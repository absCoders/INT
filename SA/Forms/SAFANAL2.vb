Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.ColorModel
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives
Imports Infragistics.UltraChart.Shared.Styles
Imports Infragistics.Win

Imports System.Drawing
Imports System.Math
Imports System.IO

Imports System.Collections
Imports System.Xml.Serialization

Public Class SAFANAL2

    ' WHY DOES IT TAKE SO LONG TO CHECK A BOX - i see now preparing summary
    ' DON'T NEED TO PERFORM ALL SQLS IF ONLY ASKING FOR SAY RTL QTY SOLD - NO NEED FOR W/S QUERIES

    ' certain views like category and reason not working
    ' sort by option - which column is sorting when we click the header?


    Dim COLUMNS_sort() As String
    Dim SATANALX As String

    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer
    Dim tblASTDSQLA As DataTable
    Dim QCOLS As New Dictionary(Of String, String)
    Dim LVL As Int16
    Dim FC_COLUMNS() As String

    Dim tblSATSLSW1 As DataTable
     
    Dim SCOPE() As String

    Dim DATA_DESCs() As String

    Dim sqlSOTINVH0() As String
    Dim sqlSUM As String
    Dim sqlICTIRECX() As String

    Private WithEvents UltraTree_DropHightLight_DrawFilter As New UltraTree_DropHightLight_DrawFilter_Class()
    Dim US_STATES() As String
    Dim USMap As MapLayer
    Dim SORTBY_COLUMN As UltraWinGrid.UltraGridColumn

 

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Form_Load_ABS()
        Catch ex As Exception
            error_has_occured = ex
        End Try
    End Sub

    Sub Form_Load_ABS()

        'Try

        If MENU_ITEM_OBJECT = "SAFANAL1" Then
        ElseIf MENU_ITEM_OBJECT = "SAFANAL2" Then
            QCOLS.Add("UNITS", "NVL(ICTIREC2.REC_UNITS,0)")
            QCOLS.Add("SALES", "NVL(ICTIREC2.REC_UNITS,0) * NVL(GMTSTYL1.PURCHASE_COST,0)")
        End If

        With dst

            ASCMAIN1.sql = ""
            For I As Int16 = 1 To 9
                ASCMAIN1.sql &= ",CTL_NO_TEXT CODE" & CStr(I)
            Next
            ASCMAIN1.sql &= ", 0 DATA_TYPE"
            For Y As Integer = 0 To 7 ' w0 thru w5, lm, tm
                ASCMAIN1.sql &= ",0.01 D" & Format(Y, "0") & Format(Y, "00")
            Next

            ASCMAIN1.sql = "Select " & Mid(ASCMAIN1.sql, 2) & " FROM TATCTLN1 where ROWNUM < 1"
            SATANALX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select * from " & SATANALX
            Create_TDA(.Tables.Add, "SATANALX", "**", 0, False)

            With .Tables.Add("SATANALP")
                For I As Int16 = 1 To 9
                    .Columns.Add("CODE" & CStr(I))
                Next
                For I As Int16 = 1 To 9
                    .Columns.Add("DESC" & CStr(I))
                Next
            End With


            Dim sqlpfx As String = ""
            For I As Int16 = 1 To 9
                sqlpfx &= "," & "CODE" & CStr(I)
            Next


            Dim sqlDATA_TYPE As String = ""
            'Dim sqlSum As String = ""
            For Y As Integer = 0 To 7
                Dim T As String = ""
                COLUMN_NAME = "D_" & Format(Y, "0")
                sqlDATA_TYPE &= ", 0.01 " & COLUMN_NAME ' & vbCrLf
                T &= "+" & COLUMN_NAME
                COLUMN_NAME = "D_" & Format(Y, "0")
                sqlDATA_TYPE &= ", 0.01 " & COLUMN_NAME & vbCrLf
                COLUMN_NAME = "D_" & Format(Y, "0")
                sqlDATA_TYPE &= ", 0.01 " & COLUMN_NAME & vbCrLf
            Next

            ASCMAIN1.sql = ""

            For DATA_TYPE As Integer = 1 To QCOLS.Count ' Each QCOL As String In QCOLS.Keys
                ASCMAIN1.sql &= Replace(sqlDATA_TYPE, "D", QCOLS.Keys(DATA_TYPE - 1))
                'ASCMAIN1.sql &= " union Select " & Mid(sqlpfx, 2) & sqlDATA_TYPE & " from " & SATANALX & " where DATA_TYPE = " & CStr(DATA_TYPE)
            Next

            ASCMAIN1.sql = "Select " & Mid(sqlpfx, 2) & ASCMAIN1.sql & " from " & SATANALX
            Create_TDA(.Tables.Add, "SATANAL1", "**", 0, False)
            .Tables("SATANAL1").Columns.Add("DESC_VALUE")
            '.Tables("SATANAL1").Columns("AVG_COST").Expression = "IIF(ISNULL(ONH_UNITS,0)=0,0,ISNULL(ONH_UNITS_X_COST,0)/ISNULL(ONH_UNITS,0))"

            For R As Integer = 0 To 11 ' REFERS TO HOW MANY DATA POINTS WE SUPPORT
                COLUMN_NAME = "COL_" & Format(R, "00")
                .Tables("SATANAL1").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
            Next

            ASCMAIN1.sql = "Select * from SATANALA"
            Create_TDA(.Tables.Add, "SATANALA", "**", 0)

            With .Tables.Add("SATANALC")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_NAME_CODE")
                .Columns.Add("COLUMN_NAME_DESC")
                .Columns.Add("TABLE_NAME_LOOKUP")
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With

            ASCMAIN1.sql = "Select CTL_NO_TEXT COLUMN_NAME, CTL_NO_TEXT CODE_VALUE, CTL_NO_TEXT DESC_VALUE from TATCTLN1"
            Create_TDA(.Tables.Add, "SATANALD", "**", 0, False, "", 2)
            .Tables("SATANALD").Columns("DESC_VALUE").MaxLength = -1 ' 100

            With .Tables.Add("SATANALO")
                .Columns.Add("DATA_CODE")
                .Columns.Add("DATA_DESC")
                .Columns.Add("SEL")
                .Columns("SEL").ReadOnly = False
                .Columns.Add("DATA_CAPTION")
                .PrimaryKey = New DataColumn() {.Columns("DATA_CODE")}
            End With

            With .Tables.Add("SATANALR")
                .Columns.Add("DATA_CODE1")
                .Columns.Add("DATA_CODE2")
                .Columns.Add("SEL1")
                .Columns.Add("SEL2")
                .Columns.Add("SEL")
                .Columns("SEL").ReadOnly = False
                .Columns.Add("SEQ", GetType(System.Int32))
                .Columns("SEQ").ReadOnly = False
                .Columns.Add("DATA_CAPTION")
                .Columns("DATA_CAPTION").ReadOnly = False
                .Columns.Add("ROW_NO", GetType(System.Int32))
                .Columns("ROW_NO").ReadOnly = False

                .PrimaryKey = New DataColumn() {.Columns("DATA_CODE1"), .Columns("DATA_CODE2")}
            End With


            Dim sql As String = ""
            For Each QCOL As String In QCOLS.Keys
                For Y As Integer = 0 To 7
                    Dim COLUMN_NAME As String = QCOL & "_" & Format(Y, "0")
                    sql &= ", 0.01 " & COLUMN_NAME
                Next
            Next

            ASCMAIN1.sql = "Select SATANALX.CODE1 STATE_CODE" _
            & sql & " from " & SATANALX & " SATANALX" _
            & " group by SATANALX.CODE1"
            ASCMAIN1.sql = "Select SATANALX.CODE1 STATE_CODE" _
            & sql & " from " & SATANALX & " SATANALX"

            Create_TDA(.Tables.Add, "SATANALS", "**", 0, False)
            For Each QCOL As String In QCOLS.Keys
                For Y As Integer = 0 To 5
                    Dim T As String = ""
                    COLUMN_NAME = QCOL & "_" & Format(Y, "0")
                    T &= "+" & COLUMN_NAME
                    COLUMN_NAME = QCOL & "_" & Format(Y, "0") & Format(6, "00")
                    .Tables("SATANALS").Columns.Add(COLUMN_NAME, GetType(System.Decimal), Mid(T, 2))
                    COLUMN_NAME = QCOL & "_" & Format(Y, "0") & Format(7, "00")
                    .Tables("SATANALS").Columns.Add(COLUMN_NAME, GetType(System.Decimal), Mid(T, 2))
                Next
            Next
            .Tables("SATANALS").Columns.Add("DATA_TYPE", GetType(System.Decimal))

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)
            With .Tables("TATSTATE")
                .Columns.Add("AMT", GetType(System.Int32))
                .Columns.Add("MAP_INDEX", GetType(System.Int32))
            End With

        End With

        Fill_Records("TATSTATE")
        Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").NewRow
        rowTATSTATE.Item("STATE_CODE") = "??"
        rowTATSTATE.Item("STATE_NAME") = "Unknown"
        dst.Tables("TATSTATE").Rows.Add(rowTATSTATE)

        dst.Tables("SATANALC").Rows.Clear()
        With dst.Tables("SATANALC").Rows
            For Each rowASTDSQLA As DataRow In Absc1.tblASTDSQLA.Rows
                Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
                Dim COLUMN_NAME2 As String = rowASTDSQLA.Item("COLUMN_NAME")
                ASCMAIN1.sql = ASCMAIN1.TACMAIN1.Get_Code_SQL_X("SAFANAL1", COLUMN_NAME, COLUMN_NAME)
                If ASCMAIN1.sql = "" Then
                    Dim COLUMN_NAME_DESC = Replace(COLUMN_NAME2, "_CODE", "_DESC")
                    .Add(New String() {COLUMN_NAME2, COLUMN_NAME2, COLUMN_NAME_DESC, COLUMN_NAME_DESC})
                Else
                    Dim DT As DataTable = ASCDATA1.GetDataTable
                    .Add(New String() {COLUMN_NAME2, COLUMN_NAME2, DT.Columns(1).ColumnName, DT.TableName})

                    For Each row As DataRow In DT.Rows
                        Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find(New String() {COLUMN_NAME2, row.Item(0)})
                        If rowSATANALD Is Nothing Then
                            rowSATANALD = dst.Tables("SATANALD").NewRow
                            rowSATANALD.Item("COLUMN_NAME") = COLUMN_NAME2
                            rowSATANALD.Item("CODE_VALUE") = row.Item(0)
                            rowSATANALD.Item("DESC_VALUE") = row.Item(1)
                            dst.Tables("SATANALD").Rows.Add(rowSATANALD)
                        End If
                    Next
                End If
            Next
        End With

        ''dst.EnforceConstraints = False
        'For Each rowSATANALC As DataRow In dst.Tables("SATANALC").Rows
        '    Dim COLUMN_NAME As String = rowSATANALC.Item("COLUMN_NAME")
        '    Dim COLUMN_NAME_CODE As String = rowSATANALC.Item("COLUMN_NAME_CODE")
        '    Dim COLUMN_NAME_DESC As String = rowSATANALC.Item("COLUMN_NAME_DESC")
        '    Dim TABLE_NAME_LOOKUP As String = rowSATANALC.Item("TABLE_NAME_LOOKUP") & ""
        '    'unrem the next line to see this form's error get caught
        '    'If COLUMN_NAME = "ITEM_CODE" Then TABLE_NAME_LOOKUP = "ICTITEM1"
        '    Fill_Records("SATANALD", "", False, "Select '" & COLUMN_NAME & "' COLUMN_NAME, " & COLUMN_NAME_CODE & " CODE_VALUE, " & COLUMN_NAME_DESC & " DESC_VALUE from " & TABLE_NAME_LOOKUP)
        'Next
        ''Fill_Records("SATANALD", "", False, "Select '" & "CUST_GROUP_CODE" & "' COLUMN_NAME, " & "CUST_CODE" & " CODE_VALUE, " & "CUST_NAME" & " DESC_VALUE from " & "ARTCUST1")

        'Me.Setup_grdSetup(grdSetup1)
        grdSATANALA.DataSource = dst.Tables("SATANALA")
        grdSATANAL1.DataSource = dst.Tables("SATANAL1")
        grdSATANALO.DataSource = dst.Tables("SATANALO")
        grdSATANALR.DataSource = dst.Tables("SATANALR")
        grdSATANALP.DataSource = dst.Tables("SATANALP")

        grdSOTINVH0.DataSource = dst.Tables("SOTINVH2")

        Dim dvw As DataView = DirectCast(grdSATANALR.DataSource, DataTable).DefaultView
        dvw.RowFilter = "SEL1 = '1' AND SEL2 = '1'"

        Format_grdSATANAL1()

        ASCMAIN1.sql = "" ' need to clear this out because of bug in ASCMAIN1.Add_Value_List - which uses sql, yet does not accept sql in parameter list - it should offer an optional parameter for sql
        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.Add_Value_List(cbeLayout, "", , New String() {":", "0:X-Month/Week", "1:Comparative"})
        Else
            ASCMAIN1.Add_Value_List(cbeLayout, "", , New String() {":", "0:X-Month/Week"})
        End If
        cbeLayout.Value = "0"

        tab1.Tabs("Inquiry").Visible = False
        tab1.Tabs("Pivot").Visible = False

        With tvwSEQ
            .Appearances.Add("DropHighLightAppearance")
            With .Appearances("DropHighLightAppearance")
                .BackColor = System.Drawing.Color.Cyan
            End With

            .Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag

            .Override.CellClickAction = UltraWinTree.CellClickAction.Default
            .ViewStyle = UltraWinTree.ViewStyle.Standard
            .AllowDrop = True
            .Override.AllowCut = DefaultableBoolean.False
            .Override.ActiveNodeAppearance.BackColor = Drawing.Color.Green ' Blue
            .Override.ActiveNodeAppearance.ForeColor = Drawing.Color.White

        End With

        'Create_Summary(grdSOTINVH0, "SO_ORDER_NO", "Count")
        'Create_Summary(grdSOTINVH0, "QTY_CASES")
        'Create_Summary(grdSOTINVH0, "QTY_UNITS")
        'Create_Summary(grdSOTINVH0, "SLSGP")
        'Create_Summary(grdSOTINVH0, "SLSNP")
        'Create_Summary(grdSOTINVH0, "STD_COST_EXT")
        'Create_Summary(grdSOTINVH0, "ADJ_COST_EXT")
        'Create_Summary(grdSOTINVH0, "EXT_REBATE")


        Check_Inquiry_Mode()

        grdSATANAL1.DisplayLayout.Override.ActiveRowAppearance.BackColor = Color.LightGreen


        'Dim dvw As DataView = dst.Tables("TATSTATE").DefaultView
        dvw = dst.Tables("TATSTATE").DefaultView
        dvw.RowFilter = "AMT <> 0"
        grdTATSTATE.DataSource = dvw

        Create_Summary(grdTATSTATE, "STATE_CODE", "Count")
        Create_Summary(grdTATSTATE, "AMT")

        With chtTotals
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            .EnableCrossHair = True
            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With

        Setup_Map()

        'Catch ex As Exception
        '    Throw ex
        '    'MsgBox("GOTCHA")
        'End Try

      

        grpOptionsSAFANAL1.Visible = (MENU_ITEM_OBJECT = "SAFANAL1")
        grpOptionsSAFANAL1.Visible = False
    End Sub

    Sub Check_Inquiry_Mode()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Generate"

                tblASTDSQLA = Absc1.grdSetupDataSource
                COLUMNS_sort = Absc1.grdSetupCOLUMNs
                If COLUMNS_sort Is Nothing Then
                    EMsg &= vbCr & "You Must Select Columns to Sort By"
                End If

                Dim SELs_Data As Int16 = Val(dst.Tables("SATANALO").Compute("COUNT (SEL)", "SEL='1'") & "")

                If SELs_Data = 0 Then
                    EMsg &= vbCr & "You cannot Generate without selecting at least 1 Data Type" '  and 1 Year"
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

            Case "Generate"

                EntryMode = "N"
                Load_Record()
                'Mode_Settings(True)

            Case "Clear"

                Clear_Settings()

            Case "Done"

                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Clear").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabGC.SelectedTab = tabGC.Tabs("Grids")
        tabGC.Tabs("Charts").Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        'Absx1.txtFor("OPS_YYYYPP").Text = ""
        'Absx1.txtFor("GROUP_NO").Text = ""
        'Absx1.txtFor("BATCH_NO").Text = ""

        EnforceConstraints(False)
        'dst.Tables("BATINVH0").Rows.Clear()
        EnforceConstraints(True)

        Setup_tab1()
        Setup_tabDetails()
        Setup_View()
        Setup_Layout_Option()
        Load_Layout_Options()
         
        SORTBY_COLUMN = Nothing
        'SORTBY_COLUMN_OLD = Nothing
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        If EntryMode = "N" Then
        Else
        End If

        EnforceConstraints(True)

        Setup_DQ()

        Generate_Inquiry(True)
        tab1.SelectedTab = tab1.Tabs("Inquiry")

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

        'BeginTrans()
        'CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "GROUP_NO"
                'sql_where = "GROUP_NO in (Select Distinct GROUP_NO from ICTPCAT1)"
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

                'Case "Load"
                '    Dim BATCH_NO As String = Split(key, ":")(0)
                '    Absx1.txtFor("BATCH_NO").Text = BATCH_NO
                '    Click_Command(command)
        End Select

        Return return_key
    End Function
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdBATGRPM5, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdTATSTATE, "CC", "Best", "Worst")
        Load_Popup_Menu(tvwDQ, "S", "Show Codes")
         Load_Popup_Menu(grdSATANAL1, "SSSS", "Show Filter", "Show GroupBox", "Show Details", "Show All Levels")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If e.SourceControl.Name = "tvwDQ" Then
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
 

        If tlb_pop.Tools.Exists("Show Details") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Details"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "N"
            tlb_sbt.Checked = Not splInquiry.Panel2Collapsed
            tlb_sbt.Tag = ""
        End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSATANAL1"
            End Select
        End If

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If

        Select Case e.Tool.Key
            Case "Show GroupBox"
                If grd IsNot Nothing AndAlso grd.Name = "grdSATANAL1" Then
                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                    grd.DisplayLayout.Bands(0).ColHeadersVisible = tlb_sbt.Checked
                End If

            Case "Show Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Codes(tlb_sbt.Checked)
                Exit Sub

            Case "Show All Levels"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).ColHeadersVisible = tlb_sbt.Checked
                'grdIMTSTATW.Visible = Not tlb_sbt.Checked
                '  UltraExplorerBar1.Groups("View").Visible = Not tlb_sbt.Checked

                Click_Node(tvwDQ.ActiveNode)
                Exit Sub

            Case "Show Details"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "N" Then
                    splInquiry.Panel2Collapsed = Not splInquiry.Panel2Collapsed
                    Setup_tab1()
                    If Not splInquiry.Panel2Collapsed Then
                        Setup_Details()
                        'splInquiry.Panel2Collapsed = False
                    Else
                        'UltraExplorerBar1.Groups("Maintain Lots").Visible = False
                        'UltraExplorerBar1.Groups("Forecast Maintenance").Visible = False
                    End If
                End If


            Case "CardView"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)

                grdSATANAL1.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked


            Case "Best"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorEnd

            Case "Worst"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorBegin

          
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

        End Select
    End Sub

    Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "BackColor"
            '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
            '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
            '    grd.ActiveRow.Cells("POS_RBG_BACKCOLOR").Value = tlb_cpt.SelectedColor.ToArgb
            '    grd.UpdateData()
            '    'Application.DoEvents()
            '    grdIMTSTATW.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
            '    Update_Record_TDA("IMTPOSS1")

            'Case "ForeColor"
            '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
            '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
            '    grd.ActiveRow.Cells("POS_RBG_FORECOLOR").Value = tlb_cpt.SelectedColor.ToArgb
            '    grd.UpdateData()
            '    'Application.DoEvents()
            '    grdIMTSTATW.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
            '    Update_Record_TDA("IMTPOSS1")

            'Case "Best"
            '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
            '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
            '    Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
            '    UltraChart1.DataBind()
            '    'grdSATCSLSS.DataBind()
            '    Application.DoEvents()
            '    grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            'Case "Worst"
            '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
            '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
            '    Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
            '    UltraChart1.DataBind()
            '    'grdSATCSLSS.DataBind()
            '    Application.DoEvents()
            '    grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub
#End Region


#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "GROUP_NO"
            Case "BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If InquiryMode Then Click_Command("Load") Else Click_Command("Edit")
                End If
        End Select
    End Sub

    Public Overrides Sub cbe_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.cbe_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)


        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "PROM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "GROUP_NO"
            Case "SUBLOC_NO"
                'If InquiryMode Then Click_Command("Load") Else Click_Command("Edit")
        End Select
    End Sub

#End Region


    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        If Absx1.txtFor("SET_DESC").Text = "" Then
            MsgBox("You must enter a Description in order to Save", _
                   MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim SET_ID As String = grpSaveAs.Tag

        If SET_ID <> "" Then
            Select Case MsgBox("Favorite " & SET_ID & " already exists." & vbCrLf & vbCrLf _
                               & "Do you want to Over-Write the Existing Favorite?" & vbCrLf & vbCrLf _
                               & "Yes = Over-write existing Favorite" & vbCrLf _
                               & "No = Create a New Favorite" & vbCrLf _
                               & "Cancel = Cancel Save Request", MsgBoxStyle.YesNoCancel, _
                               "Favorite already exists")
                Case MsgBoxResult.Cancel
                    Exit Sub
                Case MsgBoxResult.No
                    SET_ID = ASCMAIN1.Next_Control_No("SATANALA.SET_ID")
            End Select

        Else
            SET_ID = ASCMAIN1.Next_Control_No("SATANALA.SET_ID")
        End If

        Dim t As DataTable = Absc1.grdSetupDataSource
        Dim SET_ABSC As String = ""
        For Each row As DataRow In t.Select("ISNULL(SEQUENCE,0) <> 0 OR CODE_VALUES IS NOT NULL OR ISNULL(EXCLUDE,'0') = '1'")
            SET_ABSC &= row.Item("COLUMN_NAME") _
                      & vbTab & row.Item("CODE_VALUES") _
                      & vbTab & row.Item("EXCLUDE") _
                      & vbTab & row.Item("SEQUENCE") & vbCrLf
        Next

        Dim SET_DATA As String = ""
        For Each row As DataRow In dst.Tables("SATANALO").Rows
            SET_DATA &= row.Item("DATA_CODE") & vbTab _
                      & row.Item("DATA_DESC") & vbTab _
                      & row.Item("SEL") & vbTab _
                      & row.Item("DATA_CAPTION") & vbCrLf
        Next

        Dim SET_OPTIONS As String = ""

        For Each ctl As Control In grpOptionsSAFANAL1.Controls
            Dim COLUMN_NAME As String = Absx1.GetABSColumnName(ctl)
            If COLUMN_NAME <> "" Then
                SET_OPTIONS &= COLUMN_NAME & vbTab
                If TypeOf (ctl) Is ABSCS.ABSCheckBox Then
                    Dim ctl2 As ABSCS.ABSCheckBox = DirectCast(ctl, ABSCS.ABSCheckBox)
                    SET_OPTIONS &= ctl2.ABSChecked
                ElseIf TypeOf (ctl) Is UltraWinEditors.UltraOptionSet Then
                    SET_OPTIONS &= Absx1.optFor(COLUMN_NAME).Value
                End If
                SET_OPTIONS &= vbCrLf
            End If
        Next

        Dim rowSATANALA As DataRow = dst.Tables("SATANALA").Rows.Find(New Object() {MENU_ITEM_OBJECT, SET_ID})
        If rowSATANALA Is Nothing Then
            rowSATANALA = dst.Tables("SATANALA").NewRow()
            With rowSATANALA
                .Item("FORM_NAME") = MENU_ITEM_OBJECT
                .Item("SET_ID") = SET_ID
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            End With
            dst.Tables("SATANALA").Rows.Add(rowSATANALA)
        Else
            Dim row As DataRow = LookUp("SATANALA", New String() {MENU_ITEM_OBJECT, SET_ID})
            If row Is Nothing Then
                MsgBox("Favorite No Longer Exists in Database - Try Saving as New", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit Sub
            End If
        End If

        With rowSATANALA
            .Item("SET_DESC") = Absx1.txtFor("SET_DESC").Text
            .Item("SET_PUBLIC") = IIf(Absx1.chkFor("SET_PUBLIC").Checked, "1", "0")
            .Item("SET_LAYOUT") = Absx1.cbeFor("LAYOUT").Value
            .Item("SET_MW") = Absx1.optFor("MW").Value
            .Item("SET_ABSC") = SET_ABSC
            .Item("SET_DATA") = SET_DATA
            .Item("SET_OPTIONS") = SET_OPTIONS
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        End With
        Update_Record_TDA("SATANALA")

        Load_Layout_Options()
        Load_Favorite(SET_ID)

    End Sub

    Sub Load_Layout_Options()
        If IsLoading Then Exit Sub
        If SELECTION_NO = 0 Then Exit Sub

        Dim Sql As String = "Select * from SATANALA " _
        & " where (SET_PUBLIC = '1' or INIT_OPER = '" & ASCMAIN1.USER_ID & "')" _
        & " and SET_LAYOUT = '" & cbeLayout.Value & "'" _
        & " and FORM_NAME = '" & MENU_ITEM_OBJECT & "'"

        Fill_Records("SATANALA", , , Sql)
        Sort_grdColumns(grdSATANALA, "SET_DESC")

        Select Case MENU_ITEM_OBJECT
            Case "SAFANAL1"
                Select Case cbeLayout.Value
                    Case "0"
                     
                End Select

            Case "SAFANAL2"
                dst.Tables("SATANALO").Rows.Clear()
                With dst.Tables("SATANALO").Rows

                    .Add(New String() {"SLS_QTY", "Units", "0", ""})
                    .Add(New String() {"SLS_RTL", "Sales", "1", ""})
                End With

                grdSATANALO.DisplayLayout.Bands(0).SortedColumns.Clear()
                grdSATANALO.DisplayLayout.Bands(0).SortedColumns.Add("DATA_CAPTION", False)
        End Select

        Load_Layout_Columns_Selected()
        'Setup_Layout_Option()
    End Sub

    Sub Load_Layout_Columns_Selected()
        dst.Tables("SATANALR").Rows.Clear()
        Dim SEQ As Int32 = 0
        For Each row1 As DataRow In dst.Tables("SATANALO").Rows
            Dim DATA_CODE1 As String = row1.Item("DATA_CODE")
            Dim DATA_CAPTION1 As String = row1.Item("DATA_CAPTION")
            If DATA_CAPTION1 = "" Then DATA_CAPTION1 = row1.Item("DATA_DESC")
            For Each YEAR As String In New String() {"1", "2", "3", "4", "5"}
                Dim YEAR_CAPTION As String
                If YEAR = "1" Then
                    YEAR_CAPTION = "TY"
                ElseIf YEAR = "2" Then
                    YEAR_CAPTION = "LY"
                Else
                    YEAR_CAPTION = Format(Val(YEAR) - 1, "0") & "ago"
                End If
                For Each YEAR_DATA_TYPE As String In New String() {"AMT", "VARAMT", "VARPCT"}
                    Dim YEAR_DATA_TYPE_CAPTION As String = ""

                    Dim DATA_CODE2 As String = YEAR & YEAR_DATA_TYPE
                    Dim DATA_CAPTION2 As String = YEAR_CAPTION & YEAR_DATA_TYPE_CAPTION
                    Dim rowSATANALR As DataRow = dst.Tables("SATANALR").NewRow
                    rowSATANALR.Item("DATA_CODE1") = DATA_CODE1
                    rowSATANALR.Item("DATA_CODE2") = DATA_CODE2
                    rowSATANALR.Item("DATA_CAPTION") = DATA_CAPTION1 & " " & DATA_CAPTION2
                    rowSATANALR.Item("SEL") = "1"
                    rowSATANALR.Item("SEL1") = row1.Item("SEL")
                    rowSATANALR.Item("SEL2") = "0"
                    SEQ += 1
                    rowSATANALR.Item("SEQ") = SEQ

                    dst.Tables("SATANALR").Rows.Add(rowSATANALR)
                Next
            Next
        Next

        'If optLayoutOptions.Visible Then
        '    optLayoutOptions.Value = "0"
        'End If
        SEQ = 0
        For Each rowSATANALR As DataRow In dst.Tables("SATANALR").Select("", "DATA_CODE1,DATA_CODE2")
            If rowSATANALR.Item("SEL1") & "" = "1" And rowSATANALR.Item("SEL2") & "" = "1" Then
                SEQ += 1
                rowSATANALR.Item("SEQ") = SEQ
            Else
            End If
        Next

        Sort_grdColumns(grdSATANALR, "SEQ")
    End Sub

    Private Sub cbeLayout_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeLayout.ValueChanged
        Load_Layout_Options()
    End Sub

    Sub Setup_Layout_Option()

        Select Case cbeLayout.Value
            Case "0" ' X-Month/Week
               
                grdSATANAL1.DisplayLayout.Bands(0).Summaries.Clear()

                For i As Integer = 1 To 9
                    Create_Summary(grdSATANAL1, "CODE" & CStr(i), "Count")
                Next

                ReDim DATA_DESCs(11)
                With grdSATANAL1.DisplayLayout.Bands(0)
                    Dim R As Integer = 0

                    For Each rowSATANALR As DataRow In dst.Tables("SATANALR") _
                        .Select("SEL1 = '1' and SEL2 = '1' and SEL = '1'", "SEQ")
                        Dim DATA_CODE1 As String = rowSATANALR.Item("DATA_CODE1")
                        Dim DATA_CAPTION As String = rowSATANALR.Item("DATA_CAPTION")

                        If .LevelCount < R + 1 Then
                            .LevelCount = R + 1
                        End If
                        rowSATANALR.Item("ROW_NO") = R

                        COLUMN_NAME = "DATA_DESC_" & Format(R, "00")
                        .Columns(COLUMN_NAME).Hidden = False
                        .Columns(COLUMN_NAME).Level = R

                        If R >= 1 Then
                            COLUMN_NAME = "FILLER1_" & Format(R, "00")
                            .Columns(COLUMN_NAME).Hidden = False
                            .Columns(COLUMN_NAME).Level = R
                            COLUMN_NAME = "FILLER2_" & Format(R, "00")
                            .Columns(COLUMN_NAME).Hidden = False
                            .Columns(COLUMN_NAME).Level = R
                        End If

                        Dim DATA_CODE2 As String = rowSATANALR.Item("DATA_CODE2")
                        Dim YEAR As String = Mid(DATA_CODE2, 1, 1)
                        Dim YEAR_PRIOR As String = Format(Val(YEAR) + 1, "0")
                        Dim YEAR_DATA_TYPE As String = Mid(DATA_CODE2, 2)
                        For M As Integer = 1 To 8
                            COLUMN_NAME = "COL_" & Format(R, "00") & Format(M, "00")
                            .Columns(COLUMN_NAME).Format = "###,##0"
                            .Columns(COLUMN_NAME).Hidden = False
                            .Columns(COLUMN_NAME).Level = R
                            .Columns(COLUMN_NAME).Tag = DATA_CODE1 & ":" & DATA_CODE2

                            Dim DATA_EXPRESSION As String = ""
                            Select Case DATA_CODE1

                                Case "SLS_QTY", "SLS_RTL"
                                    If YEAR_DATA_TYPE = "AMT" Then
                                        DATA_EXPRESSION = CELL(DATA_CODE1, YEAR, M)
                                    ElseIf YEAR_DATA_TYPE = "VARAMT" Then
                                        DATA_EXPRESSION = CELL(DATA_CODE1, YEAR, M) & " - " & CELL(DATA_CODE1, YEAR_PRIOR, M)
                                    ElseIf YEAR_DATA_TYPE = "VARPCT" Then
                                        DATA_EXPRESSION = "100*IIF(" & CELL(DATA_CODE1, YEAR_PRIOR, M) & "=0,0,(" & CELL(DATA_CODE1, YEAR, M) & " - " & CELL(DATA_CODE1, YEAR_PRIOR, M) & ") / " & CELL(DATA_CODE1, YEAR_PRIOR, M) & ")"
                                        .Columns(COLUMN_NAME).Format = "###,##0.0"
                                    End If

                                Case Else
                                    Stop
                            End Select
                            If YEAR = "5" And (YEAR_DATA_TYPE = "VARAMT" Or YEAR_DATA_TYPE = "VARPCT") Then
                                DATA_EXPRESSION = ""
                            End If

                            dst.Tables("SATANAL1").Columns(COLUMN_NAME).Expression = DATA_EXPRESSION
                            Create_Summary(grdSATANAL1, COLUMN_NAME)
                        Next
                        DATA_DESCs(R) = DATA_CAPTION

                        Create_Summary(grdSATANAL1, "DATA_DESC_" & Format(R, "00"), "Max")

                        R += 1 : If R = 12 Then Exit For
                    Next

                    If R > 0 And R <= 11 Then
                        For R_non As Integer = 11 To R Step -1
                            COLUMN_NAME = "DATA_DESC_" & Format(R_non, "00")
                            .Columns(COLUMN_NAME).Hidden = True
                            .Columns(COLUMN_NAME).Level = 0

                            COLUMN_NAME = "FILLER1_" & Format(R_non, "00")
                            .Columns(COLUMN_NAME).Hidden = True
                            .Columns(COLUMN_NAME).Level = 0

                            COLUMN_NAME = "FILLER2_" & Format(R_non, "00")
                            .Columns(COLUMN_NAME).Hidden = True
                            .Columns(COLUMN_NAME).Level = 0

                            For M As Integer = 1 To 7
                                COLUMN_NAME = "COL_" & Format(R_non, "00") & Format(M, "00")
                                .Columns(COLUMN_NAME).Hidden = True
                                .Columns(COLUMN_NAME).Level = 0
                            Next
                        Next
                    End If

                    If R <> 0 Then
                        .LevelCount = R
                    End If

                End With


            Case "1" ' Comparative

                'Absx1.cmbFor("RYP1").Visible = (optLayoutOptions.Value = "2")
                'Absx1.cmbFor("RYW1").Visible = (optLayoutOptions.Value = "2")


        End Select

        'CreateMap()
        'CreateGraph_Totals()
        'CreateGraph_Trend()

        grdSATANAL1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Function CELL(ByVal DATA_CODE As String, ByVal YEAR As String, ByVal MONTH As Integer)
        Return DATA_CODE & "_" & Format(Val(YEAR) - 1, "0") & Format(MONTH, "00")
    End Function

    Private Sub grdSATANALA_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATANALA.AfterRowsDeleted
        Update_Record_TDA("SATANALA")
    End Sub

    Private Sub grdSATANALA_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSATANALA.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID Then
                MsgBox("You cannot delete Favorites which you did not Originate", MsgBoxStyle.OkOnly, "Cannot Perform Reqeusted Action")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Private Sub grdSATANALA_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSATANALA.ClickCellButton
        Load_Favorite(e.Cell.Row.Cells("SET_ID").Value & "")
    End Sub

    Sub Setup_DQ()

        ' COLUMNS_sort contains the columns and the sequence of those columns from ABSC1
        ' we don't use COLUMNS_sort for anything else after we read its contents in this procedure
        ' COLUMN_NAMEs are the Columns from COLUMNS_sort, in index order
        ' COLUMN_CAPTIONs are the captions for the Columns from COLUMNS_sort, in index order
        ' COLUMN_NAMEs and COLUMN_CAPTIONs represent the original order and content from COLUMNS_sort
        '  and represent the context for CODE1,CODE2,.. in the temp table

        COLUMN_NAMEs.Clear()
        COLUMN_CAPTIONs.Clear()
        For Each COLUMN_NAME As String In COLUMNS_sort
            COLUMN_NAMEs.Add(COLUMN_NAME)
            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
            COLUMN_CAPTIONs.Add(rowASTDSQLA.Item("COLUMN_CAPTION"))
        Next


        tvwSEQ.Nodes.Clear()
        'Dim anode_parent As UltraWinTree.UltraTreeNode = Nothing

        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            Dim anode As New UltraWinTree.UltraTreeNode
            anode.Text = COLUMN_CAPTIONs(i)
            anode.Key = COLUMN_NAMEs(i)
            tvwSEQ.Nodes.Add(anode)
            'If i = 0 Then
            '    tvwSEQ.Nodes.Add(anode)
            'Else
            '    anode_parent.Nodes.Add(anode)
            'End If
            'anode_parent = anode
        Next

        tvwSEQ.ExpandAll()

    End Sub

    Sub Generate_Inquiry(ByVal refresh_data As Boolean)

        grdSATANAL1.Tag = ""
        If refresh_data Then
            Dim TABLE_NAME_YP As String = "GMTCGMA0"

            ASCDATA1.ExecuteSQL("Truncate Table " & SATANALX)

            Dim sqlDATA_TYPE() As String
            ReDim sqlDATA_TYPE(QCOLS.Count)
            Dim DATA_TYPE As Integer = 0

            For Each QCOL As String In QCOLS.Keys
                DATA_TYPE += 1
                Dim sql As String = ""

                Dim XPTD As String = ""
                Dim XTOT As String = ""

                sql = "SELECT * FROM (" _
                    & "SELECT VEND_CODE, SUM (SLS_UNITS_TWK) W0" _
                    & ", SUM (SLS_UNITS_LWK) W1" _
                    & ", SUM (SLS_UNITS_2WK) W2" _
                    & ", SUM (SLS_UNITS_3WK) W3" _
                    & ", SUM (SLS_UNITS_4WK) W4" _
                    & ", SUM (SLS_UNITS_5WK) W5" _
                    & "FROM GMTCGMA0 WHERE DGC_CODE LIKE '4%' GROUP BY VEND_CODE" _
                    & ") WHERE NVL(W0,0) <> 0 OR NVL(W1,0) <> 0 OR NVL(W2,0) <> 0 OR NVL(W3,0) <> 0 OR NVL(W4,0) <> 0 OR NVL(W5,0) <> 0"

                'dst.Tables("SATANAL1").Columns(QCOL & "_" & Format(Y, "0") & Format(XMAX_now + 2, "00")).Expression = Mid(XPTD, 2)
                'dst.Tables("SATANAL1").Columns(QCOL & "_" & Format(Y, "0") & Format(XMAX_now + 1, "00")).Expression = Mid(XTOT, 2)

                sqlDATA_TYPE(DATA_TYPE) = sql
            Next


            For M As Integer = 0 To 7
                Dim G As UltraWinGrid.UltraGridGroup = grdSATANAL1.DisplayLayout.Bands(0).Groups("W" & Format(M, "00"))
                If M = 7 Then
                    G.Header.Caption = "LM MTD"
                ElseIf M = 6 Then
                    G.Header.Caption = "TM MTD"
                Else
                    G.Header.Caption = "Wk#" & Format(M, "00") ' ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(YX_MAX, M - XMAX_now), True)
                End If
            Next

            'sqlSUM = sql

            Dim X As String = Absc1.sql_GROUP_BY_cols

            'COLUMN_NAMEs_tree

            If MENU_ITEM_OBJECT = "SAFANAL1" Then
                For Each rowASTDSQLC As DataRow In Absc1.tblASTDSQLC.Select("COLUMN_NAME = 'CUST_BOUGHT_FOR'")
                    If chkDefaultBOUGHT_FOR.Checked Then
                        rowASTDSQLC.Item("COLUMN_EXPRESSION") = "NVL(SOTINVH0.CUST_BOUGHT_FOR,SOTINVH0.CUST_CODE)"
                    Else
                        rowASTDSQLC.Item("COLUMN_EXPRESSION") = ""
                    End If
                Next
                For Each rowASTDSQLD As DataRow In Absc1.tblASTDSQLD.Select("TABLE_NAME = 'ARTCUSTB'")
                    If chkDefaultBOUGHT_FOR.Checked Then
                        rowASTDSQLD.Item("COLUMN_NAME_JOIN") = "NVL(SOTINVH0.CUST_BOUGHT_FOR,SOTINVH0.CUST_CODE)"
                    Else
                        rowASTDSQLD.Item("COLUMN_NAME_JOIN") = "CUST_BOUGHT_FOR"
                    End If
                Next


                For Each rowASTDSQLC As DataRow In Absc1.tblASTDSQLC.Select("COLUMN_NAME = 'CUST_GROUP_CODE'")
                    If chkDefaultCUST_GROUP.Checked Then
                        rowASTDSQLC.Item("COLUMN_EXPRESSION") = "NVL(ARTCUST1.CUST_GROUP_CODE,SOTINVH0.CUST_CODE)"
                    Else
                        rowASTDSQLC.Item("COLUMN_EXPRESSION") = ""
                    End If
                Next

                Absc1.Get_SQL("*")
            End If

            Dim DATA_SOURCE As String = ""
            Absc1.Get_SQL("*")
            Dim TABLE_NAME_DATA_SOURCE As String = Absc1.sql_TABLE_NAME


            For DATA_TYPE = 1 To QCOLS.Count
                TABLE_NAME_YP = "GMTCGMA0"

                Dim TABLE_NAMEs As String = Absc1.sql_TABLE_NAMEs

                Dim sqlw As String = ""
                sqlw = "" _
                    & " from " & TABLE_NAME_DATA_SOURCE & TABLE_NAMEs _
                    & ASCMAIN1.SQL_Add_WHERE(Absc1.sql_JOIN & Absc1.sql_WHERE)

                ASCDATA1.ExecuteSQL("Insert into " & SATANALX & vbCrLf _
                    & " Select " & Absc1.sql_SELECT_cols & vbCrLf _
                    & ", " & CStr(DATA_TYPE) & vbCrLf _
                    & sqlDATA_TYPE(DATA_TYPE) _
                    & " group by " & Absc1.sql_GROUP_BY_cols)
            Next
        Else
            tblSATSLSW1 = Nothing
        End If


        ' Show Grid
        tab1.Tabs("Inquiry").Visible = True
        If txtSET_DESC.Text = "" Then
            tab1.Tabs("Inquiry").Text = "Custom Inquiry"
        Else
            tab1.Tabs("Inquiry").Text = txtSET_DESC.Text
        End If

        tab1.Tabs("Pivot").Visible = True

        Application.DoEvents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Collections into Selection Tree")

        ReDim COLUMN_NAME_by_Lvl(COLUMN_NAMEs.Count)
        ReDim COLUMN_CAPTION_by_Lvl(COLUMN_NAMEs.Count)
        ReDim G_by_Lvl(COLUMN_NAMEs.Count)
        ReDim SCOPE(COLUMN_NAMEs.Count)
        For G As Integer = 1 To COLUMN_NAMEs.Count
            Dim tnode As UltraWinTree.UltraTreeNode = tvwSEQ.GetNodeByKey(COLUMN_NAMEs(G - 1))
            Dim Lvl As Integer = tnode.Index + 1 ' tnode.Level + 1
            COLUMN_NAME_by_Lvl(Lvl) = COLUMN_NAMEs(G - 1)
            COLUMN_CAPTION_by_Lvl(Lvl) = tnode.Text
            G_by_Lvl(Lvl) = G
        Next

        With tvwDQ
            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
            rootColumnSet.Columns.Clear()
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count
                Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add(COLUMN_NAME_by_Lvl(Lvl))
            Next
        End With

        Dim COLUMN_NAMEs_ordered As String = ""
        Dim CODE_COLUMNs_ordered As String = ""
        For Lvl As Integer = 1 To COLUMN_NAMEs.Count
            COLUMN_NAMEs_ordered &= "," & COLUMN_NAME_by_Lvl(Lvl)
            CODE_COLUMNs_ordered &= ",CODE" & CStr(G_by_Lvl(Lvl))
        Next
        COLUMN_NAMEs_ordered = Mid(COLUMN_NAMEs_ordered, 2)
        CODE_COLUMNs_ordered = Mid(CODE_COLUMNs_ordered, 2)

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
        Dim CODE_VALUE_at_Lvl() As String = Nothing
        ReDim CODE_VALUE_at_Lvl(COLUMN_NAMEs.Count)

        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        tvwDQ.Nodes.Clear()

        Dim cur_Node_at_Lvl() As Infragistics.Win.UltraWinTree.UltraTreeNode
        ReDim cur_Node_at_Lvl(COLUMN_NAMEs.Count)
        If COLUMN_CAPTION_by_Lvl.Length = 1 Then
            aNode = tvwDQ.Nodes.Add("*", "All")
        Else
            'aNode = tvwDQ.Nodes.Add("*", "All (" & COLUMN_CAPTION_by_Lvl(1) & ")")
            aNode = tvwDQ.Nodes.Add("*", "All " & COLUMN_CAPTION_by_Lvl(1) & "s")
        End If

        cur_Node_at_Lvl(0) = aNode

        ASCMAIN1.sql = "Select Distinct " & CODE_COLUMNs_ordered & " from " & SATANALX
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        Dim last_level_set As Integer = 0

        Dim show_codes As Boolean = False
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Codes"), UltraWinToolbars.StateButtonTool)
        show_codes = tlb_sbt.Checked

        Dim images As New Dictionary(Of String, System.Drawing.Bitmap)
        images.Add("LEAF", ASCMAIN1.Get_Image(IMAGE_FOLDER, "ITEM_green"))
        images.Add("M", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M"))
        images.Add("M_OPEN", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN"))

        For Each row As DataRow In TBL.Select("", CODE_COLUMNs_ordered)
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count - 1
                If CODE_VALUE_at_Lvl(Lvl) <> row.Item(Lvl - 1) & "" Or last_level_set < Lvl Then
                    last_level_set = Lvl
                    If Lvl = 1 Then
                        aNode = tvwDQ.Nodes.Add
                    Else
                        aNode = cur_Node_at_Lvl(Lvl - 1).Nodes.Add
                    End If
                    cur_Node_at_Lvl(Lvl) = aNode

                    If Lvl = COLUMN_NAMEs.Count Then
                        ' Dim KEY As String = row.Item("ITEM_CATGY_CODE") & "/" & row.Item("COLLECTION_CODE") & "/" & row.Item("ITEM_CLASS_CODE")
                        ' aNode.Key = KEY
                        ' IF WE EVER EXPAND UPON WHAT COLUMNS TO PLACE INTO THE KEY, WE NEED TO ALSO LOOK AT TXTFINDITEM_CODE
                    End If

                    Dim CAPTION As String = "?"
                    Dim COLUMN_NAME_CODE As String = COLUMN_NAME_by_Lvl(Lvl) ' Gs(Lvl - 1)
                    Dim rowSATANALC As DataRow = dst.Tables("SATANALC").Rows.Find(COLUMN_NAME_CODE)
                    If rowSATANALC Is Nothing Then
                        CAPTION = "?"
                    Else
                        Dim COLUMN_NAME_DESC As String = rowSATANALC.Item("COLUMN_NAME_DESC")
                        Dim TABLE_NAME_LOOKUP As String = rowSATANALC.Item("TABLE_NAME_LOOKUP")
                        CAPTION = LookUp(TABLE_NAME_LOOKUP, row.Item(Lvl - 1) & "", True).Item(COLUMN_NAME_DESC) & ""
                        If CAPTION = "" Then
                            CAPTION = "?"
                        End If
                    End If

                    If show_codes Then
                        aNode.Text = row.Item(Lvl - 1) & ":" & CAPTION
                    Else
                        aNode.Text = CAPTION
                    End If

                    aNode.Tag = row.Item(Lvl - 1) & ":" & CAPTION
                    aNode.Expanded = False

                    CODE_VALUE_at_Lvl(Lvl) = row.Item(Lvl - 1) & ""
                    If last_level_set = COLUMN_NAMEs.Count - 1 Then
                        aNode.LeftImages.Add(images("LEAF"))
                    Else
                        aNode.Override.NodeAppearance.Image = images("M")
                        aNode.Override.ExpandedNodeAppearance.Image = images("M_OPEN")
                    End If

                    For iLvl As Integer = 1 To Lvl
                        aNode.Cells(iLvl - 1).Value = CODE_VALUE_at_Lvl(iLvl)
                    Next
                End If
            Next
        Next

        Dim rows() As DataRow = dst.Tables("SATANALR").Select("SEL1 = '1' and SEL2 = '1' and SEL = '1'", "SEQ")
        grdSATANALR.Tag = rows(0)("DATA_CODE1") & ":" & rows(0)("DATA_CODE2")


        Setup_View()
        Setup_tabDetails()

        If tvwDQ.Nodes.Count > 0 Then
            tvwDQ.ActiveNode = tvwDQ.Nodes(0)
            tvwDQ.Nodes(0).Selected = True
            Click_Node(tvwDQ.Nodes(0))
            SortGrid("CODES", False)
            grdSATANALR.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
            Setup_Layout_Option()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub tab1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab1.SelectedTabChanged
        Setup_tab1()
    End Sub

    Sub Setup_tab1()
        If SELECTION_NO = 0 Then Exit Sub
        If tab1.SelectedTab Is Nothing Then Exit Sub

        tabGC.Tabs("Charts").Visible = tab1.SelectedTab.Key = "Inquiry" And Not splInquiry.Panel2Collapsed ' False ' (tab1.SelectedTab.Key = "Inquiry")
        ' UltraExplorerBar1.Groups("Options").Visible = (tab1.SelectedTab.Key = "Setup")
        UltraExplorerBar1.Groups("Layout Options").Visible = False ' (tab1.SelectedTab.Key = "Setup")
        UltraExplorerBar1.Groups("Period Options").Visible = (tab1.SelectedTab.Key = "Setup")
        UltraExplorerBar1.Groups("Data Lines").Visible = (tab1.SelectedTab.Key = "Inquiry")
        UltraExplorerBar1.Groups("Data Options").Visible = Not (tab1.SelectedTab.Key = "Pivot")

        UltraExplorerBar1.Groups("Pivot Options").Visible = (tab1.SelectedTab.Key = "Pivot")
        If (tab1.SelectedTab.Key = "Pivot") Then
            grdSATANALP.Visible = False

            optPivotValue.Items.Clear()
            For Each rowSATANALR As DataRow In dst.Tables("SATANALR").Select("SEL1 = '1' AND SEL2 = '1' AND SEL = '1'", "SEQ")
                Dim DATA_CODE1 As String = rowSATANALR.Item("DATA_CODE1")
                Dim DATA_CODE2 As String = rowSATANALR.Item("DATA_CODE2")
                If Mid(DATA_CODE2, 2) = "AMT" And QCOLS.ContainsKey(DATA_CODE1) Then
                    optPivotValue.Items.Add(DATA_CODE1 & "_" & Format(Val(Mid(DATA_CODE2, 1, 1)) - 1, "0"), rowSATANALR.Item("DATA_CAPTION"))
                End If
            Next
            If optPivotValue.Items.Count > 0 Then
                optPivotValue.CheckedIndex = 0
            Else
                MsgBox("No Pivotable Data Values", MsgBoxStyle.OkOnly, "Cannot Pivot")
                tab1.SelectedTab = tab1.Tabs(0)
            End If

            optPivotBy.Items.Clear()
            For I As Int16 = 1 To COLUMN_NAME_by_Lvl.Length - 1
                Dim COLUMN_NAME As String = COLUMN_NAME_by_Lvl(I)
                Dim COLUMN_CAPTION As String = COLUMN_CAPTION_by_Lvl(I)
                optPivotBy.Items.Add(COLUMN_NAME, COLUMN_CAPTION)
            Next
            If optPivotBy.Items.Count > 1 Then
                optPivotBy.CheckedIndex = 0
            Else
                MsgBox("Need 2 or more Sorted Codes to peform a Pivot", MsgBoxStyle.OkOnly, "Cannot Pivot")
                tab1.SelectedTab = tab1.Tabs(0)
            End If

        End If

    End Sub

#Region "tvwSEQ"
    Private Sub tvwSEQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.Click

        'Try
        '    Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
        '    Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
        '    Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

        '    If tnode IsNot Nothing Then
        '        Click_Node(tvwSEQ.ActiveNode)
        '        tvwSEQ.SelectedNodes.Clear()
        '        tvwSEQ.ActiveNode.Selected = True
        '    End If


        'Catch ex As Exception

        'End Try

    End Sub

    Private Sub tvwSEQ_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwSEQ.DragDrop
        If SELECTION_NO = 0 Then Exit Sub
        Dim Node As UltraWinTree.UltraTreeNode
        Dim SelectedNodes As UltraWinTree.SelectedNodesCollection
        Dim DropNode As UltraWinTree.UltraTreeNode
        Dim i As Integer

        DropNode = UltraTree_DropHightLight_DrawFilter.DropHightLightNode

        SelectedNodes = e.Data.GetData(GetType(UltraWinTree.SelectedNodesCollection))
        SelectedNodes = SelectedNodes.Clone()

        SelectedNodes.SortByPosition()

        For i = 0 To SelectedNodes.Count - 1
            Node = SelectedNodes(i)
            Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
        Next

        'Select Case UltraTree_DropHightLight_DrawFilter.DropLinePosition
        '    Case DropLinePositionEnum.OnNode
        '        For i = 0 To SelectedNodes.Count - 1
        '            Node = SelectedNodes(i)
        '            Node.Reposition(DropNode.Nodes)
        '        Next
        '    Case DropLinePositionEnum.BelowNode
        '        For i = 0 To SelectedNodes.Count - 1
        '            Node = SelectedNodes(i)
        '            Node.Reposition(DropNode, UltraWinTree.NodePosition.Next)
        '            DropNode = Node
        '        Next
        '    Case DropLinePositionEnum.AboveNode
        '        For i = 0 To SelectedNodes.Count - 1
        '            Node = SelectedNodes(i)
        '            Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
        '        Next
        'End Select

        Absc1.Clear_grdSetup(False)
        For ii As Int16 = 0 To tvwSEQ.Nodes.Count - 1
            Absc1.Re_SEQ(tvwSEQ.Nodes(ii).Key, True)
            'Debug.Print(tvwSEQ.Nodes(ii).Key)
        Next



        Generate_Inquiry(False)

        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwSEQ_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.DragLeave
        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwSEQ_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwSEQ.DragOver
        If SELECTION_NO = 0 Then Exit Sub
        Dim Node As UltraWinTree.UltraTreeNode
        Dim PointInTree As System.Drawing.Point

        With tvwSEQ
            PointInTree = .PointToClient(New System.Drawing.Point(e.X, e.Y))

            Node = .GetNodeFromPoint(PointInTree)

            If Node Is Nothing Then
                e.Effect = DragDropEffects.None
                UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                Return
            End If

            If Me.IsParentNode(Node) And Me.IsParentNodeSelected(Me.tvwSEQ) Then
                If PointInTree.Y > (Node.Bounds.Top + 2) AndAlso PointInTree.Y < (Node.Bounds.Bottom - 2) Then
                    e.Effect = DragDropEffects.None
                    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                    Return
                End If
            End If

            'If IsAnyParentSelected(Node) Then
            '    e.Effect = DragDropEffects.None
            '    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
            '    Return
            'End If

            UltraTree_DropHightLight_DrawFilter.SetDropHighlightNode(Node, PointInTree)
            e.Effect = DragDropEffects.Move
        End With
    End Sub

    Private Sub tvwSEQ_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwSEQ.MouseUp
        tvwSEQ.SelectedNodes.Clear()
        Dim anode As Infragistics.Win.UltraWinTree.UltraTreeNode = tvwSEQ.GetNodeFromPoint(e.Location)
        If anode IsNot Nothing Then
            anode.Selected = True
            tvwSEQ.ActiveNode = anode
        End If
    End Sub

    Private Sub tvwSEQ_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles tvwSEQ.QueryContinueDrag
        If e.EscapePressed Then
            e.Action = DragAction.Cancel
            UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
        End If
    End Sub

    Private Sub tvwSEQ_SelectionDragStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.SelectionDragStart
        If SELECTION_NO = 0 Then Exit Sub
        tvwSEQ.DoDragDrop(tvwSEQ.SelectedNodes, DragDropEffects.Move)
    End Sub

#End Region
#Region "tvwDQ"
    Private Sub tvwDQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwDQ.Click

        Try
            Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
            Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
            Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

            If tnode IsNot Nothing Then
                Click_Node(tvwDQ.ActiveNode)
                tvwDQ.SelectedNodes.Clear()
                tvwDQ.ActiveNode.Selected = True
            End If


        Catch ex As Exception

        End Try

    End Sub
#End Region

    Sub Click_Node(ByVal tnode As UltraWinTree.UltraTreeNode)

        Dim SELs_Data As Int16 = Val(dst.Tables("SATANALO").Compute("COUNT (SEL)", "SEL='1'") & "")

        If SELs_Data = 0 Then
            grdSATANAL1.Visible = False
            Exit Sub
        Else
            grdSATANAL1.Visible = True
        End If


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Summary")

        If tnode IsNot Nothing Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Levels"), UltraWinToolbars.StateButtonTool)
            LVL = tnode.Level + 1
            If tnode.Key = "*" Then
                LVL = 0
            End If
            Dim COLS_Select As String = ""
            Dim COLS_Group_By As String = ""
            Dim COLS_Order_By As String = ""
            Dim sqlW As String = ""
            Dim CAPTION As String = ""

            grdSATANAL1.DisplayLayout.Bands(0).Groups("DESCRIPTION").Header.Caption = COLUMN_CAPTION_by_Lvl(LVL + 1)


            For G As Integer = 1 To COLUMN_NAME_by_Lvl.Count - 1
                If G <= LVL + 1 Or tlb_sbt.Checked Then
                    COLS_Select &= ",CODE" & CStr(G_by_Lvl(G)) & " CODE" & CStr(G)
                    COLS_Group_By &= ",CODE" & CStr(G_by_Lvl(G))
                Else
                    COLS_Select &= ",NULL CODE" & CStr(G)
                End If
                If G <= LVL Then
                    If tnode.Cells(G - 1).Text = "" Then
                        sqlW &= " and CODE" & CStr(G_by_Lvl(G)) & " is Null"
                    Else
                        sqlW &= " and CODE" & CStr(G_by_Lvl(G)) & " = '" & tnode.Cells(G - 1).Text & "'"
                    End If

                    SCOPE(G) = tnode.Cells(G - 1).Text

                    Dim CODE_VALUE As String = tnode.Cells(G - 1).Text
                    Dim DESC_VALUE As String = Get_Description(COLUMN_NAME_by_Lvl(G), CODE_VALUE)
                    CAPTION &= ", " & COLUMN_CAPTION_by_Lvl(G) & " " & CODE_VALUE & ":" & DESC_VALUE
                End If
                COLS_Order_By &= ",CODE" & CStr(G)
                With grdSATANAL1.DisplayLayout.Bands(0).Columns("CODE" & CStr(G))
                    .Hidden = (G < LVL + 1) Or (G > LVL + 1 And Not tlb_sbt.Checked)
                    .Header.Caption = COLUMN_CAPTION_by_Lvl(G)
                End With
            Next

            If LVL = 0 Then
                grdSATANAL1.Text = "All" ' tvwDQ.Nodes(0).Text
            Else
                grdSATANAL1.Text = Mid(CAPTION, 3)
            End If

            If LVL = 0 Then
                optLevelCharts.ValueList.ValueListItems(0).DisplayText = tvwDQ.Nodes(0).Text
            Else
                optLevelCharts.ValueList.ValueListItems(0).DisplayText = COLUMN_CAPTION_by_Lvl(LVL) & " " & SCOPE(LVL)
            End If
            optLevelCharts.ValueList.ValueListItems(1).DisplayText = "Individual" & vbCrLf & COLUMN_CAPTION_by_Lvl(LVL + 1)

            If COLUMN_NAME_by_Lvl.Count - 1 < 9 Then
                For G As Integer = COLUMN_NAME_by_Lvl.Count To 9
                    With grdSATANAL1.DisplayLayout.Bands(0).Columns("CODE" & CStr(G))
                        .Hidden = True
                    End With
                    COLS_Select &= ",NULL CODE" & CStr(G)
                Next
            End If

            dst.Tables("SATANAL1").Rows.Clear()

            Dim DATA_TYPE As Integer = 0
            For Each QCOL As String In QCOLS.Keys
                DATA_TYPE += 1
                Dim sqlQCOLs As String = ""

                COLUMN_NAME = QCOL
                '  sqlQCOLs &= ", SUM (" & Mid(T14, 2) & ") " & COLUMN_NAME & vbCrLf

                ASCMAIN1.sql = "Select " & Mid(COLS_Select, 2) & vbCrLf _
                & sqlQCOLs _
                       & " from " & SATANALX & ASCMAIN1.SQL_Add_WHERE(sqlW & " AND DATA_TYPE = " & CStr(DATA_TYPE)) & vbCrLf _
                & " group by " & Mid(COLS_Group_By, 2)
                'Fill_Records("SATANAL1", , , ASCMAIN1.sql)
                For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim sqlwc As String = ""
                    For G As Integer = 1 To COLUMN_NAME_by_Lvl.Count - 1
                        If G <= LVL + 1 Or tlb_sbt.Checked Then
                            If row.Item(G - 1) & "" = "" Then
                                sqlwc &= " and CODE" & CStr(G_by_Lvl(G)) & " IS NULL"
                            Else
                                sqlwc &= " and CODE" & CStr(G_by_Lvl(G)) & " = '" & row.Item(G - 1) & "'"
                            End If
                        End If
                    Next
                    sqlwc = Mid(sqlwc, 5)
                    Dim row2() As DataRow = dst.Tables("SATANAL1").Select(sqlwc)
                    Dim rowX As DataRow
                    If row2.Length = 0 Then
                        rowX = dst.Tables("SATANAL1").NewRow
                        dst.Tables("SATANAL1").Rows.Add(rowX)
                    Else
                        rowX = row2(0)
                    End If
                    For Each dcol As DataColumn In row.Table.Columns
                        If dcol.ColumnName.EndsWith(Format(6, "00")) Or dcol.ColumnName.EndsWith(Format(7, "00")) Then
                        Else
                            rowX.Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
                        End If
                    Next
                Next
            Next

            Sort_grdColumns(grdSATANAL1, "CODE" & CStr(LVL + 1))
            For Each rowSATANAL1 As DataRow In dst.Tables("SATANAL1").Rows
                Dim CODE_VALUE As String = rowSATANAL1.Item("CODE" & CStr(LVL + 1)) & ""
                Dim DESC_VALUE As String = Get_Description(COLUMN_NAME_by_Lvl(LVL + 1), CODE_VALUE)
                rowSATANAL1.Item("DESC_VALUE") = DESC_VALUE
            Next

        End If

        If tabGC.SelectedTab.Key = "Map" Or 1 = 1 Then
            CreateMap()
        End If
        CreateGraph_Totals()
        CreateGraph_Trend()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Show_Codes(ByVal tf As Boolean)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Toggling Codes")
        Show_Codes_for_Nodes(tf, tvwDQ.Nodes)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Show_Codes_for_Nodes(ByVal tf As Boolean, ByVal anodes As UltraWinTree.TreeNodesCollection)
        For Each cnode As UltraWinTree.UltraTreeNode In anodes
            If cnode.Key <> "*" Then
                Dim CAPTION As String = cnode.Tag & ""
                If tf Then
                    cnode.Text = CAPTION
                Else
                    cnode.Text = Split(CAPTION, ":")(1)
                End If

                If cnode.HasNodes Then
                    Show_Codes_for_Nodes(tf, cnode.Nodes)
                End If
            End If
        Next
    End Sub

    Private Sub grdSATANAL1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATANAL1.AfterRowActivate

        If Not splInquiry.Panel2Collapsed Then
            Show_SOTINVH0()

            If optLevelCharts.Value = "2" Then
                CreateGraph_Totals()
                CreateGraph_Trend()
                CreateMap()
            End If

        End If

    End Sub

    Sub Show_SOTINVH0()
        If grdSATANAL1.ActiveRow.IsDataRow Then
            If 1 = 1 Then Exit Sub
            If Not tabMain.Tabs("Sales Details").Selected Or MENU_ITEM_OBJECT <> "SAFANAL1" Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Fetching Sales History")

            ASCMAIN1.sql = "Select SOTINVH0.* " & sqlSOTINVH0(0)
            Dim SPEC As String = ""
            For i As Integer = 1 To LVL + 1
                Dim DATA_VALUE As String = SCOPE(i)
                If i = LVL + 1 Then
                    DATA_VALUE = grdSATANAL1.ActiveRow.Cells("CODE" & CStr(i)).Value & ""
                End If

                Dim g As Integer = G_by_Lvl(i)

                If DATA_VALUE = "" Then
                    ASCMAIN1.sql &= " and " & sqlSOTINVH0(g) & " IS NULL"
                Else
                    ASCMAIN1.sql &= " and " & sqlSOTINVH0(g) & " = '" & DATA_VALUE & "'"
                End If
                'SPEC &= "; " & COLUMN_CAPTIONs(i - 1) & " " & DATA_VALUE
                SPEC &= "; " & COLUMN_CAPTIONs(g - 1) & " " & DATA_VALUE
            Next
            Fill_Records("SOTINVH0", "", , ASCMAIN1.sql)
            Sort_grdColumns(grdSOTINVH0, "SO_ORDER_NO")
            grdSOTINVH0.Text = Mid(SPEC, 3)
            grdSOTINVH0.Visible = True

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        Else
            grdSOTINVH0.Visible = False
        End If
    End Sub

    Private Sub grdSATANAL1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATANAL1.Click
        '    Stop


    End Sub

    Private Sub grdSATANAL1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSATANAL1.DoubleClickCell
        Setup_Details()
        splInquiry.Panel2Collapsed = False
        Setup_tab1()

        Select Case e.Cell.Column.Key

            Case Else
                'tabDetails.SelectedTab = tabDetails.Tabs("Lot Details")

        End Select

    End Sub

    Sub Setup_Details()
        If grdSATANAL1.ActiveRow Is Nothing Then
            grdSOTINVH0.Visible = False
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Details")
        Show_SOTINVH0()
        Setup_tabDetails()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optView_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Setup_View()
    End Sub

    Sub Setup_View()
        If SELECTION_NO = 0 Then Exit Sub

        splInquiry.Panel2Collapsed = True
    End Sub

    Sub Format_grdSATANAL1()

        Dim G As UltraWinGrid.UltraGridGroup
        With grdSATANAL1.DisplayLayout.Bands(0)
            .LevelCount = 12

            G = .Groups.Add("CODES")
            G.Header.Fixed = True

            G.Header.Caption = "Code"
            For I As Integer = 1 To 9
                .Columns("CODE" & CStr(I)).Group = G
                G.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                G.CellAppearance.BackColor = Drawing.Color.LightYellow
            Next

            For R As Integer = 1 To 11
                COLUMN_NAME = "FILLER1_" & Format(R, "00")
                .Columns.Add(COLUMN_NAME)
                .Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).Level = R
            Next

            G = .Groups.Add("DESCRIPTION")
            G.Header.Fixed = True
            G.Header.Caption = "Description"
            .Columns("DESC_VALUE").Group = G
            G.Header.Appearance.BackColor2 = Drawing.Color.Yellow
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
            G.CellAppearance.BackColor = Drawing.Color.LightYellow

            For R As Integer = 1 To 11
                COLUMN_NAME = "FILLER2_" & Format(R, "00")
                .Columns.Add(COLUMN_NAME)
                .Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).Level = R
            Next



            G = .Groups.Add("DATA_DESC")
            G.Header.Fixed = True
            G.Header.Caption = "Data"
            G.Header.Appearance.BackColor2 = Drawing.Color.DarkGray
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
            G.CellAppearance.BackColor = Drawing.Color.LightGray

            For R As Integer = 0 To 11
                COLUMN_NAME = "DATA_DESC_" & Format(R, "00")
                .Columns.Add(COLUMN_NAME)
                .Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).Level = R
            Next
            G.Width = 80

            For M As Integer = 1 To 8
                G = .Groups.Add("M" & Format(M, "00"))

                ' Dim BTN As New Misc.UltraButton


                For R As Integer = 0 To 7
                    COLUMN_NAME = "COL_" & Format(R, "00") & Format(M, "00")
                    .Columns(COLUMN_NAME).Group = G
                    .Columns(COLUMN_NAME).Level = R
                    .Columns(COLUMN_NAME).Width = 80
                    .Columns(COLUMN_NAME).Header.Appearance.TextHAlign = HAlign.Center
                    .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Right
                    If M = 6 Or M = 7 Then
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightSkyBlue ' .LightSteelBlue
                    End If

                Next
                G.Header.Appearance.TextHAlign = HAlign.Center
                G.Header.Appearance.BackColor2 = Drawing.Color.Orange
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                G.Width = 80
            Next

            .ColHeadersVisible = False
        End With
    End Sub

    Sub Get_SQL_for_COLUMN_NAME(ByVal sender As Object, ByVal e As ABSC.ABSC.grdSetupClickCellEventArgs) _
    Handles Absc1.Get_SQL_for_COLUMN_NAME

        Dim PROD_CODEs As String = ""
        'If chkUseSmartLists.Checked Then
        '    PROD_CODEs = Absc1.SQLA("PROD_CODE", , True)
        '    'Dim row As DataRow = Absc1.tblASTDSQLA.Rows.Find("PROD_CODE")
        '    'If row IsNot Nothing Then
        '    '    Dim PROD_CODES_selected As String = row.Item("CODE_VALUES") & ""
        '    '    If PROD_CODES_selected <> "" Then
        '    '        PROD_CODEs = "'" & Replace(PROD_CODES_selected, ",", "','") & "'"
        '    '    End If
        '    'End If
        'End If

        Select Case e.COLUMN_NAME
            ' NOTE COOL IS NOT IN SMART LISTS
            ' NOTE CATGY IS NOT IN SMART LISTS
            Case "COOL_COMPLIANT", "CATEGORY_CODE"
            Case Else

                Dim COLUMN_NAME_in_ICTLOTD2 As String = e.COLUMN_NAME
                Dim COLUMN_NAME_in_LOOKUP As String = e.COLUMN_NAME
                If e.COLUMN_NAME = "DIVISION_CODE_O" Then
                    COLUMN_NAME_in_ICTLOTD2 = "PROD_DIV_CODE"
                    COLUMN_NAME_in_LOOKUP = "DIVISION_CODE"
                End If
                If e.COLUMN_NAME = "DIVISION_CODE_G" Then
                    COLUMN_NAME_in_ICTLOTD2 = "DIVISION_CODE"
                    COLUMN_NAME_in_LOOKUP = "DIVISION_CODE"
                End If
                'If chkUseSmartLists.Checked And PROD_CODEs <> "" Then
                '    e.SQL = "Select * from (" & e.SQL & ") where " & COLUMN_NAME_in_LOOKUP & " in " _
                '    & " (Select Distinct " & COLUMN_NAME_in_ICTLOTD2 & " from ICTLOTD2 where PROD_CODE in (" & PROD_CODEs & "))"

                'End If
        End Select
    End Sub

    Private Sub grdSATANALA_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATANALA.DoubleClickRow
        Load_Favorite(e.Row.Cells("SET_ID").Value & "")
    End Sub

    Private Sub grdSATANALA_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATANALA.InitializeLayout

    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub


    Private Function IsParentNode(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        'Dim Tag As String
        'Tag = Node.Tag
        'Return Split(Tag, Chr(1))(1) = "M"
    End Function

    Private Function IsParentNodeSelected(ByVal Tree As UltraWinTree.UltraTree) As Boolean
        'For Each SelectedNode As UltraWinTree.UltraTreeNode In Tree.SelectedNodes
        '    If Me.IsParentNode(SelectedNode) Then Return True
        'Next
        'Return False
    End Function

    Sub Load_Favorite(ByVal SET_ID As String)

        Dim rowSATANALA As DataRow = dst.Tables("SATANALA").Rows.Find(New Object() {MENU_ITEM_OBJECT, SET_ID})

        Absx1.txtFor("SET_DESC").Text = rowSATANALA.Item("SET_DESC") & ""
        Absx1.chkFor("SET_PUBLIC").Checked = rowSATANALA.Item("SET_PUBLIC") & "" = "1"

        lblSET_ID.Text = rowSATANALA.Item("SET_ID") & ""
        'grpSaveAs.Text = rowSATANALA.Item("SET_ID") & ""
        grpSaveAs.Tag = SET_ID

        Dim t As DataTable = Absc1.grdSetupDataSource
        For Each row As DataRow In t.Rows
            row.Item("CODE_VALUES") = ""
            row.Item("EXCLUDE") = "0"
            row.Item("SEQUENCE") = DBNull.Value
        Next
        For Each rd As String In Split(rowSATANALA.Item("SET_ABSC") & "", vbCrLf)
            Dim f() As String = Split(rd, vbTab)
            If f(0) <> "" Then
                Dim r As DataRow = t.Rows.Find(f(0))
                If r IsNot Nothing Then
                    r.Item("CODE_VALUES") = f(1)
                    r.Item("EXCLUDE") = f(2)
                    If f(3) <> "" Then
                        r.Item("SEQUENCE") = f(3)
                    End If
                End If
            End If
        Next

        For Each row As DataRow In dst.Tables("SATANALO").Rows
            row.Item("SEL") = "0"
        Next
        For Each row As DataRow In dst.Tables("SATANALR").Rows
            row.Item("SEL") = "1"
        Next
        For Each rd As String In Split(rowSATANALA.Item("SET_DATA") & "", vbCrLf)
            Dim f() As String = Split(rd, vbTab)
            If f.Length > 1 Then
                Dim r As DataRow = dst.Tables("SATANALO").Rows.Find(f(0))
                If r IsNot Nothing Then
                    r("SEL") = f(2)
                End If
            End If
        Next

        Load_Layout_Columns_Selected()
    End Sub

    Sub Clear_Settings()
        txtSET_DESC.Text = ""
        'grpSaveAs.Text = ""
        grpSaveAs.Tag = ""
        Absc1.Clear_grdSetup(True)
        tab1.Tabs("Inquiry").Text = "Inquiry"
        tab1.Tabs("Inquiry").Visible = False
        tab1.SelectedTab = tab1.Tabs("Setup")

        tab1.Tabs("Pivot").Visible = False
    End Sub

    Sub Print_Report()

        'Call Print_Report_Begin()
        'CR_params.Add("PF25", Format(YWF(25, 1)))
        'Generate_Report("DPRPLAN1", "Demand Planning", grdIMTSTATW.Text)
        'Call Print_Report_End()

    End Sub

    Function Get_Description( _
    ByVal COLUMN_NAME As String, _
    ByVal CODE_VALUE As String, _
    Optional ByVal use_code_as_default_value As Boolean = False)
        Dim DESC_VALUE As String = IIf(use_code_as_default_value, CODE_VALUE, "")
        Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find _
                       (New String() {COLUMN_NAME, CODE_VALUE})
        If rowSATANALD IsNot Nothing Then
            DESC_VALUE = rowSATANALD.Item("DESC_VALUE") & ""
        Else
            DESC_VALUE = "?"
        End If

        Return DESC_VALUE
    End Function

    Private Sub optLayoutOptions_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Layout_Option()
    End Sub

    Private Sub grdSATANALO_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSATANALO.AfterRowUpdate
        Dim DATA_CODE As String = e.Row.Cells("DATA_CODE").Value
        Dim sql As String = ""
        Dim COLUMN_NAME As String = ""
        sql = "DATA_CODE1 = '" & DATA_CODE & "'"
        COLUMN_NAME = "SEL1"
        For Each rowSATANALR As DataRow In dst.Tables("SATANALR").Select(sql)
            rowSATANALR.Item(COLUMN_NAME) = e.Row.Cells("SEL").Value
        Next

        Dim SEQ As Integer = 0
        For Each rowSATANALR As DataRow In dst.Tables("SATANALR").Select("", "DATA_CODE1,DATA_CODE2")
            If rowSATANALR.Item("SEL1") & "" = "1" And rowSATANALR.Item("SEL2") & "" = "1" Then
                SEQ += 1
                rowSATANALR.Item("SEQ") = SEQ
            Else
            End If
        Next

        Sort_grdColumns(grdSATANALR, "SEQ")
        If ScreenMode Then
            Setup_Layout_Option()
            Click_Node(tvwDQ.ActiveNode)
        End If

    End Sub

    Private Sub grdSATANAL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATANAL1.InitializeRow
        For L As Integer = 0 To grdSATANAL1.DisplayLayout.Bands(0).LevelCount - 1
            e.Row.Cells("DATA_DESC_" & Format(L, "00")).Value = DATA_DESCs(L)
        Next

        'e.Row.Cells(SORTBY_COLUMN_OLD.Key).Appearance.BackColor = Color.Empty
        'e.Row.Cells(SORTBY_COLUMN.Key).Appearance.BackColor = Color.Yellow
    End Sub

    Private Sub grdSATANALR_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSATANALR.AfterRowUpdate
        Setup_Layout_Option()
        Click_Node(tvwDQ.ActiveNode)
        'If grdSATANAL1.Tag <> "" Then
        '    SortGrid(grdSATANAL1.Tag, True)
        'End If
        SortGrid("CODES", True)
        grdSATANALR.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub


#Region "Charts"

    Sub CreateMap()
        If 1 = 1 Then Exit Sub
        UltraChart1.Visible = False
        grdTATSTATE.Visible = False
        If grdSATANAL1.Rows.Count = 0 Then Exit Sub
        If grdSATANALR.ActiveRow Is Nothing Then Exit Sub

        If Not tabMain.Tabs("Map").Selected Or MENU_ITEM_OBJECT <> "SAFANAL1" Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Data by State")

        ASCMAIN1.sql = "Select SOTINVH1.CUST_SHIP_TO_STATE STATE_CODE" _
        & sqlSUM _
        & sqlSOTINVH0(0)

        Dim SPEC As String = ""
        For i As Integer = 1 To LVL + 1
            Dim DATA_VALUE As String = SCOPE(i)
            If i = LVL + 1 Then
                DATA_VALUE = grdSATANAL1.ActiveRow.Cells("CODE" & CStr(i)).Value & ""
                If optLevelCharts.Value = "1" Then Exit For
            End If
            If DATA_VALUE = "" Then
                ASCMAIN1.sql &= " and " & sqlSOTINVH0(i) & " IS NULL"
            Else
                ASCMAIN1.sql &= " and " & sqlSOTINVH0(i) & " = '" & DATA_VALUE & "'"
            End If
            SPEC &= "; " & COLUMN_CAPTIONs(i - 1) & " " & DATA_VALUE
        Next
        ASCMAIN1.sql &= " group by SOTINVH1.CUST_SHIP_TO_STATE"

        Fill_Records("SATANALS", "", , ASCMAIN1.sql)

        For Each ROW As DataRow In dst.Tables("TATSTATE").Rows
            ROW.Item("AMT") = 0
        Next

        Dim DATA_TYPE As String = "COL_" & Format(Val(grdSATANALR.ActiveRow.Cells("ROW_NO").Value & ""), "00") & "13"
        dst.Tables("SATANALS").Columns("DATA_TYPE").Expression = dst.Tables("SATANAL1").Columns(DATA_TYPE).Expression

        For Each rowSATANALS As DataRow In dst.Tables("SATANALS").Rows
            Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(rowSATANALS.Item("STATE_CODE"))
            If rowTATSTATE IsNot Nothing Then
                rowTATSTATE.Item("AMT") = Val(rowTATSTATE.Item("AMT") & "") + Val(rowSATANALS.Item("DATA_TYPE") & "")
            Else
                rowTATSTATE = dst.Tables("TATSTATE").Rows.Find("??")
                rowTATSTATE.Item("AMT") = Val(rowTATSTATE.Item("AMT") & "") + Val(rowSATANALS.Item("DATA_TYPE") & "")
            End If
        Next

        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()
        UltraChart1.Refresh()
        tabMain.Tabs("Map").Visible = True

        grdTATSTATE.DisplayLayout.Bands(0).Columns("AMT").Header.Caption = grdSATANALR.ActiveRow.Cells("DATA_CAPTION").Value
        Sort_grdColumns(grdTATSTATE, "AMT".ToLower)
        'grdTATSTATE.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        Dim CAPTION As String = Replace(optLevelCharts.Text, vbCrLf, " ")
        If optLevelCharts.Value = "2" Then
            CAPTION &= " (" & grdSATANAL1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Value & ")"
        End If
        grdTATSTATE.Text = CAPTION

        UltraChart1.Visible = True
        grdTATSTATE.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub CreateGraph_Totals()

        chtTotals.Visible = False
        If grdSATANAL1.Rows.Count = 0 Then Exit Sub
        If grdSATANALR.ActiveRow Is Nothing Then Exit Sub

        chtTotals.DataSource = Nothing

        Dim DATA_TYPE As String = "COL_" & Format(Val(grdSATANALR.ActiveRow.Cells("ROW_NO").Value & ""), "00") & "13"

        Dim RL() As String

        chtTotals.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTotals.LabelHash = labelHash

        chtTotals.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTotals.Tooltips.FormatString = "<HIGHLOW>"

        Dim RLi As Integer = 0

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim CODE1 As String = ""
        Dim CHARTED_CODE As String = "CODE" & CStr(LVL + 1)

        Dim DTX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SATANAL1").Select(), CHARTED_CODE, "DESC_VALUE")
        DTX.Columns.Add(DATA_TYPE, GetType(System.Decimal))
        For Each rowDTX As DataRow In DTX.Rows
            Dim SQL3 As String = "ISNULL(" & CHARTED_CODE & ",'') = '" & rowDTX.Item(0) & "'"
            Dim VALUE As Decimal = Val(dst.Tables("SATANAL1").Compute("SUM(" & DATA_TYPE & ")", SQL3) & "")
            rowDTX.Item(DATA_TYPE) = VALUE
        Next

        Dim PCT_at_TOP_N As Decimal = 0
        Dim VALUE_TOTAL As Decimal = Val(DTX.Compute("SUM(" & DATA_TYPE & ")", "") & "")
        Dim VALUE_CHARTED As Decimal = 0

        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", DATA_TYPE & " DESC")
            'RL(RLi) = row.Item(CHARTED_CODE) & ":" & row("DESC_VALUE")
            RL(RLi) = row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item(CHARTED_CODE), row.Item(DATA_TYPE)})

            If optChartTrend.Value = "N" And RLi <= Val(numChartTrend.Value & "") Then
                PCT_at_TOP_N = 100 * Val(row.Item(DATA_TYPE & "00")) / VALUE_TOTAL
            End If
        Next

        'Dim CAPTION As String = optLevelCharts.ValueList.ValueListItems(0).DisplayText
        'Dim CAPTION As String = grdSATANAL1.Text
        chtTotals.TitleTop.Text = "All " & grdSATANAL1.DisplayLayout.Bands(0).Columns(CHARTED_CODE).Header.Caption & "s" & vbCrLf & grdSATANALR.ActiveRow.Cells("DATA_CAPTION").Value & ""
        chtTotals.Data.SetRowLabels(RL)
        'chtTotals.Data.SetColumnLabels(CL)

        chtTotals.DataSource = DTY
        chtTotals.PieChart.ColumnIndex = -1

        chtTotals.PieChart.OthersCategoryPercent = 2
        If optChartTrend.Value = "C" Then
            chtTotals.PieChart.OthersCategoryPercent = Val(numChartTrend.Value & "")
        Else
            chtTotals.PieChart.OthersCategoryPercent = PCT_at_TOP_N
        End If
        chtTotals.DataBind()

        chtTotals.Visible = True
    End Sub

    Sub CreateGraph_Trend()

        chtTrend.Visible = False
        If grdSATANAL1.Rows.Count = 0 Then Exit Sub
        If grdSATANALR.ActiveRow Is Nothing Then Exit Sub

        Dim periods As Integer = 5

        Dim DATA_TYPE As String = "COL_" & Format(Val(grdSATANALR.ActiveRow.Cells("ROW_NO").Value & ""), "00")
        Dim S As Integer = 1
        'If DATA_TYPE = "R" Then
        '    S = -1
        'End If

        chtTrend.DataSource = Nothing

        Dim RL() As String
        Dim CL() As String
        ReDim CL(periods)

        For i As Integer = 1 To periods
            CL(i - 1) = grdSATANAL1.DisplayLayout.Bands(0).Groups("M" & Format(i, "00")).Header.Caption
        Next

        chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTrend.LabelHash = labelHash


        'Dim CODE1 As String = COLUMN_NAME_by_Lvl(LVL)

        Dim CHARTED_CODE As String = "CODE" & CStr(LVL + 1)

        Dim CAPTION As String = grdSATANALR.ActiveRow.Cells("DATA_CAPTION").Value & ""
        If optLevelCharts.Value = "1" Then
            'CAPTION &= " for " & grdSATANAL1.Text ' optLevelCharts.ValueList.ValueListItems(0).DisplayText
            CAPTION &= " for All " & grdSATANAL1.DisplayLayout.Bands(0).Columns(CHARTED_CODE).Header.Caption & "s"
        Else
            'CAPTION &= " for " & grdSATANAL1.Text _
            '& " by " & Replace(optLevelCharts.ValueList.ValueListItems(1).DisplayText, vbCrLf, " ")
            CAPTION &= " by " & Replace(optLevelCharts.ValueList.ValueListItems(1).DisplayText, vbCrLf, " ")
        End If
        chtTrend.TitleTop.Text = CAPTION

        chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTrend.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To periods
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next


        Dim DTX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SATANAL1").Select, CHARTED_CODE, "DESC_VALUE")
        Dim T As String = ""
        For P As Integer = 1 To periods
            Dim COLUMN_NAME_period As String = DATA_TYPE & Format(periods - P, "0")
            DTX.Columns.Add(COLUMN_NAME_period, GetType(System.Decimal))
            T &= "+" & COLUMN_NAME_period
        Next
        DTX.Columns.Add(DATA_TYPE & Format(6, "0"), GetType(System.Decimal), Mid(T, 2))

        If optLevelCharts.Value = "1" Then
            DTX.Rows.Clear()
            Dim rowDTX As DataRow = DTX.NewRow
            rowDTX.Item(CHARTED_CODE) = ""
            DTX.Rows.Add(rowDTX)
        End If

        Dim RLi As Integer = 0

        Dim VALUE_TOTAL As Decimal = S * Val(dst.Tables("SATANAL1").Compute("SUM(" & DATA_TYPE & "13" & ")", "") & "")
        Dim VALUE_CHARTED As Decimal = 0

        Dim chart_all_others As Boolean = False

        ReDim RL(DTX.Rows.Count - 1)
        ''chtTrend.TitleTop.Text = "Trend " & optTD.Text & " " & optTrend.Text & ", by " & optRSTSLSA1.Text

        Dim rowDT As DataRow = Nothing

        For Each rowDTX As DataRow In DTX.Rows
            Dim SQL3 As String = "ISNULL(" & CHARTED_CODE & ",'') = '" & rowDTX.Item(0) & "'"
            If optLevelCharts.Value = "1" Then
                SQL3 = ""
            End If

            Dim this_record_is_others As Boolean = False

            Dim U00 As Decimal = S * Val(dst.Tables("SATANAL1").Compute("SUM(" & DATA_TYPE & "13" & ")", SQL3) & "")
            Dim CODE_VALUE As String = rowDTX.Item(0) & "" ' "CODE1"
            Dim DESC_VALUE As String = rowDTX.Item("DESC_VALUE") & ""

            If (optChartTrend.Value = "C" And VALUE_TOTAL > 0 AndAlso 100 * U00 / VALUE_TOTAL > Val(numChartTrend.Value & "")) _
            Or (optChartTrend.Value = "N" And RLi < Val(numChartTrend.Value & "")) Then
            Else
                this_record_is_others = True
                CODE_VALUE = "Z"
                DESC_VALUE = "All Others"
            End If

            If Not this_record_is_others Or chart_all_others Then
                If RLi <> 0 AndAlso RL(RLi - 1) = CODE_VALUE & ":" & DESC_VALUE Then
                Else
                    'RL(RLi) = CODE_VALUE & ":" & DESC_VALUE
                    RL(RLi) = DESC_VALUE
                    RLi += 1
                    rowDT = DT.NewRow
                    rowDT.Item("CODE_VALUE") = CODE_VALUE
                    rowDT.Item("DESC_VALUE") = DESC_VALUE
                    DT.Rows.Add(rowDT)
                End If

                VALUE_CHARTED += +Val(rowDTX.Item(DATA_TYPE & "13") & "")

                For P As Integer = 1 To periods
                    Dim COLUMN_NAME_period As String = DATA_TYPE & Format(P, "00")
                    Dim UP As Decimal = S * Val(dst.Tables("SATANAL1").Compute("SUM(" & COLUMN_NAME_period & ")", SQL3) & "")

                    rowDT.Item("P" & Format(P, "00")) = Val(rowDT.Item("P" & Format(P, "00")) & "") _
                                                      + UP
                Next
            End If

        Next


        chtTrend.Data.SetRowLabels(RL)
        chtTrend.Data.SetColumnLabels(CL)

        Dim CHART_CAPTION As String = ""
        Dim VALUE_PCT As Decimal = 0
        If VALUE_TOTAL <> 0 Then
            VALUE_PCT = VALUE_CHARTED / VALUE_TOTAL
        End If
        If optChartTrend.Value = "C" Then
            CHART_CAPTION = "Cut-off " & numChartTrend.Value & "%, Charting " & CStr(DT.Rows.Count) & " of " & CStr(DTX.Rows.Count) & ", " & Format(VALUE_PCT, "##.0%")
        Else
            CHART_CAPTION = "Top " & numChartTrend.Value & " of " & CStr(DTX.Rows.Count) & ", " & Format(VALUE_PCT, "##.0%")
        End If
        chtTrend.TitleBottom.Text = CHART_CAPTION

        chtTrend.DataSource = DT
        chtTrend.DataBind()
        chtTrend.Visible = True
    End Sub

    Private Sub chtTrend_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTrend.ChartDataClicked
        Select_CODE_VALUE_from_TATDASHX(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Private Sub cmdChartRedraw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChartRedraw.Click
        CreateGraph_Totals()
        CreateGraph_Trend()
    End Sub

    Private Sub tbkChartTrend_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbkChartTrend.Scroll
        chtTrend.Axis.Y.ScrollScale.Scale = (100 - Me.tbkChartTrend.Value) / 100.0
    End Sub

    Private Sub chtTotals_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTotals.ChartDataClicked
        'Select_CODE_VALUE_from_TATDASHX(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Private Sub optLevelCharts_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLevelCharts.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabGC.SelectedTab.Key = "Map" Or 1 = 1 Then
            CreateMap()
        End If
        'CreateGraph_Totals()
        CreateGraph_Trend()
    End Sub

    Sub Select_CODE_VALUE_from_TATDASHX(ByVal CODE_VALUE As String)
        Exit Sub
        'For Each grow As UltraWinGrid.UltraGridRow In grdSATANAL1.Rows
        '    If grow.Cells("CODE1").Value & "" = CODE_VALUE Then
        '        grdSATANAL1.ActiveRow = grow
        '        grdSATANAL1.Selected.Rows.Clear()
        '        grow.Selected = True
        '        Exit Sub
        '    End If
        'Next
    End Sub

#End Region

#Region "US States Map"
    Private Sub grdTATSTATE_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATSTATE.InitializeRow
        If USMap.COLORS.ContainsKey(e.Row.Cells("STATE_NAME").Text) Then
            e.Row.Cells("AMT").Appearance.ForeColor = USMap.COLORS(e.Row.Cells("STATE_NAME").Text)
        End If
    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        USMap = New MapLayer(points)

        US_STATES = USMap.STATES
        For i As Integer = 0 To USMap.STATES.Length - 1
            Dim rowTATSTATE() As DataRow = dst.Tables("TATSTATE").Select("STATE_NAME = '" & USMap.STATES(i) & "'")
            If rowTATSTATE.Length = 1 Then
                rowTATSTATE(0).Item("MAP_INDEX") = i
                rowTATSTATE(0).Item("STATE_NAME") = USMap.STATES(i)
            End If
            ' Add(New Object() {"", USMap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USMap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("TATSTATE")))
        Me.UltraChart1.LabelHash = labelRenderers
        Me.UltraChart1.Tooltips.FormatString = "<USMap>"

        ''set border
        Me.UltraChart1.Border.CornerRadius = 20
        Me.UltraChart1.Border.Thickness = 0
        Me.UltraChart1.BackColor = Color.White

        '' set color model
        'Me.UltraChart1.ColorModel.ColorBegin = Color.AliceBlue
        Me.UltraChart1.ColorModel.ColorBegin = Color.Red
        Me.UltraChart1.ColorModel.ColorEnd = Color.Blue '  Color.Yellow ' Color.FromArgb(24, 89, 165)
        Me.UltraChart1.ColorModel.AlphaLevel = 255
        Me.UltraChart1.ColorModel.ModelStyle = ColorModels.DataValueLinearRange

        '' legend
        Me.UltraChart1.Legend.Visible = True
        Me.UltraChart1.Axis.X.Extent = 10
        Me.UltraChart1.Legend.SpanPercentage = 10
        Me.UltraChart1.Legend.Location = LegendLocation.Right

        '' set the data
        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()
    End Sub

    Private Function StatesData() As StateDataInfo()
        Dim StatesDataFromDataSource() As StateDataInfo
        ReDim StatesDataFromDataSource(49)
        If SELECTION_NO <> 0 Then
            For I As Integer = 0 To US_STATES.Length - 1
                Debug.Print(US_STATES(I))
                'Dim rows() As DataRow = dst.Tables("TATSTATE").Select("STATE_NAME = '" & US_STATES(I) & "'")
                Dim rows() As DataRow = dst.Tables("TATSTATE").Select("MAP_INDEX = " & CStr(I))
                Dim SALES As Int32 = 0
                If rows.Length = 1 Then
                    SALES = Val(rows(0).Item("AMT") & "")
                End If
                StatesDataFromDataSource(I) = New StateDataInfo(US_STATES(I), SALES, "")
            Next
        End If
        'StatesDataFromDataSource(0) = New StateExpenseViewInfo("Alabama", 1915560.96, "")
        Return StatesDataFromDataSource
    End Function
#End Region


    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub

        tabGC.Tabs("Charts").Visible = tab1.SelectedTab.Key = "Inquiry" And Not splInquiry.Panel2Collapsed AndAlso (tabMain.SelectedTab.Key = "Charts" Or tabMain.SelectedTab.Key = "Map")
        optLevelCharts.Visible = (tab1.SelectedTab.Key = "Inquiry" And Not splInquiry.Panel2Collapsed AndAlso (tabMain.SelectedTab.Key = "Charts" Or tabMain.SelectedTab.Key = "Map"))

        If tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Visible Then
            If tabMain.SelectedTab.Key = "Map" Then
                CreateMap()
            ElseIf tabMain.SelectedTab.Key = "Charts" Then
                CreateGraph_Totals()
                CreateGraph_Trend()
            End If
        End If
    End Sub

    Private Sub grdSATANAL1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdSATANAL1.MouseDown
        Dim grid As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

        Dim element As UIElement = grid.DisplayLayout.UIElement.LastElementEntered
        ' See if the element is a HeaderUIElement. This will probably never happen 
        ' because the header element is filled with a TexUIElement, but it's best to 
        ' cover all the bases. 
        Dim headerElement As UltraWinGrid.HeaderUIElement = TryCast(element, UltraWinGrid.HeaderUIElement)
        If headerElement Is Nothing Then
            ' See if the element has a HeaderUIElement in it's parent chain. 
            Try
                headerElement = TryCast(element.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)), UltraWinGrid.HeaderUIElement)

            Catch ex As Exception

            End Try
        End If
        ' We failed to find a HeaderUIElement, so we must not be on a header. 
        If headerElement Is Nothing Then
            Return
        End If
        ' A HeaderUIElement could be the element for the grid caption, a group, or a column. 
        ' Check if this head has a Group. 
        Dim groupHeader As UltraWinGrid.GroupHeader = TryCast(headerElement.Header, UltraWinGrid.GroupHeader)
        If groupHeader Is Nothing Then
            Return
        End If

        Dim group As UltraWinGrid.UltraGridGroup = groupHeader.Group
        SortGrid(group.Key, False)

    End Sub

    Sub SortGrid(ByVal groupKey As String, ByVal no_flip As Boolean)

        With grdSATANAL1.DisplayLayout.Bands(0)
            If groupKey = "CODES" Or groupKey = "DESCRIPTION" _
            Or (Len(groupKey) = 3 And Mid(groupKey, 1, 1) = "M" And Mid(groupKey, 2, 2) >= "01" And Mid(groupKey, 2, 2) <= Format(5, "00")) Then
                'If New String() {"CODES", "DESCRIPTION", "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12"}.Contains(groupKey) Then
                Dim DESC As Boolean = False

                SORTBY_COLUMN = .Groups(groupKey).Columns(0)
                Dim SORTBY As String = grdSATANALR.Tag
                For Each gcol As UltraWinGrid.UltraGridColumn In .Groups(groupKey).Columns
                    If Not gcol.Hidden Then
                        If gcol.Tag = SORTBY Then
                            SORTBY_COLUMN = gcol
                        End If
                    End If
                Next

                Dim GROUP_KEYs As New List(Of String)
                GROUP_KEYs.Add("CODES")
                GROUP_KEYs.Add("DESCRIPTION")
                For M As Integer = 1 To 5
                    GROUP_KEYs.Add("M" & Format(M, "00"))
                Next

                For Each GROUP_KEY As String In GROUP_KEYs
                    'For Each GROUP_KEY As String In New String() {"CODES", "DESCRIPTION", "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12"}
                    If GROUP_KEY = groupKey Then
                        If grdSATANAL1.Tag = groupKey And .SortedColumns.Count <> 0 Then
                            If no_flip Then
                                If .SortedColumns(0).SortIndicator = UltraWinGrid.SortIndicator.Descending Then
                                    DESC = True
                                End If
                            Else
                                If .SortedColumns(0).SortIndicator = UltraWinGrid.SortIndicator.Ascending Then
                                    DESC = True
                                End If
                            End If
                        End If
                        If DESC Then
                            .Groups(GROUP_KEY).Header.Appearance.ForeColor = Color.Red
                        Else
                            .Groups(GROUP_KEY).Header.Appearance.ForeColor = Color.Green
                        End If

                    Else
                        .Groups(GROUP_KEY).Header.Appearance.ForeColor = Color.Empty
                    End If
                Next
                .SortedColumns.Clear()
                .SortedColumns.Add(SORTBY_COLUMN, DESC, False)
                grdSATANAL1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
                grdSATANAL1.Tag = groupKey
            End If
        End With

    End Sub

    Private Sub cmdGeneratePivot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGeneratePivot.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating Pivot")

        grdSATANALP.Visible = False
        grdSATANALP.DisplayLayout.Bands(0).SortedColumns.Clear()

        dst.Tables("SATANALP").Rows.Clear()

        Dim C As Int16 = dst.Tables("SATANALP").Columns.Count - 18
        If C > 0 Then
            For I2 As Int16 = 0 To C - 1
                Dim COLUMN_NAME As String = "VAL_" & Format(I2, "000")
                Dim summary As UltraWinGrid.SummarySettings
                For Each summary In grdSATANALP.DisplayLayout.Bands(0).Summaries
                    grdSATANALP.DisplayLayout.Bands(0).Summaries.Remove(summary)
                Next
            Next
        End If


        grdSATANALP.DataSource = Nothing

        grdSATANALP.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdSATANALP.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show

        If C > 0 Then
            For I2 As Int16 = 0 To C - 1
                dst.Tables("SATANALP").Columns.Remove("VAL_" & Format(I2, "000"))
            Next
        End If

        Dim COLUMN_NAME_pivot As String = optPivotBy.Value
        Dim COLUMN_CODE_pivot As Int16 = optPivotBy.CheckedIndex + 1

        Dim I As Int16 = 0
        Dim S As String = ""
        Dim DSQL As String = ""
        Dim P As New List(Of String)

        Dim COLUMN_FORMAT As String = grdSATANAL1.DisplayLayout.Bands(0).Columns("COL_" & Format(optPivotValue.CheckedIndex, "00") & Format(7, "00")).Format

        Dim COLUMN_NAME_DSQL_BASE As String = optPivotValue.Value
        Dim COLUMN_NAME_DSQL As String = ""
        Dim COLUMN_NAME_WHERE As String = ""
        Dim QCOL As String = Mid(COLUMN_NAME_DSQL_BASE, 1, COLUMN_NAME_DSQL_BASE.Length - 2)
       
        ASCMAIN1.sql = "Select Distinct CODE" & CStr(COLUMN_CODE_pivot) & " from " & SATANALX
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CODE" & CStr(COLUMN_CODE_pivot))
            Dim CODE_VALUE As String = row.Item(0) & ""
            I += 1
            Dim COLUMN_NAME As String = "VAL_" & Format(I, "000")
            S &= "+ISNULL(" & COLUMN_NAME & ",0)"
            dst.Tables("SATANALP").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
            P.Add(CODE_VALUE)
            DSQL &= ", SUM (DECODE(" & "CODE" & CStr(COLUMN_CODE_pivot) & "," & IIf(CODE_VALUE = "", "NULL", "'" & CODE_VALUE & "'") & "," & Mid(COLUMN_NAME_DSQL, 2) & ",0)) VAL_" & Format(I, "000") & vbCrLf
        Next
        dst.Tables("SATANALP").Columns.Add("VAL_" & Format(0, "000"), GetType(System.Decimal), Mid(S, 2))
        ' dst.Tables("SATANALP").Columns.Add("VAL_" & Format(0, "000"), GetType(System.Decimal))
        grdSATANALP.Visible = True
        grdSATANALP.DataSource = dst.Tables("SATANALP")

        Dim SSQL As String = ""
        Dim GSQL As String = ""

        Dim clr(9) As System.Drawing.Color
        clr(1) = Color.Blue
        clr(2) = Color.Green
        clr(3) = Color.Orange
        clr(4) = Color.Violet
        clr(5) = Color.Red
        clr(6) = Color.Blue
        clr(7) = Color.Green
        clr(8) = Color.Orange
        clr(9) = Color.Violet

        Dim gby As String = ""
        With grdSATANALP.DisplayLayout.Bands(0)

            .ColHeaderLines = 2
            For I3 As Int16 = 1 To 9
                Dim COLUMN_NAME As String = "CODE" & Format(I3, "0")
                Dim COLUMN_NAME_DESC As String = "DESC" & Format(I3, "0")
                .Columns(COLUMN_NAME).Hidden = (I3 > optPivotBy.Items.Count) Or I3 = COLUMN_CODE_pivot
                .Columns(COLUMN_NAME_DESC).Hidden = .Columns(COLUMN_NAME).Hidden
                .Columns(COLUMN_NAME_DESC).Header.VisiblePosition = .Columns(COLUMN_NAME).Header.VisiblePosition + 1

                '.Columns(COLUMN_NAME).HiddenWhenGroupBy = DefaultableBoolean.True
                '.Columns(COLUMN_NAME_DESC).HiddenWhenGroupBy = DefaultableBoolean.True

                If .Columns(COLUMN_NAME).Hidden Then
                    SSQL &= ", NULL " & COLUMN_NAME & vbCrLf
                Else
                    SSQL &= "," & COLUMN_NAME & vbCrLf
                    GSQL &= "," & COLUMN_NAME & vbCrLf
                    gby &= "/" & COLUMN_CAPTION_by_Lvl(I3)

                    .Columns(COLUMN_NAME).Header.Caption = COLUMN_CAPTION_by_Lvl(I3)
                    .Columns(COLUMN_NAME).Width = 60
                    .Columns(COLUMN_NAME_DESC).Header.Caption = "Description"
                    .Columns(COLUMN_NAME_DESC).Width = 130
                End If
                .Columns(COLUMN_NAME).Header.Appearance.ForeColor = clr(I3)
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Beige
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                .Columns(COLUMN_NAME).CellAppearance.ForeColor = clr(I3)

                .Columns(COLUMN_NAME_DESC).Header.Appearance.ForeColor = clr(I3)
                .Columns(COLUMN_NAME_DESC).Header.Appearance.BackColor2 = Color.Beige
                .Columns(COLUMN_NAME_DESC).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                .Columns(COLUMN_NAME_DESC).CellAppearance.BackColor = Color.Beige
                .Columns(COLUMN_NAME_DESC).CellAppearance.ForeColor = clr(I3)

            Next

            For I3 As Int16 = 1 To P.Count
                COLUMN_NAME = "VAL_" & Format(I3, "000")

                Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find _
                    (New String() {COLUMN_NAME_pivot, P(I3 - 1)})
                Dim DESC_VALUE As String = ""
                If rowSATANALD IsNot Nothing Then
                    DESC_VALUE = rowSATANALD.Item("DESC_VALUE") & ""
                Else
                    DESC_VALUE = "?"
                End If

                .Columns(COLUMN_NAME).Header.Caption = P(I3 - 1) & vbLf & DESC_VALUE
                .Columns(COLUMN_NAME).Format = COLUMN_FORMAT
                .Columns(COLUMN_NAME).Width = 100
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightBlue
                '.Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                Create_Summary(grdSATANALP, COLUMN_NAME)
            Next

            COLUMN_NAME = "VAL_000"
            .Columns(COLUMN_NAME).Header.Caption = "Total"
            .Columns(COLUMN_NAME).Format = COLUMN_FORMAT
            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LimeGreen
            '.Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.White
            .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
            .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            Create_Summary(grdSATANALP, COLUMN_NAME)
        End With

        Dim QC() As String = QCOLS.Keys.ToArray
        Dim DATA_TYPE As Integer = 0
        For DT As Integer = 0 To QC.Length - 1
            If QC(DT) = QCOL Then
                DATA_TYPE = DT + 1
            End If
        Next

        ASCMAIN1.sql = "Select " & Mid(SSQL, 2) _
        & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL" _
        & DSQL & " from " & SATANALX _
        & " where DATA_TYPE = " & CStr(DATA_TYPE) _
        & " AND (" & Mid(COLUMN_NAME_WHERE, 4) & ")" _
        & " group by " & Mid(GSQL, 2)

        For Each ROW As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSATANALP As DataRow = dst.Tables("SATANALP").NewRow
            rowSATANALP.ItemArray = ROW.ItemArray
            dst.Tables("SATANALP").Rows.Add(rowSATANALP)

            For I3 As Int16 = 1 To COLUMN_NAME_by_Lvl.Count - 1
                If I3 <> COLUMN_CODE_pivot Then
                    Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find _
                        (New String() {COLUMN_NAME_by_Lvl(I3), rowSATANALP.Item("CODE" & CStr(I3)) & ""})
                    Dim DESC_VALUE As String = ""
                    If rowSATANALD IsNot Nothing Then
                        DESC_VALUE = rowSATANALD.Item("DESC_VALUE") & ""
                    End If
                    rowSATANALP.Item("DESC" & CStr(I3)) = DESC_VALUE

                    If COLUMN_CODE_pivot <> COLUMN_NAME_by_Lvl.Count - 1 And I3 <> COLUMN_NAME_by_Lvl.Count - 1 _
                    Or COLUMN_CODE_pivot = COLUMN_NAME_by_Lvl.Count - 1 And I3 <> COLUMN_NAME_by_Lvl.Count - 2 Then
                        grdSATANALP.DisplayLayout.Bands(0).SortedColumns.Add("CODE" & CStr(I3), False, True)
                        'grdSATANALP.DisplayLayout.Bands(0).SortedColumns("CODE" & CStr(I3)).Header.Caption = "aa"
                    Else
                        grdSATANALP.DisplayLayout.Bands(0).SortedColumns.Add("CODE" & CStr(I3), False, False)
                    End If

                End If
            Next
        Next
        'grdSATANALP.DisplayLayout.Override.GroupByRowDescriptionMask = "[value]"
        grdSATANALP.Text = optPivotValue.Text & ", by " & Mid(gby, 2) & ", Pivot by " & optPivotBy.Text

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub grdSATANALP_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATANALP.AfterRowActivate

    End Sub

    Private Sub grdSATANALP_InitializeGroupByRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdSATANALP.InitializeGroupByRow
        If e.Row.IsGroupByRow Then
            Dim gbyrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(e.Row, UltraWinGrid.UltraGridGroupByRow)
            Dim COLUMN_NAME = gbyrow.Column.Key
            Dim LVL As Int16 = Val(Mid(COLUMN_NAME, 5))

            Dim DESC_VALUE As String = ""
            Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find(New String() {COLUMN_NAME_by_Lvl(LVL), gbyrow.Value & ""})
            If rowSATANALD IsNot Nothing Then
                DESC_VALUE = rowSATANALD.Item("DESC_VALUE") & ""
            End If
            gbyrow.Description = COLUMN_CAPTION_by_Lvl(LVL) & " : " & gbyrow.Value & " " & DESC_VALUE

        End If

    End Sub

    Private Sub grdSATANALP_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATANALP.InitializeRow
    End Sub

    Private Sub grdSATANALR_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSATANALR.BeforeCellUpdate
        If e.NewValue & "" <> "1" Then
            Dim SELs As Int16 = Val(dst.Tables("SATANALR").Compute("COUNT(SEL)", "SEL1= '1' and SEL2 = '1' and SEL = '1'") & "")
            If SELs <= 1 Then
                MsgBox("You Cannot De-Select the Last Line of Data", MsgBoxStyle.OkOnly, "There would be No Rows to Show, Silly")
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub grdSATANALR_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSATANALR.ClickCellButton
        If e.Cell.Row.Cells("SEL").Value & "" = "1" Then
            grdSATANALR.Tag = e.Cell.Row.Cells("DATA_CODE1").Value & ":" & e.Cell.Row.Cells("DATA_CODE2").Value
        End If
        grdSATANALR.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        If grdSATANAL1.Tag <> "" And grdSATANAL1.Tag <> "CODES" And grdSATANAL1.Tag <> "DESCRIPTION" Then
            SortGrid(grdSATANAL1.Tag, True)
        End If
    End Sub

    Private Sub grdSATANALR_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATANALR.InitializeLayout

    End Sub

    Private Sub grdSATANALR_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATANALR.InitializeRow
        If grdSATANALR.Tag = e.Row.Cells("DATA_CODE1").Value & ":" & e.Row.Cells("DATA_CODE2").Value Then
            e.Row.Cells("SORT").ButtonAppearance.Image = e.Row.Cells("SORT").Column.Header.Appearance.Image
        Else
            e.Row.Cells("SORT").ButtonAppearance.Image = Nothing
        End If
    End Sub

    Private Sub grdSATANALR_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdSATANALR.MouseUp
        If grdSATANALR.ActiveRow IsNot Nothing AndAlso grdSATANALR.ActiveRow.DataChanged _
        AndAlso grdSATANALR.ActiveCell IsNot Nothing AndAlso grdSATANALR.ActiveCell.Column.Key = "SEL" Then
            grdSATANALR.ActiveRow.Update()
        End If
    End Sub

    Private Sub SplitContainer10_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub chkVARAMT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Set_SEL2()
    End Sub

    Private Sub chkVARPCT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Set_SEL2()
    End Sub

    Private Sub cbeYears_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Set_SEL2()
    End Sub

    Sub Set_SEL2()
        Me.Cursor = Cursors.WaitCursor
        For Each rowSATANALR As DataRow In dst.Tables("SATANALR").Rows
 rowSATANALR.Item("SEL2") = "0"
        Next
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chtTrend_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles chtTrend.DoubleClick
        Dim FILENAME As String = Replace(Replace(Replace(chtTrend.TitleTop.Text, " ", "_"), vbCrLf, ""), ":", "_") & ".jpg"
        chtTrend.SaveTo(ASCMAIN1.Folders("Temp") & FILENAME, System.Drawing.Imaging.ImageFormat.Jpeg)
        Show_Document(ASCMAIN1.Folders("Temp") & FILENAME)
    End Sub

    Private Sub chtTotals_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles chtTotals.DoubleClick
        Dim FILENAME As String = Replace(Replace(Replace(chtTotals.TitleTop.Text, " ", "_"), vbCrLf, ""), ":", "_") & ".jpg"
        chtTotals.SaveTo(ASCMAIN1.Folders("Temp") & FILENAME, System.Drawing.Imaging.ImageFormat.Jpeg)
        Show_Document(ASCMAIN1.Folders("Temp") & FILENAME)
    End Sub

    Public Overrides Function Excel_Export(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        If grd.Name = "grdSATANAL1" And grd.DisplayLayout.GroupByBox.Hidden Then

            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Levels"), UltraWinToolbars.StateButtonTool)

            If tlb_sbt.Checked Then
                MyBase.Gembox_Excel_Export(grd)
            Else
                Dim n_orig As UltraWinTree.UltraTreeNode = tvwDQ.ActiveNode

                GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

                Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile

                If dst.Tables.Contains("ASTGRIDC") Then
                    dst.Tables("ASTGRIDC").Rows.Clear()
                Else
                    With dst.Tables.Add("ASTGRIDC")
                        .Columns.Add("SHEET", GetType(System.Int32))
                        .Columns.Add("ROW", GetType(System.Int32))
                        .Columns.Add("COL", GetType(System.Int32))
                        .Columns.Add("COLOR1", GetType(System.Int64))
                        .Columns.Add("COLOR2", GetType(System.Int64))
                        .Columns.Add("GRADIENT", GetType(System.Int32))
                    End With
                End If

                Dim N As UltraWinTree.TreeNodesCollection = Nothing
                If tvwDQ.ActiveNode.Index = 0 Then
                    N = tvwDQ.Nodes
                Else
                    Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSATANAL1, False, , , tab1.Tabs("Inquiry").Text)
                    N = tvwDQ.ActiveNode.Nodes
                End If

                For Each NODE As UltraWinTree.UltraTreeNode In N
                    Click_Node(NODE)
                    Application.DoEvents()
                    Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSATANAL1, False, , , tab1.Tabs("Inquiry").Text)
                Next
                Gembox_Export_to_Excel_Show(myWorkbook)

                Click_Node(n_orig)
            End If



        ElseIf grd.Name = "grdSATANALP" Then
            MyBase.Gembox_Excel_Export(grd)
        ElseIf grd.Name = "grdSetup" Then
            MyBase.Gembox_Excel_Export(grd)
        Else
            MyBase.Excel_Export(grd)
        End If
        Return Nothing
    End Function

End Class