
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

Public Class ASFLOGX1

    Dim ASTLOGX1 As String
    Dim sqlASTLOGX1 As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        sqlASTLOGX1 = "SELECT X.* , TO_CHAR(X.LOG_DATE, 'HH24MI') HHMM, " _
             & "Y.LOG_DATE LOG_DATE_END, extract( second from  X.LOG_DATE - Y.LOG_DATE) SECS" _
             & " from ASTLOGX1 X, ASTLOGX1 Y" _
             & " where X.LOG_DATE > SYSDATE - 2/24" _
             & "   and Y.LOG_ID = X.LOG_ID -1"
        ASTLOGX1 = ASCMAIN1.Temp_Table(sqlASTLOGX1)

        With dst

            ASCMAIN1.sql = "Select * from " & ASTLOGX1
            Create_TDA(.Tables.Add, "ASTLOGX1", "**", 0, False, "N")

            Dim sqlASTLOGX2 As String = "Select HHMM, CONTROL_NO, SUM (SECS) SECS from " & ASTLOGX1 _
                & " where CONTROL_NO NOT IN ('ORDRBATCH','JOBBATCH','WEBTICK')" _
                & " group by HHMM, CONTROL_NO"
            ASCMAIN1.sql = sqlASTLOGX2
            Create_TDA(.Tables.Add, "ASTLOGX2", "**", 0, False)

            Dim sqlASTLOGX3 As String = "Select HHMM, CONTROL_NO, SUM (SECS) SECS from " & ASTLOGX1 _
                & " where CONTROL_NO IN ('ORDRBATCH','JOBBATCH','WEBTICK')" _
                & " group by HHMM, CONTROL_NO"
            ASCMAIN1.sql = sqlASTLOGX3
            Create_TDA(.Tables.Add, "ASTLOGX3", "**", 0, False)

            ASCMAIN1.sql = "Select " _
                & " HHMM, SUM (ORDW) ORDW, SUM (ORDD) ORDD, SUM (COMW) COMW, SUM (COMD) COMD, SUM (TIME) TIME from (" _
                & "Select HHMM" _
                & ", Sum (CASE WHEN LENGTH(CONTROL_NO) = 10 THEN SECS ELSE 0 END) ORDW" _
                & ", Sum (CASE WHEN LENGTH(CONTROL_NO) = 6 THEN SECS ELSE 0 END) ORDD" _
                & ", 0 COMW, 0 COMD, 0 TIME" _
                & " from (" & sqlASTLOGX2 & ") group by HHMM" _
                & " union " _
                & "Select HHMM" _
                & ", 0 ORDW, 0 ORDD" _
                & ", Sum (CASE WHEN CONTROL_NO = 'ORDRBATCH' THEN SECS ELSE 0 END) COMW" _
                & ", Sum (CASE WHEN CONTROL_NO = 'JOBBATCH' THEN SECS ELSE 0 END) COMW" _
                & ", Sum (CASE WHEN CONTROL_NO = 'WEBTICK' THEN SECS ELSE 0 END) TIME" _
                & " from (" & sqlASTLOGX3 & ") group by HHMM" _
                & ") group by HHMM"
            Create_TDA(.Tables.Add, "ASTLOGX0", "**", 0, False, "", 1)
            .Tables("ASTLOGX0").Columns.Add("TOTAL", GetType(System.Decimal), "ORDW+ORDD+COMW+COMD+TIME")

            ASCMAIN1.sql = "" _
                & "Select HHMM" _
                & ", Sum (CASE WHEN LENGTH(CONTROL_NO) = 10 THEN 1 ELSE 0 END) ORDW" _
                & ", Sum (CASE WHEN LENGTH(CONTROL_NO) = 6 THEN 1 ELSE 0 END) ORDD" _
                & " from (" & sqlASTLOGX2 & ") group by HHMM"
            Create_TDA(.Tables.Add, "ASTLOGXC", "**", 0, False, "", 1)
            With .Tables("ASTLOGXC")
                .Columns("ORDW").DataType = GetType(System.Int32)
                .Columns("ORDD").DataType = GetType(System.Int32)
            End With
            'With .Tables.Add("ASTLOGX0")
            '    .Columns.Add("HHMM")
            '    .Columns.Add("ORDW", GetType(System.Decimal))
            '    .Columns.Add("ORDD", GetType(System.Decimal))
            '    .Columns.Add("COMW", GetType(System.Decimal))
            '    .Columns.Add("COMD", GetType(System.Decimal))
            '    .Columns.Add("TIME", GetType(System.Decimal))
            'End With
        End With


        grdASTLOGX0.DataSource = dst.Tables("ASTLOGX0")
        grdASTLOGXC.DataSource = dst.Tables("ASTLOGXC")

        grdASTLOGX1.DataSource = dst.Tables("ASTLOGX1")
        grdASTLOGX2.DataSource = dst.Tables("ASTLOGX2")
        grdASTLOGX3.DataSource = dst.Tables("ASTLOGX3")

        Create_Summary(grdASTLOGX1, "LOG_ID", "Count")
        Create_Summary(grdASTLOGX1, "SECS")

        Create_Summary(grdASTLOGX2, "HHMM", "Count")
        Create_Summary(grdASTLOGX2, "SECS")
        Create_Summary(grdASTLOGX3, "HHMM", "Count")
        Create_Summary(grdASTLOGX3, "SECS")

        Create_Summary(grdASTLOGX0, "HHMM", "Count")
        Create_Summary(grdASTLOGX0, New String() {"ORDW", "ORDD", "COMW", "COMD", "TIME", "TOTAL"})
        Create_Summary(grdASTLOGXC, "HHMM", "Count")
        Create_Summary(grdASTLOGXC, New String() {"ORDW", "ORDD"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
            End With
            .Groups("Parameters").Visible = Not ScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Tables")
        Me.Cursor = Cursors.WaitCursor

        ASCDATA1.ExecuteSQL("Truncate Table " & ASTLOGX1)
        Dim Hours As Integer = Val(numHours.Value)
        ASCDATA1.ExecuteSQL("Insert into " & ASTLOGX1 & " " & Replace(sqlASTLOGX1, "2/24", CStr(Hours) & "/24"))

        Fill_Records("ASTLOGX1")
        Sort_grdColumns(grdASTLOGX1, "LOG_ID")
        Fill_Records("ASTLOGX2")
        Sort_grdColumns(grdASTLOGX2, "HHMM")
        Fill_Records("ASTLOGX3")
        Sort_grdColumns(grdASTLOGX3, "HHMM")
        Fill_Records("ASTLOGX0")
        Sort_grdColumns(grdASTLOGX0, "HHMM")
        Fill_Records("ASTLOGXC")
        Sort_grdColumns(grdASTLOGXC, "HHMM")

        CreateGraph_Trend()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("TATDATA1")
            CommitTrans("Save Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub

    Sub Clear_Record()

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "USER_ID" Then

        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "USER_ID" Then

        End If
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        ' Load_Popup_Menu(grdASTLOGX1, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

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



        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region


    Sub CreateGraph_Trend()

        chtTrend.Visible = False
        If grdASTLOGX0.Rows.Count = 0 Then Exit Sub

        Dim MINUTES As Integer = dst.Tables("ASTLOGX0").Rows.Count

        If MINUTES > 0 Then
            Dim DT As New DataTable
            DT.Columns.Add("DATA_TYPE")

            For Each rowASTLOGX0 As DataRow In dst.Tables("ASTLOGX0").Select("", "HHMM")
                Dim HHMM As String = rowASTLOGX0.Item("HHMM")
                Dim rowASTLOGXC As DataRow = dst.Tables("ASTLOGXC").Rows.Find(HHMM)
                DT.Columns.Add("HHMM " & HHMM, GetType(System.Decimal))
            Next

            DT.PrimaryKey = New DataColumn() {DT.Columns(0)}

            For Each C As String In New String() {"ORDW", "ORDD", "COMW", "COMD", "TIME"}
                DT.Rows.Add(New Object() {C})
            Next

            For Each rowASTLOGX0 As DataRow In dst.Tables("ASTLOGX0").Select("", "HHMM")
                Dim HHMM As String = rowASTLOGX0.Item("HHMM")
                DT.Rows.Find("ORDW").Item("HHMM " & HHMM) = rowASTLOGX0.Item("ORDW")
                DT.Rows.Find("ORDD").Item("HHMM " & HHMM) = rowASTLOGX0.Item("ORDD")
                DT.Rows.Find("COMW").Item("HHMM " & HHMM) = rowASTLOGX0.Item("COMW")
                DT.Rows.Find("COMD").Item("HHMM " & HHMM) = rowASTLOGX0.Item("COMD")
                DT.Rows.Find("TIME").Item("HHMM " & HHMM) = rowASTLOGX0.Item("TIME")
            Next

            Dim periods As Integer = dst.Tables("ASTLOGX0").Rows.Count

            'Dim DATA_TYPE As String = "COL_" & Format(Val(grdSATANALR.ActiveRow.Cells("ROW_NO").Value & ""), "00")
            Dim S As Integer = 1
            chtTrend.DataSource = Nothing

            Dim RL() As String = {"Web OrdPrc", "Del OrdPrc", "Web OrdCom", "Del OrdCom", "Idle"}
            Dim CL() As String
            ReDim CL(periods)

            For i As Integer = 1 To periods
                CL(i - 1) = Mid(DT.Columns(i).ColumnName, 6)
            Next

            chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

            Dim labelHash As New Hashtable()
            labelHash.Add("HIGHLOW", New MyCustomTooltip(dst))
            chtTrend.LabelHash = labelHash

            'Dim CODE1 As String = COLUMN_NAME_by_Lvl(LVL)

            Dim CAPTION As String = "How I spent my Minute"

            chtTrend.TitleTop.Text = CAPTION

            chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
            chtTrend.Tooltips.FormatString = "<HIGHLOW>"

            'Dim DTX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SATANAL1").Select, CHARTED_CODE, "DESC_VALUE")
            'Dim T As String = ""
            'For P As Integer = 1 To periods
            '    Dim COLUMN_NAME_period As String = DATA_TYPE & Format(periods - P, "0")
            '    DTX.Columns.Add(COLUMN_NAME_period, GetType(System.Decimal))
            '    T &= "+" & COLUMN_NAME_period
            'Next
            'DTX.Columns.Add(DATA_TYPE & Format(XMAX_now + 1, "0"), GetType(System.Decimal), Mid(T, 2))
            'Dim RLi As Integer = 0

            'Dim VALUE_TOTAL As Decimal = S * Val(dst.Tables("SATANAL1").Compute("SUM(" & DATA_TYPE & "13" & ")", "") & "")
            'Dim VALUE_CHARTED As Decimal = 0

            'Dim chart_all_others As Boolean = False

            'ReDim RL(DTX.Rows.Count - 1)
            ' ''chtTrend.TitleTop.Text = "Trend " & optTD.Text & " " & optTrend.Text & ", by " & optRSTSLSA1.Text

            Dim rowDT As DataRow = Nothing

            chtTrend.Data.SetRowLabels(RL)
            chtTrend.Data.SetColumnLabels(CL)

            Dim CHART_CAPTION As String = "CAPTION 2"
            chtTrend.TitleBottom.Text = CHART_CAPTION

            chtTrend.DataSource = DT
            chtTrend.DataBind()
            chtTrend.Visible = True

        End If


    End Sub

End Class

Public Class MyCustomTooltip
    Implements IRenderLabel
    Dim _dst As DataSet

    Public Sub New(ByVal dst As DataSet)
        _dst = dst
    End Sub 'New

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        Dim row As DataRow = _dst.Tables("ASTLOGXC").Rows.Find(Context("ITEM_LABEL"))
        If row Is Nothing Then Return String.Empty

        Dim ORDW As Integer = Val(row.Item("ORDW") & "")
        Dim ORDD As Integer = Val(row.Item("ORDD") & "")
        Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0")) & " Secs" _
        & vbCrLf & CStr(ORDW) & " Web Orders" _
        & vbCrLf & CStr(ORDD) & " Del Orders"

    End Function 'ToString 
End Class 'MyCustomTooltip
