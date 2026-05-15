Imports System.Drawing
Imports Infragistics.UltraChart.Resources

Public Class ASFBWAN1

    Dim PacketNo = 0
    Dim dte_adj As TimeSpan
    Dim dte_start As DateTime

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Show_Filter(grdASTBWAN1, True)

        With dst

            With .Tables.Add("ASTBWAN1").Columns
                .Add("PacketNo", GetType(System.Int32))
                .Add("FacilityCode", GetType(System.Int32))
                .Add("Facility")
                .Add("SeverityCode", GetType(System.Int32))
                .Add("Severity")
                .Add("Timestamp")
                .Add("Hostname")
                .Add("Message")
                .Add("Conforms", GetType(System.Boolean))
                .Add("Packet")
                '.Add("PacketB", GetType(System.Byte))
                .Add("PacketB")
                .Add("SourceAddress")
                .Add("SourcePort", GetType(System.Int32))

                .Add("msg_time", GetType(System.DateTime))
                .Add("msg_proto")
                .Add("msg_src")
                .Add("msg_srcport")
                .Add("msg_dst")
                .Add("msg_dstport")
                .Add("msg_duration", GetType(System.Int32))
                .Add("msg_sent", GetType(System.Int32))
                .Add("msg_rcvd", GetType(System.Int32))
                .Add("msg_pkts_sent", GetType(System.Int32))
                .Add("msg_pkts_rcvd", GetType(System.Int32))
                .Add("msg_msg")
                .Add("msg_op")
                .Add("msg_dstname")
            End With

            With .Tables("ASTBWAN1")
                .PrimaryKey = New DataColumn() {.Columns("PacketNo")}
            End With




            With .Tables.Add("ASTBWAN2").Columns
                .Add("PacketNo", GetType(System.Int32))
                .Add("Key")
                .Add("Text")
            End With

            With .Tables("ASTBWAN2")
                .PrimaryKey = New DataColumn() {.Columns("PacketNo"), .Columns("Key")}
            End With


            With .Tables.Add("ASTBWANK").Columns
                .Add("Key")
                .Add("LastText")
                .Add("Uses", GetType(System.Int32))
                .Add("LastPacketNo", GetType(System.Int32))
            End With

            With .Tables("ASTBWANK")
                .PrimaryKey = New DataColumn() {.Columns("Key")}
            End With


            With .Tables.Add("ASTBWANA").Columns
                .Add("IP")
                .Add("HostName")
                .Add("RequestID")
                .Add("Packets_as_src", GetType(System.Int32))
                .Add("sent_as_src", GetType(System.Int64))
                .Add("rcvd_as_src", GetType(System.Int64))
                .Add("Packets_as_dst", GetType(System.Int32))
                .Add("sent_as_dst", GetType(System.Int64))
                .Add("rcvd_as_dst", GetType(System.Int64))
            End With

            With .Tables("ASTBWANA")
                .PrimaryKey = New DataColumn() {.Columns("IP")}
            End With
        End With

        grdASTBWAN1.DataSource = dst.Tables("ASTBWAN1")
        grdASTBWAN2.DataSource = dst.Tables("ASTBWAN2")

        grdASTBWANK.DataSource = dst.Tables("ASTBWANK")
        grdASTBWANA.DataSource = dst.Tables("ASTBWANA")


        With grdASTBWAN1.DisplayLayout.Bands(0)
            .Columns("msg_time").Format = "HH:mm:ss"
        End With

        Create_Summary(grdASTBWAN1, "PacketNo", "Count")
        Create_Summary(grdASTBWAN1, "msg_sent")
        Create_Summary(grdASTBWAN1, "msg_rcvd")
        Create_Summary(grdASTBWAN1, "msg_pkts_sent")
        Create_Summary(grdASTBWAN1, "msg_pkts_rcvd")

        Create_Summary(grdASTBWAN2, "PacketNo", "Count")

        Create_Summary(grdASTBWANA, "IP", "Count")
        Create_Summary(grdASTBWANA, "Packets_as_src")
        Create_Summary(grdASTBWANA, "sent_as_src")
        Create_Summary(grdASTBWANA, "rcvd_as_src")
        Create_Summary(grdASTBWANA, "Packets_as_dst")
        Create_Summary(grdASTBWANA, "sent_as_dst")
        Create_Summary(grdASTBWANA, "rcvd_as_dst")

        Create_Summary(grdASTBWANK, "Key", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

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
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode
        Syslog1.Active = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading")
        Me.Cursor = Cursors.WaitCursor

        PacketNo = 0
        dst.Tables("ASTBWAN1").Clear()
        dst.Tables("ASTBWAN2").Clear()
        dst.Tables("ASTBWANK").Clear()
        dst.Tables("ASTBWANA").Clear()

        Sort_grdColumns(grdASTBWAN1, "PacketNo")

        Show_Fields(False)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        UltraExplorerBar1.Groups("Charts").Visible = False

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTBWAN1, "SSS", "Show Filter", "Show GroupBox", "Show syslog Fields")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If
        If tlb_pop.Tools.Exists("Show syslog Fields") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show syslog Fields"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("FacilityCode").Hidden
        End If



        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Show syslog Fields"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Fields(tlb_sbt.Checked)

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "grdASTBWAN1"
    Private Sub grdASTBWAN1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTBWAN1.AfterRowActivate
        Setup_ASTBWAN2_Packets()
    End Sub

    Private Sub grdASTBWAN1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTBWAN1.InitializeRow

    End Sub

#End Region

    Private Sub grdASTBWAN1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTBWAN1.InitializeLayout

    End Sub

    Private Delegate Sub Packet_Received_Delegate(ByVal e As nsoftware.IPWorks.SyslogPacketInEventArgs)

    Sub Packet_Received(ByVal e As nsoftware.IPWorks.SyslogPacketInEventArgs)


        PacketNo += 1

        Dim rowASTBWAN1 As DataRow = dst.Tables("ASTBWAN1").NewRow
        rowASTBWAN1.Item("PacketNo") = PacketNo
        rowASTBWAN1.Item("FacilityCode") = e.FacilityCode
        rowASTBWAN1.Item("Facility") = e.Facility
        rowASTBWAN1.Item("SeverityCode") = e.SeverityCode
        rowASTBWAN1.Item("Severity") = e.Severity
        rowASTBWAN1.Item("Timestamp") = e.Timestamp
        rowASTBWAN1.Item("Hostname") = e.Hostname

        rowASTBWAN1.Item("Message") = e.Message


        Dim msg As String = e.Message

        Do While msg <> ""
            Dim i As Integer = InStr(msg, "=")
            If i = 0 Then Stop
            Dim msgkey As String = Mid(msg, 1, i - 1)
            msg = Mid(msg, i + 1)
            Dim j As Integer
            Dim q As Boolean = False
            If Mid(msg, 1, 1) = Chr(34) Then
                j = InStr(Mid(msg, 2) & Chr(34) & " ", Chr(34) & " ") + 2
                q = True
            Else
                j = InStr(msg & " ", " ")
            End If
            Dim msgtext = Mid(msg, 1, j - 1)
            If q Then
                msgtext = Mid(msgtext, 2, Len(msgtext) - 2)
            End If
            msg = Mid(msg, j + 1)

            dst.Tables("ASTBWAN2").Rows.Add(New Object() {PacketNo, msgkey, msgtext})

            Dim rowASTBWANK As DataRow = dst.Tables("ASTBWANK").Rows.Find(msgkey)
            If rowASTBWANK Is Nothing Then
                rowASTBWANK = dst.Tables("ASTBWANK").NewRow
                rowASTBWANK.Item("Key") = msgkey
                dst.Tables("ASTBWANK").Rows.Add(rowASTBWANK)
            End If
            rowASTBWANK.Item("LastText") = msgtext
            rowASTBWANK.Item("LastPacketNo") = PacketNo
            rowASTBWANK.Item("Uses") = Val(rowASTBWANK.Item("Uses") & "") + 1

            If New String() {"time", "proto", "src", "srcport", "dst", "dstport", _
                             "duration", "sent", "rcvd", "pkts_sent", "pkts_rcvd", _
                             "msg", "op", "dstname"}.Contains(msgkey) Then
                rowASTBWAN1.Item("msg_" & msgkey) = msgtext
            End If

            If msgkey = "from_ip" Then
                rowASTBWAN1.Item("msg_src") = msgtext
            ElseIf msgkey = "to_ip" Then
                rowASTBWAN1.Item("msg_dst") = msgtext
            ElseIf msgkey = "source_port" Then
                rowASTBWAN1.Item("msg_srcport") = msgtext
            ElseIf msgkey = "dest_port" Then
                rowASTBWAN1.Item("msg_dstport") = msgtext
            ElseIf msgkey = "protocol" Then
                rowASTBWAN1.Item("msg_proto") = msgtext
            ElseIf msgkey = "protocol" Then
                rowASTBWAN1.Item("msg_proto") = msgtext
            ElseIf msgkey = "protocol" Then
                rowASTBWAN1.Item("msg_proto") = msgtext
            ElseIf msgkey = "protocol" Then
                rowASTBWAN1.Item("msg_proto") = msgtext
            End If

            If msgkey = "time" Then
                If PacketNo = 1 Then
                    dte_start = Now
                    dte_adj = Now.Subtract(rowASTBWAN1.Item("msg_time"))
                End If

                rowASTBWAN1.Item("msg_time") = CDate(rowASTBWAN1.Item("msg_time")).Add(dte_adj)
            End If

        Loop


        Dim rowASTBWANA As DataRow
        Dim IP As String

        IP = rowASTBWAN1.Item("msg_src") & ""
        If IP <> "" Then
            rowASTBWANA = dst.Tables("ASTBWANA").Rows.Find(IP)
            If rowASTBWANA Is Nothing Then
                rowASTBWANA = dst.Tables("ASTBWANA").NewRow
                rowASTBWANA.Item("IP") = IP
                dst.Tables("ASTBWANA").Rows.Add(rowASTBWANA)

                'Ipinfo1.HostAddress = IP
                'rowASTBWANA.Item("RequestId") = Ipinfo1.RequestId

            End If
            rowASTBWANA.Item("Packets_as_src") = Val(rowASTBWANA.Item("Packets_as_src") & "") + Val(rowASTBWAN1.Item("msg_pkts_sent") & "") + Val(rowASTBWAN1.Item("msg_pkts_rcvd") & "")
            rowASTBWANA.Item("sent_as_src") = Val(rowASTBWANA.Item("sent_as_src") & "") + Val(rowASTBWAN1.Item("msg_sent") & "")
            rowASTBWANA.Item("rcvd_as_src") = Val(rowASTBWANA.Item("rcvd_as_src") & "") + Val(rowASTBWAN1.Item("msg_rcvd") & "")
        End If

        IP = rowASTBWAN1.Item("msg_dst") & ""
        If IP <> "" Then
            rowASTBWANA = dst.Tables("ASTBWANA").Rows.Find(IP)
            If rowASTBWANA Is Nothing Then
                rowASTBWANA = dst.Tables("ASTBWANA").NewRow
                rowASTBWANA.Item("IP") = IP
                dst.Tables("ASTBWANA").Rows.Add(rowASTBWANA)

                'Ipinfo1.HostAddress = IP
                'rowASTBWANA.Item("RequestId") = Ipinfo1.RequestId
            End If
            rowASTBWANA.Item("Packets_as_dst") = Val(rowASTBWANA.Item("Packets_as_dst") & "") + Val(rowASTBWAN1.Item("msg_pkts_sent") & "") + Val(rowASTBWAN1.Item("msg_pkts_rcvd") & "")
            rowASTBWANA.Item("sent_as_dst") = Val(rowASTBWANA.Item("sent_as_dst") & "") + Val(rowASTBWAN1.Item("msg_rcvd") & "")
            rowASTBWANA.Item("rcvd_as_dst") = Val(rowASTBWANA.Item("rcvd_as_dst") & "") + Val(rowASTBWAN1.Item("msg_sent") & "")
        End If


        rowASTBWAN1.Item("Conforms") = e.Conforms
        rowASTBWAN1.Item("Packet") = e.Packet

        Dim enc As System.Text.ASCIIEncoding = New System.Text.ASCIIEncoding()
        Dim str As String = enc.GetString(e.PacketB)

        rowASTBWAN1.Item("PacketB") = str
        rowASTBWAN1.Item("SourceAddress") = e.SourceAddress
        rowASTBWAN1.Item("SourcePort") = e.SourcePort
        dst.Tables("ASTBWAN1").Rows.Add(rowASTBWAN1)
    End Sub


    Private Sub Syslog1_OnPacketIn(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.SyslogPacketInEventArgs) Handles Syslog1.OnPacketIn
        Me.Invoke(New Packet_Received_Delegate(AddressOf Packet_Received), e)
    End Sub

    Sub Show_Fields(ByVal tf As Boolean)
        For Each COLUMN_NAME In New String() _
        {"FacilityCode", "Facility", "SeverityCode", "Severity", _
         "Timestamp", "Hostname", "Message", "Conforms", "Packet", "PacketB", "SourceAddress", "SourcePort"}
            With grdASTBWAN1.DisplayLayout.Bands(0)
                .Columns(COLUMN_NAME).Hidden = Not tf
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            End With
        Next
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Select Case tabMain.SelectedTab.Key
            Case "Packets"
                grdASTBWAN2.Parent = splPacketDetails.Panel1
                Setup_ASTBWAN2_Packets()

                grdASTBWAN1.Parent = splPackets.Panel1
                grdASTBWAN1.DisplayLayout.CaptionVisible = DefaultableBoolean.False
                Dim dvw As DataView = dst.Tables("ASTBWAN1").DefaultView
                dvw.RowFilter = ""
                Sort_grdColumns(grdASTBWAN1, "PacketNo")

            Case "Message Keys"
                grdASTBWAN2.Parent = splKeys.Panel2
                Setup_ASTBWAN2_Keys()

            Case "IPs"
                grdASTBWAN1.Parent = tabIP.Tabs("Packets").TabPage
                grdASTBWAN1.DisplayLayout.CaptionVisible = DefaultableBoolean.True
                Setup_ASTBWAN1_IP()
                Prepare_Charts()
        End Select

        Setup_tabIP()
    End Sub

    Sub Setup_ASTBWAN1_IP()
        If grdASTBWANA.ActiveRow Is Nothing Then
            splIP.Panel2Collapsed = True
        Else
            If grdASTBWANA.ActiveRow.IsDataRow Then
                Dim dvw As DataView = dst.Tables("ASTBWAN1").DefaultView
                dvw.RowFilter = "msg_src = '" & grdASTBWANA.ActiveRow.Cells("IP").Value & "' or msg_dst = '" & grdASTBWANA.ActiveRow.Cells("IP").Value & "'"
            End If
            splIP.Panel2Collapsed = False
        End If
    End Sub

    Sub Setup_ASTBWAN2_Packets()
        If grdASTBWAN1.ActiveRow Is Nothing Then
            splPackets.Panel2Collapsed = True
        Else
            If grdASTBWAN1.ActiveRow.IsDataRow Then
                txtMessage.Text = grdASTBWAN1.ActiveRow.Cells("Message").Text

                Dim dvw As DataView = dst.Tables("ASTBWAN2").DefaultView
                dvw.RowFilter = "PacketNo = " & grdASTBWAN1.ActiveRow.Cells("PacketNo").Value
            End If

            splPackets.Panel2Collapsed = Not (grdASTBWAN1.ActiveRow.IsDataRow)
        End If
    End Sub

    Sub Setup_ASTBWAN2_Keys()
        If grdASTBWANK.ActiveRow Is Nothing Then
            splKeys.Panel2Collapsed = True
        Else
            If grdASTBWANK.ActiveRow.IsDataRow Then
                Dim dvw As DataView = dst.Tables("ASTBWAN2").DefaultView
                dvw.RowFilter = "Key = '" & grdASTBWANK.ActiveRow.Cells("Key").Value & "'"
            End If

            splKeys.Panel2Collapsed = Not (grdASTBWANK.ActiveRow.IsDataRow)
        End If
    End Sub

    'Private Sub Ipinfo1_OnRequestComplete(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.IpinfoRequestCompleteEventArgs)
    '    Me.Invoke(New IPInfo_Received_Delegate(AddressOf IPInfo_Received), e)
    'End Sub

    'Private Delegate Sub IPInfo_Received_Delegate(ByVal e As nsoftware.IPWorks.IpinfoRequestCompleteEventArgs)

    'Sub IPInfo_Received(ByVal e As nsoftware.IPWorks.IpinfoRequestCompleteEventArgs)
    '    Stop
    '    Dim RequestID As String = e.RequestId
    '    Dim row As DataRow = dst.Tables("ASTBWANA").Select("RequestID = '" & RequestID & "'")(0)
    'End Sub



#Region "Charts"

    Sub Prepare_Charts()
        If SELECTION_NO = 0 Then Exit Sub

        CreateGraph_Totals()
        CreateGraph_Trend()
        chtTotals.Visible = True
        chtTrend.Visible = True
        
        For Each COLUMN_NAME As String In New String() _
        {"Packets_as_src", "sent_as_src", "rcvd_as_src", "Packets_as_dst", "sent_as_dst", "rcvd_as_dst"}
            With grdASTBWANA.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellAppearance
                If COLUMN_NAME = optChartTrendData.Value Then
                    .BackColor = Color.Yellow
                Else
                    .BackColor = Color.Empty
                End If
            End With
        Next
        Dim COLUMN_NAME_charted As String = optChartTrendData.Value
        Sort_grdColumns(grdASTBWANA, COLUMN_NAME_charted.ToLower)
    End Sub

    Sub CreateGraph_Totals()

        Dim chtIsVisible As Boolean = chtTotals.Visible
        chtTotals.Visible = False

        chtTotals.DataSource = Nothing

        Dim DATA_TYPE As String = optChartTrendData.Value

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

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

        Dim DTX As DataTable = dst.Tables("ASTBWANA")

        Dim PCT_at_TOP_N As Decimal = 0
        Dim VALUE_TOTAL As Decimal = Val(DTX.Compute("SUM(" & DATA_TYPE & ")", "") & "")
        Dim VALUE_CHARTED As Decimal = 0

        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", DATA_TYPE & " DESC")
            RL(RLi) = row.Item("IP") & ":" & row("HOSTNAME")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("IP"), row.Item(DATA_TYPE)})

            If optChartTrend.Value = "N" And RLi <= Val(numChartTrend.Value & "") Then
                PCT_at_TOP_N = 100 * Val(row.Item(DATA_TYPE & "00")) / VALUE_TOTAL
            End If
        Next

        chtTotals.TitleTop.Text = "Total - " & DATA_TYPE
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

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_Trend()

        Dim chtIsVisible As Boolean = chtTrend.Visible
        chtTrend.Visible = False

        Dim periods As Integer = 10
        Dim dte_finish As DateTime = Now

        Dim DATA_TYPE As String = optChartTrendData.Value
        Dim S As Integer = 1
        'If DATA_TYPE = "R" Then
        '    S = -1
        'End If

        chtTrend.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(periods)

        Dim dte() As DateTime
        ReDim dte(periods)

        Dim tcks_total As Int64 = dte_finish.Subtract(dte_start).Ticks

        dte(0) = dte_start
        For i As Integer = 1 To periods
            dte(i) = dte(i - 1).AddTicks(dte_finish.Subtract(dte_start).Ticks / periods)
            CL(i - 1) = Format(dte(i), "HH:mm")
        Next

        chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTrend.LabelHash = labelHash

        chtTrend.TitleTop.Text = "Trend - " & optChartTrendData.Text

        chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTrend.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To periods
            DT.Columns.Add("P" & Format(P - 1, "0"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0


        Dim DTX As New DataTable
        With DTX.Columns
            .Add("IP")
            .Add("HOSTNAME")
            Dim TOTAL As String = ""
            For P As Integer = 0 To 9
                .Add("P" & Format(P, "0"), GetType(System.Decimal))
                TOTAL &= "+ISNULL(P" & Format(P, "0") & ",0)"
            Next
            .Add(DATA_TYPE, GetType(System.Decimal), Mid(TOTAL, 2))
        End With
        DTX.PrimaryKey = New DataColumn() {DTX.Columns("IP")}

        For Each rowASTBWAN1 As DataRow In dst.Tables("ASTBWAN1").Rows
            Dim IP As String
            If New String() {"Packets_as_src", "sent_as_src", "rcvd_as_src"}.Contains(optChartTrendData.Value) Then
                IP = rowASTBWAN1.Item("msg_src") & ""
            Else
                IP = rowASTBWAN1.Item("msg_dst") & ""
            End If
            If IP <> "" Then
                Dim row As DataRow = DTX.Rows.Find(IP)
                If row Is Nothing Then
                    row = DTX.NewRow
                    row.Item("IP") = IP
                    Dim rowASTBWANA As DataRow = dst.Tables("ASTBWANA").Rows.Find(IP)
                    row.Item("HOSTNAME") = rowASTBWANA.Item("HOSTNAME")
                    DTX.Rows.Add(row)
                End If
                Dim DTM As DateTime = Now ' rowASTBWAN1.Item("msg_time")
                Dim tcks As Int64 = 0
                If DTM.CompareTo(dte_start) < 0 Then
                    'Stop
                    tcks = -1
                ElseIf DTM.CompareTo(dte_finish) > 0 Then
                    'Stop
                    tcks = -1
                Else
                    tcks = DTM.Subtract(dte_start).Ticks
                    tcks_total = dte_finish.Subtract(dte_start).Ticks
                    Dim P As Integer = 9 * (tcks / tcks_total)
                    If New String() {"Packets_as_src", "Packets_as_dst"}.Contains(optChartTrendData.Value) Then
                        row.Item("P" & CStr(P)) = Val(row.Item("P" & CStr(P)) & "") + Val(rowASTBWAN1.Item("msg_pkts_sent") & "") + Val(rowASTBWAN1.Item("msg_pkts_rcvd") & "")
                    ElseIf New String() {"sent_as_src", "sent_as_dst"}.Contains(optChartTrendData.Value) Then
                        row.Item("P" & CStr(P)) = Val(row.Item("P" & CStr(P)) & "") + Val(rowASTBWAN1.Item("msg_sent") & "")
                    Else
                        row.Item("P" & CStr(P)) = Val(row.Item("P" & CStr(P)) & "") + Val(rowASTBWAN1.Item("msg_rcvd") & "")
                    End If
                End If
            End If
        Next

        Dim sqlw As String = ""

        Dim VALUE_TOTAL As Decimal = S * Val(DTX.Compute("SUM(" & DATA_TYPE & ")", sqlw) & "")
        Dim VALUE_CHARTED As Decimal = 0

        Dim chart_all_others As Boolean = False

        ReDim RL(DTX.Rows.Count - 1)
        ''chtTrend.TitleTop.Text = "Trend " & optTD.Text & " " & optTrend.Text & ", by " & optRSTSLSA1.Text

        Dim rowDT As DataRow = Nothing


        For Each row As DataRow In DTX.Select(sqlw, DATA_TYPE & " DESC")
            Dim this_record_is_others As Boolean = False

            Dim VALUE_this_record As Decimal = S * Val(row.Item(DATA_TYPE) & "")
            Dim CODE_VALUE As String = row.Item("IP") & ""
            Dim DESC_VALUE As String = row.Item("HOSTNAME") & ""

            If (optChartTrend.Value = "C" And VALUE_TOTAL > 0 AndAlso 100 * VALUE_this_record / VALUE_TOTAL > Val(numChartTrend.Value & "")) _
            Or (optChartTrend.Value = "N" And RLi < Val(numChartTrend.Value & "")) Then
            Else
                this_record_is_others = True
                CODE_VALUE = "Z"
                DESC_VALUE = "All Others"
            End If

            If Not this_record_is_others Or chart_all_others Then
                If RLi <> 0 AndAlso RL(RLi - 1) = CODE_VALUE & ":" & DESC_VALUE Then
                Else
                    RL(RLi) = CODE_VALUE & ":" & DESC_VALUE
                    RLi += 1
                    rowDT = DT.NewRow
                    rowDT.Item("CODE_VALUE") = CODE_VALUE
                    rowDT.Item("DESC_VALUE") = DESC_VALUE
                    DT.Rows.Add(rowDT)
                End If

                VALUE_CHARTED += +Val(row.Item(DATA_TYPE) & "")

                For P As Integer = 1 To periods
                    Dim COLUMN_NAME_period As String = "P" & Format(P - 1, "0")
                    Dim VALUE As Decimal = Val(row.Item(COLUMN_NAME_period) & "")

                    VALUE = VALUE / (dte_finish.Subtract(dte_start).TotalSeconds / periods)

                    rowDT.Item("P" & Format(P - 1, "0")) = Val(rowDT.Item("P" & Format(P - 1, "0")) & "") _
                                                      + S * VALUE
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

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub chtTrend_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTrend.ChartDataClicked
        Select_CODE_VALUE_from_Chart(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Private Sub cmdChartRedraw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChartRedraw.Click

        Prepare_Charts()
        'CreateGraph_Totals()
        'CreateGraph_Trend()
    End Sub

    Private Sub tbkChartTrend_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbkChartTrend.Scroll
        chtTrend.Axis.Y.ScrollScale.Scale = (100 - Me.tbkChartTrend.Value) / 100.0
    End Sub

    Private Sub chtTotals_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTotals.ChartDataClicked
        Select_CODE_VALUE_from_Chart(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Sub Select_CODE_VALUE_from_Chart(ByVal CODE_VALUE As String)
        'Stop
        'For Each grow As UltraWinGrid.UltraGridRow In grdTATDASHX.Rows
        '    If grow.Cells("CODE1").Value & "" = CODE_VALUE Then
        '        grdTATDASHX.ActiveRow = grow
        '        grdTATDASHX.Selected.Rows.Clear()
        '        grow.Selected = True
        '        Exit Sub
        '    End If
        'Next
    End Sub

#End Region




    Public Class MyCustomTooltip
        Implements IRenderLabel

        Public Sub New()

        End Sub 'New

        Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
            'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
            'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
            Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

        End Function 'ToString 
    End Class 'MyCustomTooltip

    Private Sub tabIP_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabIP.SelectedTabChanged
        Setup_tabIP()
    End Sub

    Sub Setup_tabIP()
        If SELECTION_NO = 0 Then Exit Sub

        If Not ScreenMode Then
            UltraExplorerBar1.Groups("Charts").Visible = False
        Else
            UltraExplorerBar1.Groups("Charts").Visible = (tabMain.SelectedTab.Key = "IPs") And (tabIP.SelectedTab.Key = "Usage")
        End If
    End Sub

    Private Sub grdASTBWANK_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTBWANK.AfterRowActivate
        Setup_ASTBWAN2_Keys()
    End Sub

    Private Sub grdASTBWANK_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTBWANK.InitializeLayout

    End Sub

    Private Sub optChartTrendData_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChartTrendData.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        Prepare_Charts()
    End Sub

    Private Sub grdASTBWANA_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTBWANA.AfterRowActivate
        Setup_ASTBWAN1_IP()

        Dim IP As String = grdASTBWANA.ActiveRow.Cells("IP").Text
        grdASTBWAN1.Text = "Packets to or from IP " & IP

    End Sub

    Private Sub grdASTBWANA_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTBWANA.InitializeLayout

    End Sub
End Class