
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Data
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms

Imports Infragistics.UltraChart.Shared.Styles
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Resources


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ Summary description for GaugeLayerDemo.
    '/ </summary>

    Public Class Map
        Inherits System.Windows.Forms.Form
        Private ultraChart1 As Infragistics.Win.UltraWinChart.UltraChart
        Private splitter2 As System.Windows.Forms.Splitter
        Private splitter1 As System.Windows.Forms.Splitter
        Private ultraExpandableGroupBox1 As Infragistics.Win.Misc.UltraExpandableGroupBox
        Private ultraExpandableGroupBoxPanel1 As Infragistics.Win.Misc.UltraExpandableGroupBoxPanel
        Private ultraLabel1 As Infragistics.Win.Misc.UltraLabel
        Private ultraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox

        '/ <summary>
        '/ Required designer variable.
        '/ </summary>
        Private components As System.ComponentModel.Container = Nothing


        Public Sub New()
            '
            ' Required for Windows Form Designer support
            '
            InitializeComponent()
        End Sub 'New

        '
        ' TODO: Add any constructor code after InitializeComponent call
        '

        '/ <summary>
        '/ Clean up any resources being used.
        '/ </summary>
        Protected Overloads Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub 'Dispose

#Region "Windows Form Designer generated code"

        '/ <summary>
        '/ Required method for Designer support - do not modify
        '/ the contents of this method with the code editor.
        '/ </summary>
        Private Sub InitializeComponent()
            Dim gradientEffect1 As New Infragistics.UltraChart.Resources.Appearance.GradientEffect()
            Dim pieChartAppearance1 As New Infragistics.UltraChart.Resources.Appearance.PieChartAppearance()
            Me.ultraChart1 = New Infragistics.Win.UltraWinChart.UltraChart()
            Me.splitter2 = New System.Windows.Forms.Splitter()
            Me.ultraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
            Me.splitter1 = New System.Windows.Forms.Splitter()
            Me.ultraExpandableGroupBox1 = New Infragistics.Win.Misc.UltraExpandableGroupBox()
            Me.ultraExpandableGroupBoxPanel1 = New Infragistics.Win.Misc.UltraExpandableGroupBoxPanel()
            Me.ultraLabel1 = New Infragistics.Win.Misc.UltraLabel()
            CType(Me.ultraChart1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.ultraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.ultraExpandableGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.ultraExpandableGroupBox1.SuspendLayout()
            Me.ultraExpandableGroupBoxPanel1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            '			'UltraChart' properties's serialization: Since 'ChartType' changes the way axes look,
            '			'ChartType' must be persisted ahead of any Axes change made in design time.
            '			 
            Me.ultraChart1.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.PieChart
            ' 
            ' ultraChart1
            ' 
            Me.ultraChart1.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            Me.ultraChart1.Axis.X.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.X2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y.MajorGridLines.Color = System.Drawing.Color.Transparent
            Me.ultraChart1.Axis.Y.MinorGridLines.Color = System.Drawing.Color.Transparent
            Me.ultraChart1.Axis.Y.MinorGridLines.Visible = True
            Me.ultraChart1.Axis.Y.TickmarkInterval = 10
            Me.ultraChart1.Axis.Y.TickmarkStyle = Infragistics.UltraChart.Shared.Styles.AxisTickStyle.Smart
            Me.ultraChart1.Axis.Y2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Border.CornerRadius = 5
            Me.ultraChart1.ColorModel.AlphaLevel = CType(150, System.Byte)
            Me.ultraChart1.Effects.Effects.Add(gradientEffect1)
            Me.ultraChart1.ForeColor = System.Drawing.SystemColors.ControlText
            Me.ultraChart1.Location = New System.Drawing.Point(8, 128)
            Me.ultraChart1.Name = "ultraChart1"
            pieChartAppearance1.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            pieChartAppearance1.Labels.FontColor = System.Drawing.Color.DimGray
            pieChartAppearance1.RadiusFactor = 80
            Me.ultraChart1.PieChart = pieChartAppearance1
            Me.ultraChart1.Size = New System.Drawing.Size(432, 384)
            Me.ultraChart1.TabIndex = 22
            Me.ultraChart1.Tooltips.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8F)
            Me.ultraChart1.Tooltips.HighlightFillColor = System.Drawing.Color.DimGray
            Me.ultraChart1.Tooltips.HighlightOutlineColor = System.Drawing.Color.DarkGray
            ' 
            ' splitter2
            ' 
            Me.splitter2.Dock = System.Windows.Forms.DockStyle.Top
            Me.splitter2.Name = "splitter2"
            Me.splitter2.Size = New System.Drawing.Size(680, 3)
            Me.splitter2.TabIndex = 26
            Me.splitter2.TabStop = False
            ' 
            ' ultraGroupBox2
            ' 
            Me.ultraGroupBox2.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right
            Me.ultraGroupBox2.Location = New System.Drawing.Point(448, 128)
            Me.ultraGroupBox2.Name = "ultraGroupBox2"
            Me.ultraGroupBox2.Size = New System.Drawing.Size(224, 376)
            Me.ultraGroupBox2.TabIndex = 24
            ' 
            ' splitter1
            ' 
            Me.splitter1.Dock = System.Windows.Forms.DockStyle.Right
            Me.splitter1.Location = New System.Drawing.Point(677, 3)
            Me.splitter1.Name = "splitter1"
            Me.splitter1.Size = New System.Drawing.Size(3, 523)
            Me.splitter1.TabIndex = 25
            Me.splitter1.TabStop = False
            ' 
            ' ultraExpandableGroupBox1
            ' 
            Me.ultraExpandableGroupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            Me.ultraExpandableGroupBox1.ContentPadding.Bottom = 5
            Me.ultraExpandableGroupBox1.ContentPadding.Left = 5
            Me.ultraExpandableGroupBox1.ContentPadding.Right = 5
            Me.ultraExpandableGroupBox1.ContentPadding.Top = 5
            Me.ultraExpandableGroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.ultraExpandableGroupBoxPanel1})
            Me.ultraExpandableGroupBox1.ExpandedSize = New System.Drawing.Size(664, 96)
            Me.ultraExpandableGroupBox1.HeaderClickAction = Infragistics.Win.Misc.GroupBoxHeaderClickAction.None
            Me.ultraExpandableGroupBox1.HeaderPosition = Infragistics.Win.Misc.GroupBoxHeaderPosition.TopOutsideBorder
            Me.ultraExpandableGroupBox1.Location = New System.Drawing.Point(8, 8)
            Me.ultraExpandableGroupBox1.Name = "ultraExpandableGroupBox1"
            Me.ultraExpandableGroupBox1.Size = New System.Drawing.Size(664, 112)
            Me.ultraExpandableGroupBox1.TabIndex = 23
            Me.ultraExpandableGroupBox1.Text = "Custom Map Layer"
            Me.ultraExpandableGroupBox1.ViewStyle = Infragistics.Win.Misc.GroupBoxViewStyle.Office2003
            ' 
            ' ultraExpandableGroupBoxPanel1
            ' 
            Me.ultraExpandableGroupBoxPanel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.ultraLabel1})
            Me.ultraExpandableGroupBoxPanel1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.ultraExpandableGroupBoxPanel1.Location = New System.Drawing.Point(7, 31)
            Me.ultraExpandableGroupBoxPanel1.Name = "ultraExpandableGroupBoxPanel1"
            Me.ultraExpandableGroupBoxPanel1.Size = New System.Drawing.Size(650, 74)
            Me.ultraExpandableGroupBoxPanel1.TabIndex = 0
            ' 
            ' ultraLabel1
            ' 
            Me.ultraLabel1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.ultraLabel1.Name = "ultraLabel1"
            Me.ultraLabel1.Size = New System.Drawing.Size(650, 74)
            Me.ultraLabel1.TabIndex = 1
            Me.ultraLabel1.Text = "The UltraChart uses a unique layer architecture to create and render chart images" + ".  The controls API allows you to hook into this architecture and use it to crea" + "te completely cusomized charts, such as the guage chart displayed below."
            ' 
            ' Map
            ' 
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.BackColor = System.Drawing.Color.White
            Me.ClientSize = New System.Drawing.Size(680, 526)
            Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.splitter1, Me.ultraChart1, Me.splitter2, Me.ultraGroupBox2, Me.ultraExpandableGroupBox1})
            Me.Name = "Map"
            Me.Text = "Custom Map Layer"
            CType(Me.ultraChart1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.ultraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.ultraExpandableGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ultraExpandableGroupBox1.ResumeLayout(False)
            Me.ultraExpandableGroupBoxPanel1.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub 'InitializeComponent 
#End Region




        Private Sub Map_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            '' create the layer
            Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), "../Customization/US_STATES.xml")
            Dim mp As New MapLayer(points)

            '' set the layer
            Me.ultraChart1.ChartType = ChartType.Composite
            Me.ultraChart1.CompositeChart.ChartAreas.Add(New ChartArea())
            Me.ultraChart1.UserLayerIndex = New String() {"USMap"}
            Me.ultraChart1.Layer.Add("USMap", mp)

            '' set the tooltip.
            Dim labelRenderers As New Hashtable()
            labelRenderers.Add("USMap", New USMapLabelRenderer(ExpenseByStateData()))
            Me.ultraChart1.LabelHash = labelRenderers
            Me.ultraChart1.Tooltips.FormatString = "<USMap>"

            ''set border
            Me.ultraChart1.Border.CornerRadius = 20
            Me.ultraChart1.Border.Thickness = 0
            Me.ultraChart1.BackColor = Color.White

            '' set color model
            Me.ultraChart1.ColorModel.ColorBegin = Color.Red ' Color.AliceBlue
            Me.ultraChart1.ColorModel.ColorEnd = Color.FromArgb(24, 89, 165)
            Me.ultraChart1.ColorModel.AlphaLevel = 255
            Me.ultraChart1.ColorModel.ModelStyle = ColorModels.CustomRandom ' .DataValueLinearRange

            '' legend
            Me.ultraChart1.Legend.Visible = True
            Me.ultraChart1.Axis.X.Extent = 10
            Me.ultraChart1.Legend.SpanPercentage = 10
            Me.ultraChart1.Legend.Location = LegendLocation.Right

            '' set the data
            Me.ultraChart1.Data.DataSource = StatesExpenseView()
            Me.ultraChart1.Data.DataBind()
        End Sub 'Map_Load 

#Region "ExpenseByStateData Method"

        Public Function ExpenseByStateData() As DataTable
            Dim StateInfo As StateExpenseViewInfo() = StatesExpenseView()

            Dim dt As New DataTable("States")
            dt = InitializeStatesTable()

            Dim i As Integer
            For i = 0 To StateInfo.Length - 1
                Select Case StateInfo(i).State

                    Case "Alabama"
                        dt.Rows.Find("Alabama")(1) = StateInfo(i).Amount
                    Case "Alaska"
                        dt.Rows.Find("Alaska")(1) = StateInfo(i).Amount
                    Case "Arkansas"
                        dt.Rows.Find("Arkansas")(1) = StateInfo(i).Amount
                    Case "Arizona"
                        dt.Rows.Find("Arizona")(1) = StateInfo(i).Amount
                    Case "California"
                        dt.Rows.Find("California")(1) = StateInfo(i).Amount
                    Case "Colorado"
                        dt.Rows.Find("Colorado")(1) = StateInfo(i).Amount
                    Case "Connecticut"
                        dt.Rows.Find("Connecticut")(1) = StateInfo(i).Amount
                    Case "Delaware"
                        dt.Rows.Find("Delaware")(1) = StateInfo(i).Amount
                    Case "Florida"
                        dt.Rows.Find("Florida")(1) = StateInfo(i).Amount
                    Case "Georgia"
                        dt.Rows.Find("Georgia")(1) = StateInfo(i).Amount
                    Case "Hawaii"
                        dt.Rows.Find("Hawaii")(1) = StateInfo(i).Amount
                    Case "Idaho"
                        dt.Rows.Find("Idaho")(1) = StateInfo(i).Amount
                    Case "Illinois"
                        dt.Rows.Find("Illinois")(1) = StateInfo(i).Amount
                    Case "Indiana"
                        dt.Rows.Find("Indiana")(1) = StateInfo(i).Amount
                    Case "Iowa"
                        dt.Rows.Find("Iowa")(1) = StateInfo(i).Amount
                    Case "Kansas"
                        dt.Rows.Find("Kansas")(1) = StateInfo(i).Amount
                    Case "Kentucky"
                        dt.Rows.Find("Kentucky")(1) = StateInfo(i).Amount
                    Case "Louisiana"
                        dt.Rows.Find("Louisiana")(1) = StateInfo(i).Amount
                    Case "Maine"
                        dt.Rows.Find("Maine")(1) = StateInfo(i).Amount
                    Case "Massachusetts"
                        dt.Rows.Find("Massachusetts")(1) = StateInfo(i).Amount
                    Case "Maryland"
                        dt.Rows.Find("Maryland")(1) = StateInfo(i).Amount
                    Case "Michigan"
                        dt.Rows.Find("Michigan")(1) = StateInfo(i).Amount
                    Case "Minnesota"
                        dt.Rows.Find("Minnesota")(1) = StateInfo(i).Amount
                    Case "Missouri"
                        dt.Rows.Find("Missouri")(1) = StateInfo(i).Amount
                    Case "Mississippi"
                        dt.Rows.Find("Mississippi")(1) = StateInfo(i).Amount
                    Case "Montana"
                        dt.Rows.Find("Montana")(1) = StateInfo(i).Amount
                    Case "North Carolina"
                        dt.Rows.Find("North Carolina")(1) = StateInfo(i).Amount
                    Case "North Dakota"
                        dt.Rows.Find("North Dakota")(1) = StateInfo(i).Amount
                    Case "Nebraska"
                        dt.Rows.Find("Nebraska")(1) = StateInfo(i).Amount
                    Case "New Hampshire"
                        dt.Rows.Find("New Hampshire")(1) = StateInfo(i).Amount
                    Case "New Jersey"
                        dt.Rows.Find("New Jersey")(1) = StateInfo(i).Amount
                    Case "New Mexico"
                        dt.Rows.Find("New Mexico")(1) = StateInfo(i).Amount
                    Case "Nevada"
                        dt.Rows.Find("Nevada")(1) = StateInfo(i).Amount
                    Case "New York"
                        dt.Rows.Find("New York")(1) = StateInfo(i).Amount
                    Case "Ohio"
                        dt.Rows.Find("Ohio")(1) = StateInfo(i).Amount
                    Case "Oklahoma"
                        dt.Rows.Find("Oklahoma")(1) = StateInfo(i).Amount
                    Case "Oregon"
                        dt.Rows.Find("Oregon")(1) = StateInfo(i).Amount
                    Case "Pennsylvania"
                        dt.Rows.Find("Pennsylvania")(1) = StateInfo(i).Amount
                    Case "Rhode Island"
                        dt.Rows.Find("Rhode Island")(1) = StateInfo(i).Amount
                    Case "South Carolina"
                        dt.Rows.Find("South Carolina")(1) = StateInfo(i).Amount
                    Case "South Dakota"
                        dt.Rows.Find("South Dakota")(1) = StateInfo(i).Amount
                    Case "Tennessee"
                        dt.Rows.Find("Tennessee")(1) = StateInfo(i).Amount
                    Case "Texas"
                        dt.Rows.Find("Texas")(1) = StateInfo(i).Amount
                    Case "Utah"
                        dt.Rows.Find("Utah")(1) = StateInfo(i).Amount
                    Case "Virginia"
                        dt.Rows.Find("Virginia")(1) = StateInfo(i).Amount
                    Case "Vermont"
                        dt.Rows.Find("Vermont")(1) = StateInfo(i).Amount
                    Case "Washington"
                        dt.Rows.Find("Washington")(1) = StateInfo(i).Amount
                    Case "Wisconsin"
                        dt.Rows.Find("Wisconsin")(1) = StateInfo(i).Amount
                    Case "West Virginia"
                        dt.Rows.Find("West Virginia")(1) = StateInfo(i).Amount
                    Case "Wyoming"
                        dt.Rows.Find("Wyoming")(1) = StateInfo(i).Amount
                End Select
            Next i

            Return dt
        End Function 'ExpenseByStateData
#End Region

#Region "InitializeStatesTable"

        Private Function InitializeStatesTable() As DataTable
            Dim dt As New DataTable()
            Dim chartCol As New DataColumn("State", GetType(String))

            dt.Columns.Add(chartCol)
            dt.PrimaryKey = New DataColumn() {dt.Columns("State")}
            chartCol = New DataColumn("Amount", GetType([Decimal]))
            dt.Columns.Add(chartCol)

            dt.Rows.Add(New [Object]() {"Alabama", 0})
            dt.Rows.Add(New [Object]() {"Alaska", 0})
            dt.Rows.Add(New [Object]() {"Arizona", 0})
            dt.Rows.Add(New [Object]() {"Arkansas", 0})
            dt.Rows.Add(New [Object]() {"California", 0})
            dt.Rows.Add(New [Object]() {"Colorado", 0})
            dt.Rows.Add(New [Object]() {"Connecticut", 0})
            dt.Rows.Add(New [Object]() {"Delaware", 0})
            dt.Rows.Add(New [Object]() {"Florida", 0})
            dt.Rows.Add(New [Object]() {"Georgia", 0})
            dt.Rows.Add(New [Object]() {"Hawaii", 0})
            dt.Rows.Add(New [Object]() {"Idaho", 0})
            dt.Rows.Add(New [Object]() {"Illinois", 0})
            dt.Rows.Add(New [Object]() {"Indiana", 0})
            dt.Rows.Add(New [Object]() {"Iowa", 0})
            dt.Rows.Add(New [Object]() {"Kansas", 0})
            dt.Rows.Add(New [Object]() {"Kentucky", 0})
            dt.Rows.Add(New [Object]() {"Louisiana", 0})
            dt.Rows.Add(New [Object]() {"Maine", 0})
            dt.Rows.Add(New [Object]() {"Maryland", 0})
            dt.Rows.Add(New [Object]() {"Massachusetts", 0})
            dt.Rows.Add(New [Object]() {"Michigan", 0})
            dt.Rows.Add(New [Object]() {"Minnesota", 0})
            dt.Rows.Add(New [Object]() {"Mississippi", 0})
            dt.Rows.Add(New [Object]() {"Missouri", 0})
            dt.Rows.Add(New [Object]() {"Montana", 0})
            dt.Rows.Add(New [Object]() {"Nebraska", 0})
            dt.Rows.Add(New [Object]() {"Nevada", 0})
            dt.Rows.Add(New [Object]() {"New Hampshire", 0})
            dt.Rows.Add(New [Object]() {"New Jersey", 0})
            dt.Rows.Add(New [Object]() {"New Mexico", 0})
            dt.Rows.Add(New [Object]() {"New York", 0})
            dt.Rows.Add(New [Object]() {"North Carolina", 0})
            dt.Rows.Add(New [Object]() {"North Dakota", 0})
            dt.Rows.Add(New [Object]() {"Ohio", 0})
            dt.Rows.Add(New [Object]() {"Oklahoma", 0})
            dt.Rows.Add(New [Object]() {"Oregon", 0})
            dt.Rows.Add(New [Object]() {"Pennsylvania", 0})
            dt.Rows.Add(New [Object]() {"Rhode Island", 0})
            dt.Rows.Add(New [Object]() {"South Carolina", 0})
            dt.Rows.Add(New [Object]() {"South Dakota", 0})
            dt.Rows.Add(New [Object]() {"Tennessee", 0})
            dt.Rows.Add(New [Object]() {"Texas", 0})
            dt.Rows.Add(New [Object]() {"Utah", 0})
            dt.Rows.Add(New [Object]() {"Vermont", 0})
            dt.Rows.Add(New [Object]() {"Virginia", 0})
            dt.Rows.Add(New [Object]() {"Washington", 0})
            dt.Rows.Add(New [Object]() {"West Virginia", 0})
            dt.Rows.Add(New [Object]() {"Wisconsin", 0})
            dt.Rows.Add(New [Object]() {"Wyoming", 0})

            Return dt
        End Function 'InitializeStatesTable
#End Region

#Region "Create StateExpenseView Data"

        Private Function StatesExpenseView() As StateExpenseViewInfo()
            Dim sevi(49) As StateExpenseViewInfo

            sevi(0) = New StateExpenseViewInfo("Alabama", 1915560.96, "")
            sevi(1) = New StateExpenseViewInfo("Alaska", 0, "")
            sevi(2) = New StateExpenseViewInfo("Arizona", 1915560.96, "")
            sevi(3) = New StateExpenseViewInfo("Arkansas", 9577804.8, "")
            sevi(4) = New StateExpenseViewInfo("California", 9577804.8, "")
            sevi(5) = New StateExpenseViewInfo("Colorado", 957780.48, "")
            sevi(6) = New StateExpenseViewInfo("Conneticut", 1915560.96, "")
            sevi(7) = New StateExpenseViewInfo("Delaware", 2873341.44, "")
            sevi(8) = New StateExpenseViewInfo("Florida", 957780.48, "")
            sevi(9) = New StateExpenseViewInfo("Georgia", 0, "")
            sevi(10) = New StateExpenseViewInfo("Hawaii", 0, "")
            sevi(11) = New StateExpenseViewInfo("Idaho", 0, "")
            sevi(12) = New StateExpenseViewInfo("Illinois", 957780.48, "")
            sevi(13) = New StateExpenseViewInfo("Indiana", 1915560.96, "")
            sevi(14) = New StateExpenseViewInfo("Iowa", 1.34, "")
            sevi(15) = New StateExpenseViewInfo("Kansas", 957780.48, "")
            sevi(16) = New StateExpenseViewInfo("Kentuky", 0, "")
            sevi(17) = New StateExpenseViewInfo("Louisana", 0, "")
            sevi(18) = New StateExpenseViewInfo("Maine", 0, "")
            sevi(19) = New StateExpenseViewInfo("Maryland", 5746682.88, "")
            sevi(20) = New StateExpenseViewInfo("Massachusettes", 1915560.96, "")
            sevi(21) = New StateExpenseViewInfo("Michigan", 957780.48, "")
            sevi(22) = New StateExpenseViewInfo("Minnesota", 957780.48, "")
            sevi(23) = New StateExpenseViewInfo("Missippi", 0, "")
            sevi(24) = New StateExpenseViewInfo("Missouri", 0, "")
            sevi(25) = New StateExpenseViewInfo("Montana", 0, "")
            sevi(26) = New StateExpenseViewInfo("Nebraska", 0, "")
            sevi(27) = New StateExpenseViewInfo("Nevada", 957780.48, "")
            sevi(28) = New StateExpenseViewInfo("New Hampshire", 0, "")
            sevi(29) = New StateExpenseViewInfo("New Jersey", 2873341.44, "")
            sevi(30) = New StateExpenseViewInfo("New Mexico", 0, "")
            sevi(31) = New StateExpenseViewInfo("New York", 0, "")
            sevi(32) = New StateExpenseViewInfo("North Carolina", 1915560.96, "")
            sevi(33) = New StateExpenseViewInfo("North Dakota", 0, "")
            sevi(34) = New StateExpenseViewInfo("Ohio", 1915560.96, "")
            sevi(35) = New StateExpenseViewInfo("Oklahoma", 0, "")
            sevi(36) = New StateExpenseViewInfo("Oregon", 957780.48, "")
            sevi(37) = New StateExpenseViewInfo("Pennsylvania", 2873341.44, "")
            sevi(38) = New StateExpenseViewInfo("Rhode Island", 0, "")
            sevi(39) = New StateExpenseViewInfo("South Carolina", 0, "")
            sevi(40) = New StateExpenseViewInfo("South Dakota", 0, "")
            sevi(41) = New StateExpenseViewInfo("Tennessee", 0, "")
            sevi(42) = New StateExpenseViewInfo("Texas", 957780.48, "")
            sevi(43) = New StateExpenseViewInfo("Utah", 0, "")
            sevi(44) = New StateExpenseViewInfo("Vermont", 0, "")
            sevi(45) = New StateExpenseViewInfo("Virginia", 1915560.96, "")
            sevi(46) = New StateExpenseViewInfo("Washington", 957780.48, "")
            sevi(47) = New StateExpenseViewInfo("West Virginia", 957780.48, "")
            sevi(48) = New StateExpenseViewInfo("Wisconson", 957780.48, "")
            sevi(49) = New StateExpenseViewInfo("Wyoming", 0, "")

            Return sevi
        End Function 'StatesExpenseView
#End Region
    End Class 'Map

    '/ <summary>
    '/ This custom tooltip uses the State data information
    '/ to construct a custom tooltip based on the current
    '/ state
    '/ </summary>

    Public Class USMapLabelRenderer
        Implements IRenderLabel 'ToDo: Add Implements Clauses for implementation methods of these interface(s)


        Public Sub New(ByVal info As DataTable)
            Me._InformationPerState = info
        End Sub 'New ''New
        Private _InformationPerState As DataTable

#Region "IRenderLabel Members"

        '/ <summary>
        '/ Locate the proper data value for the current state, 
        '/ construct and return the proper tooltip string
        '/ </summary>
        '/ <param name="Context"></param>
        '/ <returns></returns>
        Overloads Function ToString(ByVal Context As Hashtable) As String Implements Infragistics.UltraChart.Resources.IRenderLabel.ToString
            Dim row As Integer
            If Not (Context("DATA_ROW") Is Nothing) Then
                row = CInt(Context("DATA_ROW"))
            Else
                row = CInt(Context("ITEM_NUMBER"))
            End If

            Return _InformationPerState.Rows(row)(0) + ": Expense=" + System.Convert.ToDouble(_InformationPerState.Rows(row)(1)).ToString("0.0") ' +": Revenue=" + System.Convert.ToDouble(_InformationPerState.Rows[row][2]).ToString("0.0");
        End Function 'IRenderLabel.ToString
#End Region
    End Class 'USMapLabelRenderer ''USMapLabelRenderer
End Namespace 'ChartSamplesExplorerCS.Customization