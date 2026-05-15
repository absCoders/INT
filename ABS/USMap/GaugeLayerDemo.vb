
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Shared.Styles

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ Summary description for GaugeLayerDemo.
    '/ </summary>

    Public Class GaugeLayerDemo
        Inherits System.Windows.Forms.Form
        Private ultraChart1 As Infragistics.Win.UltraWinChart.UltraChart
        Private splitter2 As System.Windows.Forms.Splitter
        Private splitter1 As System.Windows.Forms.Splitter
        Private ultraExpandableGroupBox1 As Infragistics.Win.Misc.UltraExpandableGroupBox
        Private ultraExpandableGroupBoxPanel1 As Infragistics.Win.Misc.UltraExpandableGroupBoxPanel
        Private ultraLabel1 As Infragistics.Win.Misc.UltraLabel
        Private ultraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
        Private ultraGroupBox5 As Infragistics.Win.Misc.UltraGroupBox
        Private WithEvents skinEditor As Infragistics.Win.UltraWinEditors.UltraComboEditor

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
            Me.ultraGroupBox5 = New Infragistics.Win.Misc.UltraGroupBox()
            Me.skinEditor = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
            Me.splitter1 = New System.Windows.Forms.Splitter()
            Me.ultraExpandableGroupBox1 = New Infragistics.Win.Misc.UltraExpandableGroupBox()
            Me.ultraExpandableGroupBoxPanel1 = New Infragistics.Win.Misc.UltraExpandableGroupBoxPanel()
            Me.ultraLabel1 = New Infragistics.Win.Misc.UltraLabel()
            CType(Me.ultraChart1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.ultraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.ultraGroupBox2.SuspendLayout()
            CType(Me.ultraGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.ultraGroupBox5.SuspendLayout()
            CType(Me.skinEditor, System.ComponentModel.ISupportInitialize).BeginInit()
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
            Me.ultraChart1.Axis.X.Labels.Flip = False
            Me.ultraChart1.Axis.X.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.X.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.X.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.X.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X.ScrollScale.Height = 10
            Me.ultraChart1.Axis.X.ScrollScale.Visible = False
            Me.ultraChart1.Axis.X.ScrollScale.Width = 15
            Me.ultraChart1.Axis.X.TickmarkInterval = 0
            Me.ultraChart1.Axis.X2.Labels.Flip = False
            Me.ultraChart1.Axis.X2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.X2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.X2.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.X2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.X2.ScrollScale.Height = 10
            Me.ultraChart1.Axis.X2.ScrollScale.Visible = False
            Me.ultraChart1.Axis.X2.ScrollScale.Width = 15
            Me.ultraChart1.Axis.X2.TickmarkInterval = 0
            Me.ultraChart1.Axis.Y.Labels.Flip = False
            Me.ultraChart1.Axis.Y.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Y.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.Y.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y.MajorGridLines.Color = System.Drawing.Color.Transparent
            Me.ultraChart1.Axis.Y.MinorGridLines.Color = System.Drawing.Color.Transparent
            Me.ultraChart1.Axis.Y.MinorGridLines.Visible = True
            Me.ultraChart1.Axis.Y.ScrollScale.Height = 10
            Me.ultraChart1.Axis.Y.ScrollScale.Visible = False
            Me.ultraChart1.Axis.Y.ScrollScale.Width = 15
            Me.ultraChart1.Axis.Y.TickmarkInterval = 10
            Me.ultraChart1.Axis.Y.TickmarkStyle = Infragistics.UltraChart.Shared.Styles.AxisTickStyle.Smart
            Me.ultraChart1.Axis.Y2.Labels.Flip = False
            Me.ultraChart1.Axis.Y2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Y2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Y2.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.FormatString = ""
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.Y2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Y2.ScrollScale.Height = 10
            Me.ultraChart1.Axis.Y2.ScrollScale.Visible = False
            Me.ultraChart1.Axis.Y2.ScrollScale.Width = 15
            Me.ultraChart1.Axis.Y2.TickmarkInterval = 0
            Me.ultraChart1.Axis.Z.Labels.Flip = False
            Me.ultraChart1.Axis.Z.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Z.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.Z.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z.ScrollScale.Height = 10
            Me.ultraChart1.Axis.Z.ScrollScale.Visible = False
            Me.ultraChart1.Axis.Z.ScrollScale.Width = 15
            Me.ultraChart1.Axis.Z.TickmarkInterval = 0
            Me.ultraChart1.Axis.Z2.Labels.Flip = False
            Me.ultraChart1.Axis.Z2.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            Me.ultraChart1.Axis.Z2.Labels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z2.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            Me.ultraChart1.Axis.Z2.Labels.OrientationAngle = 0
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.Flip = False
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.HorizontalAlign = System.Drawing.StringAlignment.Near
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.OrientationAngle = 0
            Me.ultraChart1.Axis.Z2.Labels.SeriesLabels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z2.Labels.VerticalAlign = System.Drawing.StringAlignment.Center
            Me.ultraChart1.Axis.Z2.ScrollScale.Height = 10
            Me.ultraChart1.Axis.Z2.ScrollScale.Visible = False
            Me.ultraChart1.Axis.Z2.ScrollScale.Width = 15
            Me.ultraChart1.Axis.Z2.TickmarkInterval = 0
            Me.ultraChart1.Border.CornerRadius = 5
            Me.ultraChart1.ColorModel.AlphaLevel = CType(150, System.Byte)
            Me.ultraChart1.Data.EmptyStyle.LineStyle.DrawStyle = Infragistics.UltraChart.Shared.Styles.LineDrawStyle.Dash
            Me.ultraChart1.Data.EmptyStyle.LineStyle.EndStyle = Infragistics.UltraChart.Shared.Styles.LineCapStyle.NoAnchor
            Me.ultraChart1.Data.EmptyStyle.LineStyle.MidPointAnchors = False
            Me.ultraChart1.Data.EmptyStyle.LineStyle.StartStyle = Infragistics.UltraChart.Shared.Styles.LineCapStyle.NoAnchor
            Me.ultraChart1.Effects.Effects.Add(gradientEffect1)
            Me.ultraChart1.ForeColor = System.Drawing.SystemColors.ControlText
            Me.ultraChart1.Location = New System.Drawing.Point(8, 128)
            Me.ultraChart1.Name = "ultraChart1"
            pieChartAppearance1.Labels.Font = New System.Drawing.Font("Verdana", 7.0F)
            pieChartAppearance1.RadiusFactor = 80
            Me.ultraChart1.PieChart = pieChartAppearance1
            Me.ultraChart1.Size = New System.Drawing.Size(432, 384)
            Me.ultraChart1.TabIndex = 22
            Me.ultraChart1.Tooltips.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8F)
            Me.ultraChart1.Tooltips.HighlightDataPoint = False
            Me.ultraChart1.Tooltips.UseControl = False
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
            Me.ultraGroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.ultraGroupBox5})
            Me.ultraGroupBox2.Location = New System.Drawing.Point(448, 128)
            Me.ultraGroupBox2.Name = "ultraGroupBox2"
            Me.ultraGroupBox2.Size = New System.Drawing.Size(224, 376)
            Me.ultraGroupBox2.SupportThemes = False
            Me.ultraGroupBox2.TabIndex = 24
            ' 
            ' ultraGroupBox5
            ' 
            Me.ultraGroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.skinEditor})
            Me.ultraGroupBox5.Dock = System.Windows.Forms.DockStyle.Top
            Me.ultraGroupBox5.Location = New System.Drawing.Point(3, 2)
            Me.ultraGroupBox5.Name = "ultraGroupBox5"
            Me.ultraGroupBox5.Size = New System.Drawing.Size(218, 42)
            Me.ultraGroupBox5.SupportThemes = False
            Me.ultraGroupBox5.TabIndex = 22
            Me.ultraGroupBox5.Text = "Skin:"
            ' 
            ' skinEditor
            ' 
            Me.skinEditor.Dock = System.Windows.Forms.DockStyle.Fill
            Me.skinEditor.Location = New System.Drawing.Point(3, 16)
            Me.skinEditor.Name = "skinEditor"
            Me.skinEditor.Nullable = False
            Me.skinEditor.Size = New System.Drawing.Size(212, 21)
            Me.skinEditor.TabIndex = 0
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
            Me.ultraExpandableGroupBox1.SupportThemes = False
            Me.ultraExpandableGroupBox1.TabIndex = 23
            Me.ultraExpandableGroupBox1.Text = "Custom Gauge Layer"
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
			Me.ultraLabel1.Text = "The UltraChart uses a unique layer architecture to create and render chart images.  The controls API allows you to hook into this architecture and use it to create completely cusomized charts, such as the guage chart displayed below."
            ' 
            ' GaugeLayerDemo
            ' 
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.BackColor = System.Drawing.Color.White
            Me.ClientSize = New System.Drawing.Size(680, 526)
            Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.splitter1, Me.ultraChart1, Me.splitter2, Me.ultraGroupBox2, Me.ultraExpandableGroupBox1})
            Me.Name = "GaugeLayerDemo"
            Me.Text = "Custom Gauge Layer"
            CType(Me.ultraChart1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.ultraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ultraGroupBox2.ResumeLayout(False)
            CType(Me.ultraGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ultraGroupBox5.ResumeLayout(False)
            CType(Me.skinEditor, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.ultraExpandableGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ultraExpandableGroupBox1.ResumeLayout(False)
            Me.ultraExpandableGroupBoxPanel1.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub 'InitializeComponent 
#End Region

        Private needle1, needle2, needle3 As Needle
        Private theLayer As GaugeLayer


        Private Sub InitializeSkinsList()
            Me.skinEditor.Items.Add("(None)")
            Dim skinsDir As New DirectoryInfo(Config.DialImagesPath)
            Dim skinsFiles As FileInfo() = skinsDir.GetFiles()
            Dim fi As FileInfo
            For Each fi In skinsFiles
                Me.skinEditor.Items.Add(fi)
            Next fi
            Me.skinEditor.SelectedIndex = 0

        End Sub 'InitializeSkinsList


        Private Sub GaugeLayerDemo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

            Me.theLayer = New GaugeLayer()

            ' add the layer
            Me.theLayer.Appearance = New GaugeAppearance()

            Me.theLayer.Appearance.Radius = 180
            Me.theLayer.Appearance.Center = New Point(190, 190)

            Me.theLayer.Appearance.DialPE.Stroke = Color.Black
            Me.theLayer.Appearance.DialPE.StrokeWidth = 5

            Me.theLayer.ChartComponent = Me.ultraChart1

            Me.theLayer.Appearance.StartAngle = -30
            Me.theLayer.Appearance.EndAngle = 190
            Me.theLayer.Appearance.TextLoc = 85
            Me.theLayer.Appearance.TickStart = 65
            Me.theLayer.Appearance.TickEnd = 80

            Me.needle1 = New Needle(35, New PaintElement(Color.Red))
            Me.needle1.PE.StrokeWidth = 18
            Me.needle1.Length = 75
            Me.needle1.PE.Fill = Color.White

            Me.theLayer.Appearance.Needles.Add(Me.needle1)

            Me.needle2 = New Needle(80, New PaintElement(Color.Red))
            Me.needle2.PE.StrokeWidth = 8
            Me.needle2.Length = 60
            Me.needle2.PE.Fill = Color.Black

            Me.theLayer.Appearance.Needles.Add(Me.needle2)

            Me.needle3 = New Needle(40, New PaintElement(Color.Red))
            Me.needle3.PE.StrokeWidth = 14
            Me.needle3.Length = 70
            Me.needle3.PE.Fill = Color.Silver

            Me.theLayer.Appearance.Needles.Add(Me.needle3)

            Me.ultraChart1.Layer.Add("GaugeLayer", Me.theLayer)
            Me.ultraChart1.UserLayerIndex = New String() {"GaugeLayer"}

            ' Set axes
            Me.ultraChart1.ChartType = ChartType.PieChart
            Me.ultraChart1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:0>"
            Me.ultraChart1.Axis.Y.MajorGridLines.Thickness = 5
            Me.ultraChart1.Axis.Y.MajorGridLines.DrawStyle = LineDrawStyle.Solid
            Me.ultraChart1.Axis.Y.MajorGridLines.Color = Color.Yellow
            Me.ultraChart1.Axis.Y.Labels.FontColor = Color.Yellow
            Me.ultraChart1.Axis.Y.LineColor = Color.Yellow

            ' Set data source (this is meaningless to the gauge layer)
            Me.ultraChart1.Data.DataSource = Infragistics.UltraChart.Data.DemoTable.Table()
            Me.ultraChart1.Data.DataBind()
			Me.ultraChart1.BackgroundImage = System.Drawing.Image.FromFile(Config.ImagePath + "/chart_gray_bg.jpg")

            Me.theLayer.Appearance.Sections.Add(New GaugeSection(50))
            Me.theLayer.Appearance.Sections(0).StartWidth = 75
            Me.theLayer.Appearance.Sections(0).EndWidth = 80
            Me.theLayer.Appearance.Sections(0).PE.ElementType = PaintElementType.Gradient
            Me.theLayer.Appearance.Sections(0).PE.FillGradientStyle = GradientStyle.Horizontal
            Me.theLayer.Appearance.Sections(0).PE.Fill = Color.Green
            Me.theLayer.Appearance.Sections(0).PE.FillStopColor = Color.Yellow
            Me.theLayer.Appearance.Sections.Add(New GaugeSection(30))
            Me.theLayer.Appearance.Sections(1).StartWidth = 70
            Me.theLayer.Appearance.Sections(1).EndWidth = 80
            Me.theLayer.Appearance.Sections(1).PE.ElementType = PaintElementType.Gradient
            Me.theLayer.Appearance.Sections(1).PE.FillGradientStyle = GradientStyle.Horizontal
            Me.theLayer.Appearance.Sections(1).PE.Fill = Color.Yellow
            Me.theLayer.Appearance.Sections(1).PE.FillStopColor = Color.Orange
            Me.theLayer.Appearance.Sections.Add(New GaugeSection(20))
            Me.theLayer.Appearance.Sections(2).StartWidth = 65
            Me.theLayer.Appearance.Sections(2).EndWidth = 80
            Me.theLayer.Appearance.Sections(2).PE.ElementType = PaintElementType.Gradient
            Me.theLayer.Appearance.Sections(2).PE.FillGradientStyle = GradientStyle.Horizontal
            Me.theLayer.Appearance.Sections(2).PE.Fill = Color.Orange
            Me.theLayer.Appearance.Sections(2).PE.FillStopColor = Color.Red

            Me.InitializeSkinsList()

            Me.ultraChart1.InvalidateLayers()

        End Sub 'GaugeLayerDemo_Load


        Private Sub skinEditor_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles skinEditor.ValueChanged
            If Not (Me.skinEditor.SelectedItem.DataValue.GetType Is GetType(FileInfo)) Then
                Return
            End If

            Dim imageFile As FileInfo = CType(Me.skinEditor.SelectedItem.DataValue, FileInfo) '

            'ToDo: Error processing original source shown below
            '
            '   FileInfo imageFile = this.skinEditor.SelectedItem.DataValue as FileInfo;
            '----------------------------------------------------------------^--- Syntax error: ';' expected
            Dim imageUsed As Boolean = Not (imageFile Is Nothing)
            If imageUsed Then
                Dim img As New Bitmap(imageFile.FullName)
                Me.theLayer.Appearance.DialPE = New PaintElement(img)
                Me.ultraChart1.InvalidateLayers()
            Else
                Me.theLayer.Appearance.DialPE = New PaintElement(Color.WhiteSmoke, Color.CornflowerBlue, GradientStyle.Elliptical)
            End If
            ' our gauge layer uses the following properties to determine how
            ' sections and section labels are drawn ... if a skin image is
            ' being used then we want to turn these things off.
            Me.ultraChart1.Axis.Y.Visible = imageUsed
            Me.ultraChart1.Axis.Y.Labels.Visible = imageUsed
            Me.ultraChart1.Axis.Y.MajorGridLines.Visible = imageUsed
            Me.ultraChart1.Axis.Y.MinorGridLines.Visible = imageUsed

        End Sub 'skinEditor_ValueChanged 
    End Class 'GaugeLayerDemo
End Namespace 'ChartSamplesExplorerCS.Customization