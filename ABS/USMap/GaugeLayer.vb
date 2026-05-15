Imports System
Imports System.Collections
Imports System.Drawing
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.ComponentModel
Imports Infragistics.UltraChart.Core.ColorModel ' for IColorModel
Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Resources   ' for IChartComponent
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Resources.Editor
Imports Infragistics.UltraChart.Shared.Styles
Imports Infragistics.UltraChart.Core.Util

'/ <summary>
'/ Summary description for MyLayer.
'/ </summary>
Public Class GaugeLayer
	Implements ILayer

	' Field storage of Properties
	Protected innerBounds As New Rectangle(0, 0, 0, 0)

	' Numeric ruler for mapping purposes
	Private _Ruler As NumericRuler
	Private _Labels As New Hashtable()

	Public Sub New()
		' Create a ruler.
		Me._Ruler = New NumericRuler()
	End Sub

	Public Sub FillSceneGraph(ByVal scene As SceneGraph) Implements ILayer.FillSceneGraph


		' Draw the axes, use the settings from the y axis properties.
		Dim YApp As AxisAppearance = CType(Me.ChartComponent.GetChartAppearance(ChartAppearanceTypes.AxisY), AxisAppearance)

		' Check if appearance and Y-axis appearance are defined.
		If Not _Appearance Is Nothing AndAlso Not YApp Is Nothing Then

			' initilize the variable
			Dim min As Double = 0
			Dim max As Double = 100
			Dim delta As Double = 10

			If Me._Appearance.Layout = DialLayout.Automatic Then
				Me._Appearance.Radius = Math.Min(Me.innerBounds.Width / 3, Me.innerBounds.Height / 3)
				Me._Appearance.Center = New Point(Me.innerBounds.X + Me.innerBounds.Width / 2, Me.innerBounds.Y + Me.innerBounds.Height / 2)
			End If
			Dim sec As GaugeSection
			' check for axis settings
			' Custom
			If YApp.RangeType = AxisRangeType.Custom Then
				min = YApp.RangeMin
				max = YApp.RangeMax
			Else
				' Automatic: add up the specified sections.
				If Me._Appearance.Sections.Count > 0 Then
					min = 0

					For Each sec In Me._Appearance.Sections
						max += sec.Value
					Next sec
				End If
			End If

			' calculate the increment
			delta = (max - min) / 10

			' check for y-axis appearance
			If YApp.TickmarkStyle = AxisTickStyle.DataInterval Then
				delta = YApp.TickmarkInterval
			Else
				delta = (max - min) * YApp.TickmarkPercentage / 100
			End If

			' do setting in the ruler.
			_Ruler.Maximum = max
			_Ruler.Minimum = min

			If Me.Appearance.Direction = Direction.RightToLeft Then
				_Ruler.MapMinimum = _Appearance.StartAngle
				_Ruler.MapMaximum = _Appearance.EndAngle
			Else
				_Ruler.MapMinimum = _Appearance.EndAngle
				_Ruler.MapMaximum = _Appearance.StartAngle
			End If

			' copy scoll-scale.
			_Ruler.Scale = YApp.ScrollScale.Scale
			_Ruler.Scroll = YApp.ScrollScale.Scroll

			' start from the minimum
			Dim d_i As Double = CType(_Ruler.WindowMinimum, Double)

			' calculate various radii.
			Dim r1 As Integer = Me._Appearance.TickStart * Me._Appearance.Radius / 100
			Dim r2 As Integer = Me._Appearance.TickEnd * Me._Appearance.Radius / 100
			Dim r3 As Integer = Me._Appearance.TextLoc * Me._Appearance.Radius / 100

			' draw dial or background.
			Dim dial As New Ellipse(Me._Appearance.Center, Me._Appearance.Radius)
			dial.PE = Me.Appearance.DialPE

			' add dial background.
			scene.Add(dial)


			' draw the sections
			Dim presentVal As Double = 0
			Dim lastVal As Double = CType(_Ruler.WindowMinimum, Double)

			For Each sec In Me._Appearance.Sections
				presentVal = lastVal + sec.Value

				Dim ang0 As Integer = -CType(_Ruler.Map(lastVal), Integer)
				Dim ang1 As Integer = -CType(_Ruler.Map(presentVal), Integer)

				Dim w As New Wedge(Me._Appearance.Center, sec.EndWidth * Me._Appearance.Radius / 100, ang0, (ang1 - ang0))
				w.PE = sec.PE
				w.RadiusInner = sec.StartWidth * Me._Appearance.Radius / 100
				scene.Add(w)

				lastVal = presentVal
			Next sec


			' sanity check for increment. Without this it will go into infinite loop.
			If delta < 2 * Double.Epsilon Then delta = 5 * Double.Epsilon

			' loop thru and add the items.
			Do While d_i < CType(_Ruler.WindowMaximum, Double) + 2 * Double.Epsilon + delta
				' convert the tickmark value to angle
				Dim ang As Integer = CType(_Ruler.Map(d_i), Integer)

				' see if major grid lines are visible.
				If YApp.MajorGridLines.Visible Then
                    Dim p1 As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, r1, -Geometry.DegreeToRadian(ang))
                    Dim p2 As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, r2, -Geometry.DegreeToRadian(ang))

					Dim l As New Line(p1, p2)
					l.PE.Stroke = YApp.MajorGridLines.Color
					l.lineStyle.DrawStyle = YApp.MajorGridLines.DrawStyle
					l.PE.StrokeWidth = YApp.MajorGridLines.Thickness

					scene.Add(l)
				End If

				' see if major grid lines are visible.
				If YApp.MinorGridLines.Visible Then
					If d_i + delta / 2 < CType(_Ruler.WindowMaximum, Double) Then
						' convert the tickmark value to angle
						Dim ang1 As Integer = CType(_Ruler.Map(d_i + delta / 2), Integer)

						Dim tfp As Integer = Math.Abs((r2 - r1) / 4)
                        Dim p1 As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, r1 + tfp, -Geometry.DegreeToRadian(ang1))
                        Dim p2 As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, r2 - tfp, -Geometry.DegreeToRadian(ang1))

						Dim l As New Line(p1, p2)
						l.PE.Stroke = YApp.MinorGridLines.Color
						l.lineStyle.DrawStyle = YApp.MinorGridLines.DrawStyle
						l.PE.StrokeWidth = YApp.MinorGridLines.Thickness

						scene.Add(l)
					End If
				End If

				' see if labels are visible.
				If YApp.Labels.Visible Then
                    Dim p3 As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, r3, -Geometry.DegreeToRadian(ang))
					_Labels("DATA_VALUE") = d_i

					Dim t As New Infragistics.UltraChart.Core.Primitives.Text(p3, LabelFormatter.replaceKeywords(_Labels, YApp.Labels.ItemFormatString), YApp.Labels.LabelStyle.Copy())
					t.labelStyle.VerticalAlign = StringAlignment.Center
					t.labelStyle.HorizontalAlign = StringAlignment.Center

					t.labelStyle.Orientation = TextOrientation.Custom
					t.labelStyle.RotationAngle = ang - 90
					scene.Add(t)
				End If

				d_i += delta
			Loop

			If YApp.Visible Then
				Dim ls As New LineStyle()
				ls.DrawStyle = YApp.LineDrawStyle

				Dim el As New Arc(Me._Appearance.Center, (r1 + r2) / 2, CType(Me._Appearance.StartAngle, Single), -CType(Math.Abs(Me._Appearance.EndAngle - Me._Appearance.StartAngle), Single), ls)
				el.PE.Stroke = YApp.LineColor
				el.PE.StrokeWidth = YApp.LineThickness
				scene.Add(el)
			End If


			' sort needles according to needle length. shortest comes on the top.
			Dim ar(Me.Appearance.Needles.Count) As Double
			Dim i As Integer
			For i = 0 To Me.Appearance.Needles.Count - 1
				ar(i) = Me.Appearance.Needles(i).Length
			Next i

			Dim order() As Integer
			If ar.Length > 0 Then order = MiscFunctions.GetSortedOrderDouble(ar)

			' draw the needles.
			For i = 0 To Me.Appearance.Needles.Count - 1
				Dim nd As Needle = Me.Appearance.Needles(order(i))

				Dim theta_i As Integer = CType(_Ruler.Map(nd.Value), Integer)

                Dim p As Point = Infragistics.UltraChart.Core.Util.Geometry.AngularToCartesian(Me._Appearance.Center, nd.Length * Me._Appearance.Radius / 100, Geometry.DegreeToRadian(-theta_i))

				Dim l As New Line(Me._Appearance.Center, p)
				l.lineStyle.EndStyle = LineCapStyle.ArrowAnchor
				l.lineStyle.StartStyle = LineCapStyle.RoundAnchor

				l.PE = nd.PE

				scene.Add(l)
			Next i
		End If
	End Sub
	Private _Appearance As GaugeAppearance
	Public Property Appearance() As GaugeAppearance
		Get
			Return Me._Appearance
		End Get
		Set(ByVal Value As GaugeAppearance)
			Me._Appearance = Value
		End Set
	End Property

#Region "ILayer Implementation"
	Private _ChartColorModel As IColorModel
	Public Property ChartColorModel() As Infragistics.UltraChart.Core.ColorModel.IColorModel Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartColorModel
		Get
			Return Me._ChartColorModel
		End Get
		Set(ByVal Value As Infragistics.UltraChart.Core.ColorModel.IColorModel)
			Me._ChartColorModel = Value
		End Set
	End Property
	Private _ChartComponent As IChartComponent
	Public Property ChartComponent() As Infragistics.UltraChart.Resources.IChartComponent Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartComponent
		Get
			Return Me._ChartComponent
		End Get
		Set(ByVal Value As Infragistics.UltraChart.Resources.IChartComponent)
			Me._ChartComponent = Value
		End Set
	End Property
	Private _ChartCore As ChartCore
	Public Property ChartCore() As Infragistics.UltraChart.Core.ChartCore Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartCore
		Get
			Return Me._ChartCore
		End Get
		Set(ByVal Value As Infragistics.UltraChart.Core.ChartCore)
			Me._ChartCore = Value
		End Set
	End Property
	Private _ChartData As IChartData
	Public Property ChartData() As Infragistics.UltraChart.Data.IChartData Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartData
		Get
			Return Me._ChartData
		End Get
		Set(ByVal Value As Infragistics.UltraChart.Data.IChartData)
			Me._ChartData = Value
		End Set
	End Property

	Public Function GetDataInvalidMessage() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.GetDataInvalidMessage
		Return "Gauge Layer"
	End Function

	Public Function GetInnerBounds() As System.Drawing.Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.GetInnerBounds
		Return Me.innerBounds
	End Function
	Private _Grid As Hashtable
	Public Property Grid() As System.Collections.Hashtable Implements Infragistics.UltraChart.Core.Layers.ILayer.Grid
		Get
			Return Me._Grid
		End Get
		Set(ByVal Value As System.Collections.Hashtable)
			Me._Grid = Value
		End Set
	End Property
	Private _LayerID As String
	Public Property LayerID() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.LayerID
		Get
			return me._LayerID
		End Get
		Set(ByVal Value As String)
			Me._LayerID = Value
		End Set
	End Property
	Private _OuterBound As Rectangle
	Public Property OuterBound() As System.Drawing.Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.OuterBound
		Get

			OuterBound = Me._OuterBound
		End Get
		Set(ByVal Value As System.Drawing.Rectangle)
			Me._OuterBound = Value
			' Always remember to recalculate the InnerBound when OuterBound is changed.
			Me.CalculateInnerBounds()
		End Set
	End Property
	Private _Visible As Boolean
	Public Property Visible() As Boolean Implements Infragistics.UltraChart.Core.Layers.ILayer.Visible
		Get
			Return Me._Visible
		End Get
		Set(ByVal Value As Boolean)
			Me._Visible = Value
		End Set
	End Property
#End Region
	Public Sub CalculateInnerBounds()
		Me.innerBounds = New Rectangle(Me._OuterBound.X, Me._OuterBound.Y, Me._OuterBound.Width, Me._OuterBound.Height)
	End Sub
End Class




Public Enum Direction
	LeftToRight
	RightToLeft
End Enum

Public Enum DialLayout
	Automatic
	Manual
End Enum

Public Class GaugeAppearance

	Private _Center As Point
	Public Property Center() As Point
		Get
			Return Me._Center
		End Get
		Set(ByVal Value As Point)
			Me._Center = Value
		End Set
	End Property

	Private _Radius As Integer
	Public Property Radius() As Integer
		Get
			Return Me._Radius
		End Get
		Set(ByVal Value As Integer)
			Me._Radius = Value
		End Set
	End Property
	Private _StartAngle As Double = -45
	Public Property StartAngle() As Double
		Get
			Return Me._StartAngle
		End Get
		Set(ByVal Value As Double)
			Me._StartAngle = Value
		End Set
	End Property
	Private _EndAngle As Double = 180
	Public Property EndAngle() As Double
		Get
			Return Me._EndAngle
		End Get
		Set(ByVal Value As Double)
			Me._EndAngle = Value
		End Set
	End Property

	Private _TickStart As Integer = 70
	Public Property TickStart() As Integer
		Get
			Return Me._TickStart
		End Get
		Set(ByVal Value As Integer)
			Me._TickStart = Value
		End Set
	End Property

	Private _TickEnd As Integer = 90
	Public Property TickEnd() As Integer
		Get
			Return Me._TickEnd
		End Get
		Set(ByVal Value As Integer)
			Me._TickEnd = Value
		End Set
	End Property

	Private _TextLoc = 94
	Public Property TextLoc() As Integer
		Get
			Return Me._TextLoc
		End Get
		Set(ByVal Value As Integer)
			Me._TextLoc = Value
		End Set
	End Property

	Private _Direction As Direction
	Public Property Direction() As Direction
		Get
			Return Me._Direction
		End Get
		Set(ByVal Value As Direction)
			Me._Direction = Value
		End Set
	End Property

	Private _Needles As New NeedleCollection()
	Public ReadOnly Property Needles() As NeedleCollection
		Get
			Return Me._needles
		End Get
	End Property

	Private _Sections As New GaugeSectionCollection()
	Public ReadOnly Property Sections() As GaugeSectionCollection
		Get
			Return Me._Sections
		End Get
	End Property

	Private _DialPE As New PaintElement(Color.White, Color.Blue, GradientStyle.Elliptical)
	<TypeConverter(GetType(PaintElementConverter))> _
	 Public Property DialPE() As PaintElement
		Get
			Return Me._DialPE
		End Get
		Set(ByVal Value As PaintElement)
			Me._DialPE = Value
		End Set
	End Property

	Private _Layout As DialLayout = DialLayout.Automatic
	Public Property Layout() As DialLayout
		Get
			Return Me._Layout
		End Get
		Set(ByVal Value As DialLayout)
			Me._Layout = Value
		End Set
	End Property
End Class

Public Class Needle

	Public Sub New()
	End Sub

	Public Sub New(ByVal val As Double)
		Me._Value = val
	End Sub

	Public Sub New(ByVal val As Double, ByVal pe As PaintElement)
		Me.New(val)
		Me._PE = pe
	End Sub

	Private _Value As Double
	Public Property Value() As Double
		Get
			Return Me._Value
		End Get
		Set(ByVal Value As Double)
			Me._Value = Value
		End Set
	End Property

	Private _PE As New PaintElement()
	<TypeConverter(GetType(PaintElementConverter))> _
	 Public Property PE() As PaintElement
		Get
			Return Me._PE
		End Get
		Set(ByVal Value As PaintElement)
			_PE = Value
		End Set
	End Property

	Private _Length As Integer = 90
	Public Property Length() As Integer
		Get
			Return _Length
		End Get
		Set(ByVal Value As Integer)
			_Length = Value
		End Set
	End Property

End Class

Public Class NeedleCollection
	Inherits CollectionBase

	Default Public Property Item(ByVal index As Integer) As Needle
		Get
			Return CType(Me.List(index), Needle)
		End Get
		Set(ByVal Value As Needle)
			Me.List(index) = Value
		End Set
	End Property
	Public Function Add(ByVal value As Needle) As Integer
		Return Me.List.Add(value)
	End Function
	Public Function IndexOf(ByVal value As Needle) As Integer
		Return Me.List.IndexOf(value)
	End Function
	Public Sub Insert(ByVal index As Integer, ByVal value As Needle)
		Me.List.Insert(index, value)
	End Sub
	Public Sub Remove(ByVal value As Needle)
		Me.List.Remove(value)
	End Sub
	Public Function Contains(ByVal value As Needle) As Boolean
		' If value is not of type Needle, this will return false.
		Return Me.List.Contains(value)
	End Function
	Protected Overrides Sub OnInsert(ByVal index As Integer, ByVal value As Object)
		If Not TypeOf value Is Needle Then
			Throw New ArgumentException("value must be of type Needle.", "value")
		End If
	End Sub
	Protected Overrides Sub OnRemove(ByVal index As Integer, ByVal value As Object)
		If Not TypeOf value Is Needle Then
			Throw New ArgumentException("value must be of type Needle.", "value")
		End If
	End Sub

	protected overrides sub OnSet( index as integer, oldValue as object, newValue as object)
		If Not TypeOf newValue Is Needle Then
			Throw New ArgumentException("newValue must be of type Needle.", "newValue")
		End If
	End Sub
	Protected Overrides Sub OnValidate(ByVal value As Object)
		If Not TypeOf value Is Needle Then
			Throw New ArgumentException("value must be of type Needle.")
		End If
	End Sub
End Class

Public Class GaugeSection

	Public Sub New()

	End Sub
	Public Sub New(ByVal val As Double)
		Me._Value = val
	End Sub
	Public Sub New(ByVal val As Double, ByVal pe As PaintElement)
		Me.new(val)
		Me._PE = pe
	End Sub

	Private _Value As Double
	Public Property Value() As Double
		Get
			Return Me._Value
		End Get
		Set(ByVal Value As Double)
			Me._Value = Value
		End Set
	End Property

	Private _PE As New PaintElement()
	<TypeConverter(GetType(PaintElementConverter))> _
	Public Property PE() As PaintElement
		Get
			Return Me._PE
		End Get
		Set(ByVal Value As PaintElement)
			Me._PE = Value
		End Set
	End Property

	Private _StartWidth As Integer = 40
	Public Property StartWidth() As Integer
		Get
			Return Me._StartWidth
		End Get
		Set(ByVal Value As Integer)
			If Value <= 0 OrElse Value >= EndWidth Then
				Throw New ArgumentOutOfRangeException("StartWidth", Value, "StartWidth must be a value greater than zero and less than EndWidth.")
			End If
			Me._StartWidth = Value
		End Set
	End Property
	Private _EndWidth As Integer = 80
	Public Property EndWidth() As Integer
		Get
			Return Me._EndWidth
		End Get
		Set(ByVal Value As Integer)
			If Value <= 0 OrElse Value <= StartWidth Then
				Throw New ArgumentOutOfRangeException("EndWidth", Value, "EndWidth must be a value greater than zero and greater than StartWidth.")
			End If
			Me._EndWidth = Value
		End Set
	End Property
End Class

Public Class GaugeSectionCollection
	Inherits CollectionBase
	Default Public Property Item(ByVal index As Integer) As GaugeSection
		Get
			Return CType(Me.List(index), GaugeSection)
		End Get
		Set(ByVal Value As GaugeSection)
			Me.List(index) = Value
		End Set
	End Property
	Public Function Add(ByVal value As GaugeSection) As Integer
		Return Me.List.Add(value)
	End Function

	Public Function IndexOf(ByVal value As GaugeSection) As Integer
		Return Me.List.IndexOf(value)
	End Function
	Public Sub Insert(ByVal index As Integer, ByVal value As GaugeSection)
		List.Insert(index, value)
	End Sub
	Public Sub Remove(ByVal value As GaugeSection)
		List.Remove(value)
	End Sub
	Public Function Contains(ByVal value As GaugeSection) As Boolean
		' If value is not of type GaugeSection, this will return false.
		Return Me.List.Contains(value)
	End Function
	Protected Overrides Sub OnInsert(ByVal index As Integer, ByVal value As Object)
		If Not TypeOf value Is GaugeSection Then
			Throw New ArgumentException("value must be of type GaugeSection.", "value")
		End If
	End Sub
	Protected Overrides Sub OnRemove(ByVal index As Integer, ByVal value As Object)
		If Not TypeOf value Is GaugeSection Then
			Throw New ArgumentException("value must be of type GaugeSection.", "value")
		End If
	End Sub
	Protected Overrides Sub OnSet(ByVal index As Integer, ByVal oldValue As Object, ByVal newValue As Object)
		If Not TypeOf newValue Is GaugeSection Then
			Throw New ArgumentException("newValue must be of type GaugeSection.", "newValue")
		End If
	End Sub
	Protected Overrides Sub OnValidate(ByVal value As Object)
		If Not TypeOf value Is GaugeSection Then
			Throw New ArgumentException("value must be of type GaugeSection.")
		End If
	End Sub
End Class