
Imports System
Imports System.Collections
Imports System.Drawing

Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.ColorModel
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives
Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ The primary custom layer class.  This class defines each state shaped polygon, 
    '/ colors it according to the cooresponding data point value and then adds the
    '/ polygon to the SceneGraph.  This class implements ILayer which allows it to
    '/ be added to the charts Layers class as a custom layer.
    '/ </summary>

    Public Class MapLayer
        Implements ILayer 'ToDo: Add Implements Clauses for implementation methods of these interface(s)
        Private shapeFile As shapeFile = Nothing


        Public Sub New(ByVal filename As String)
            'Load the shape file which contains each states shape.
            shapeFile = shapeFile.Load(filename)
        End Sub 'New

        Public Shared STATES As String() = {"Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"}


        '/ <summary>
        '/ Method which loops through each state, locates the appropriate polygon
        '/ shape and then determines how it sohuld be added to the SceneGraph
        '/ </summary>
        '/ <param name="scene"></param>
        Public Sub FillSceneGraph(ByVal scene As SceneGraph) Implements Infragistics.UltraChart.Core.Layers.ILayer.FillSceneGraph
            'Create a background Box for the layer and color it white
            '            Box bkgnd = new Box(this._OuterBound);
            '            bkgnd.PE.Fill = Color.White;
            '            bkgnd.PE.FillOpacity = 255;
            '            scene.Add(bkgnd);
            Dim i As Integer
            For i = 0 To STATES.Length - 1
                Dim state As String = STATES(i)

                If state.StartsWith("Michigan") Then
                    'Since Michigan requires two polygons (for the LP and UP) we have to treat it different
                    AddPolygons(i, New PolygonShape() {shapeFile("Michigan0"), shapeFile("Michigan1")}, scene)
                ElseIf state.StartsWith("Hawaii") Then
                    'Since Hawaii is several polygons, we have to treat it different
                    AddPolygons(i, New PolygonShape() {shapeFile("Hawaii0"), shapeFile("Hawaii1"), shapeFile("Hawaii2"), shapeFile("Hawaii3"), shapeFile("Hawaii4")}, scene)
                Else
                    AddPolygons(i, New PolygonShape() {shapeFile(state)}, scene)
                End If
            Next i
        End Sub 'FillSceneGraph


        '/ <summary>
        '/ Method which creates each new polygon and sets its properties 
        '/ and actually adds the polygon to the SceneGraph
        '/ </summary>
        '/ <param name="index"></param>
        '/ <param name="polygonshapes"></param>
        '/ <param name="scene"></param>
        Private Sub AddPolygons(ByVal index As Integer, ByVal polygonshapes() As PolygonShape, ByVal scene As SceneGraph)
            Dim i As Integer
            For i = 0 To polygonshapes.Length - 1
                Dim polygon As New polygon(Infragistics.UltraChart.Core.Util.Transform.viewingTransform(shapeFile.Bounds, Me.OuterBound, polygonshapes(i).Points.ToArray(), True))

                Dim objectValue As Double = CDbl(Me.ChartData.GetObjectValue(index, 0))

                Console.WriteLine(objectValue.ToString())

                polygon.PE.Fill = Me._ChartColorModel.getFillColor(index, 0, objectValue)
                polygon.PE.Stroke = Me._ChartColorModel.getOutlineColor(index, 0, objectValue)
                polygon.Caps = PCaps.HitTest Or PCaps.Tooltip Or PCaps.Skin

                polygon.Row = index
                polygon.Column = 0
                polygon.Value = polygonshapes(i).Name
                polygon.Layer = Me

                scene.Add(polygon)
            Next i
        End Sub 'AddPolygons

#Region "ILayer Members"

        Private innerBounds As Rectangle

        Public Function GetInnerBounds() As Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.GetInnerBounds
            Return Me.innerBounds
        End Function 'GetInnerBounds


        Public Function GetDataInvalidMessage() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.GetDataInvalidMessage
            Return "United States"
        End Function 'GetDataInvalidMessage

        Private _Grid As New Hashtable()

        Public Property Grid() As Hashtable Implements Infragistics.UltraChart.Core.Layers.ILayer.Grid
            Get
                Return _Grid
            End Get
            Set(ByVal Value As Hashtable)
                _Grid = Value
            End Set
        End Property

        Private _LayerID As String

        Public Property LayerID() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.LayerID
            Get
                Return _LayerID
            End Get
            Set(ByVal Value As String)
                _LayerID = Value
            End Set
        End Property

        Private _ChartCore As ChartCore

        Public Property ChartCore() As ChartCore Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartCore
            Get
                Return _ChartCore
            End Get
            Set(ByVal Value As ChartCore)
                _ChartCore = Value
            End Set
        End Property

        Private _ChartData As IChartData

        Public Property ChartData() As IChartData Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartData
            Get
                Return _ChartData
            End Get
            Set(ByVal Value As IChartData)
                _ChartData = Value
            End Set
        End Property

        Private _ChartColorModel As IColorModel

        Public Property ChartColorModel() As IColorModel Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartColorModel
            Get
                Return _ChartColorModel
            End Get
            Set(ByVal Value As IColorModel)
                _ChartColorModel = Value
            End Set
        End Property

        Private _Visible As Boolean

        Public Property Visible() As Boolean Implements Infragistics.UltraChart.Core.Layers.ILayer.Visible
            Get
                Return _Visible
            End Get
            Set(ByVal Value As Boolean)
                _Visible = Value
            End Set
        End Property

        Private _ChartComponent As IChartComponent

        Public Property ChartComponent() As IChartComponent Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartComponent
            Get
                Return _ChartComponent
            End Get
            Set(ByVal Value As IChartComponent)
                _ChartComponent = Value
            End Set
        End Property

        Private _OuterBound As New Rectangle(0, 0, 0, 0)

        Public Property OuterBound() As Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.OuterBound
            Get
                Return _OuterBound
            End Get
            Set(ByVal Value As Rectangle)
                _OuterBound = Value
                CalculateInnerBounds()
            End Set
        End Property


        Protected Sub CalculateInnerBounds()
            Me.innerBounds = New Rectangle(Me._OuterBound.X, Me._OuterBound.Y, Me._OuterBound.Width, Me._OuterBound.Height)
        End Sub 'CalculateInnerBounds

#End Region
    End Class 'MapLayer
End Namespace 'ChartSamplesExplorerCS.Customization