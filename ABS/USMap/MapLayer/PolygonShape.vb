
Imports System
Imports System.Collections
Imports System.Drawing
Imports System.Xml.Serialization


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ This custom collection allows us to collect the Points 
    '/ required for each state polygon into a single location
    '/ </summary>

    Public Class PointCollection
        Inherits CollectionBase

        Public Overridable Function Add(ByVal point As Point) As Integer
            Return Me.List.Add(point)
        End Function 'Add


        Default Public Overridable Property Item(ByVal index As Integer) As Point
            Get
                Return CType(Me.List(index), Point)
            End Get
            Set(ByVal Value As Point)
                Me(index) = Value
            End Set
        End Property


        Public Overridable Function ToArray() As Point()
            Dim points(Me.Count - 1) As Point
            Dim current As Integer
            For current = 0 To (Me.Count) - 1
                points(current) = Me(current)
            Next current
            Return points
        End Function 'ToArray
    End Class 'PointCollection

    '/ <summary>
    '/ The PolygonShape class contains each states polygon shape,
    '/ which is deserialized from an external XML file
    '/ </summary>

    Public Class PolygonShape
        Private _Name As String


        <XmlAttributeAttribute()> _
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal Value As String)
                _Name = Value
            End Set
        End Property

        Private _Points As New PointCollection()

        Public ReadOnly Property Points() As PointCollection
            Get
                Return _Points
            End Get
        End Property

        Private BoundsUptoDate As Boolean = False
        Private _Bounds As Rectangle

        Public ReadOnly Property Bounds() As Rectangle
            Get
                If Not Me.BoundsUptoDate Then
                    Dim minX As Integer = Int32.MaxValue
                    Dim minY As Integer = Int32.MaxValue
                    Dim maxX As Integer = Int32.MinValue
                    Dim maxY As Integer = Int32.MinValue


                    Dim p As Point
                    For Each p In Me._Points
                        If p.X < minX Then
                            minX = p.X
                        End If
                        If p.X > maxX Then
                            maxX = p.X
                        End If
                        If p.Y < minY Then
                            minY = p.Y
                        End If
                        If p.Y > maxY Then
                            maxY = p.Y
                        End If
                    Next p
                    Me._Bounds = New Rectangle(minX, minY, maxX - minX, maxY - minY)
                    BoundsUptoDate = True
                End If
                Return Me._Bounds
            End Get
        End Property
    End Class 'PolygonShape ''PolygonShape
End Namespace 'ChartSamplesExplorerCS.Customization