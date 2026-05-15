
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Xml.Serialization


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ This simple class is responsible for loading a Shape File, which contains
    '/ the polygon points for the map we are going to draw.
    '/ </summary>

    Public Class ShapeFile
        Private _Shapes As New PolygonShapeCollection()


        Public ReadOnly Property Shapes() As PolygonShapeCollection
            Get
                Return _Shapes
            End Get
        End Property


        '/ <summary>
        '/ Loads the shapes from an external file
        '/ </summary>
        '/ <param name="filename"></param>
        '/ <returns></returns>
        Public Overloads Shared Function Load(ByVal filename As String) As ShapeFile
            Dim serializer As New XmlSerializer(GetType(ShapeFile))
            Dim result As ShapeFile = Nothing
            Dim reader As New StreamReader(filename)
            result = Load(reader)
            reader.Close()
            Return result
        End Function 'Load
        ''Load
        '/ <summary>
        '/ Loads the shapes from a TextReader
        '/ </summary>
        '/ <param name="reader"></param>
        '/ <returns></returns>
        Public Overloads Shared Function Load(ByVal reader As TextReader) As ShapeFile
            Dim serializer As New XmlSerializer(GetType(ShapeFile))
            Dim result As ShapeFile = Nothing
            result = CType(serializer.Deserialize(reader), ShapeFile)
            Return result
        End Function 'Load
        ''Load
        '/ <summary>
        '/ Save the existing shapes to an XML file
        '/ </summary>
        '/ <param name="filename"></param>
        Public Sub Save(ByVal filename As String)
            Dim writer As New StreamWriter(filename)
            Dim serializer As New XmlSerializer(GetType(ShapeFile))
            serializer.Serialize(writer, Me)
            writer.Close()
        End Sub 'Save ''Save
        Private BoundsUptoDate As Boolean = False
        Private _Bounds As Rectangle


        Public ReadOnly Property Bounds() As Rectangle
            Get
                If Not Me.BoundsUptoDate Then
                    Dim minX As Integer = Int32.MaxValue
                    Dim minY As Integer = Int32.MaxValue
                    Dim maxX As Integer = Int32.MinValue
                    Dim maxY As Integer = Int32.MinValue

                    Dim ps As PolygonShape
                    For Each ps In Me.Shapes
                        If ps.Bounds.X < minX Then
                            minX = ps.Bounds.X
                        End If
                        If ps.Bounds.Right > maxX Then
                            maxX = ps.Bounds.Right
                        End If
                        If ps.Bounds.Y < minY Then
                            minY = ps.Bounds.Y
                        End If
                        If ps.Bounds.Bottom > maxY Then
                            maxY = ps.Bounds.Bottom
                        End If
                    Next ps

                    Me._Bounds = New Rectangle(minX, minY, maxX - minX, maxY - minY)
                    BoundsUptoDate = True
                End If
                Return Me._Bounds
            End Get
        End Property


        Default Public Property Item(ByVal id As String) As PolygonShape
            Get
                Return Me._Shapes(id)
            End Get
            Set(ByVal Value As PolygonShape)
                Me._Shapes(id) = Value
            End Set
        End Property
    End Class 'ShapeFile ''ShapeFile
End Namespace 'ChartSamplesExplorerCS.Customization