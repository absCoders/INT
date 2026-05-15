
Imports System
Imports System.Collections


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ This custom collection is used to store all of the
    '/ polygon shapes that we need to draw the US map
    '/ </summary>

    Public Class PolygonShapeCollection
        Inherits CollectionBase


        Default Public Property Item(ByVal id As String) As PolygonShape
            Get
                Return SearchForId(id)
            End Get
            Set(ByVal Value As PolygonShape)
                Dim e As PolygonShape = SearchForId(id)
                If e Is Nothing Then
                    Me.Add(Value)
                Else
                    Me(Me.IndexOf(e)) = Value
                End If
            End Set
        End Property


        Private Function SearchForId(ByVal id As String) As PolygonShape
            Dim result As PolygonShape = Nothing

            Dim ef As PolygonShape
            For Each ef In Me
                If ef.Name.Equals(id) Then
                    Return ef
                End If
            Next ef

            Return result
        End Function 'SearchForId 
        ''SearchForId


        Default Public Property Item(ByVal index As Integer) As PolygonShape
            Get
                Return CType(Me(index), PolygonShape)
            End Get
            Set(ByVal Value As PolygonShape)
                Me(index) = Value
            End Set
        End Property


        Public Function Add(ByVal value As PolygonShape) As Integer
            Return List.Add(value)
        End Function 'Add
        ''Add
        Public Function IndexOf(ByVal value As PolygonShape) As Integer
            Return Me.IndexOf(value)
        End Function 'IndexOf
        ''IndexOf
        Public Sub Insert(ByVal index As Integer, ByVal value As PolygonShape)
            Me.Insert(index, value)
        End Sub 'Insert
        ''Insert
        Public Sub Remove(ByVal value As PolygonShape)
            Me.Remove(value)
        End Sub 'Remove
        ''Remove
        Public Function Contains(ByVal value As PolygonShape) As Boolean
            '' If value is not of type PolygonShape, this will return false.
            Return Me.Contains(value)
        End Function 'Contains ''Contains
    End Class 'PolygonShapeCollection
End Namespace 'ChartSamplesExplorerCS.Customization ''PolygonShapeCollection