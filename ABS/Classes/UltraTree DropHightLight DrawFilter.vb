Imports Infragistics.Win.UltraWinTree

'Enumerates the possiblt Drop States
<Flags()> Public Enum DropLinePositionEnum
    None = 0
    OnNode = 1
    AboveNode = 2
    BelowNode = 4
    All = OnNode Or AboveNode Or BelowNode
End Enum

Public Class UltraTree_DropHightLight_DrawFilter_Class
    'This class has to implement the DrawFilter Interface
    'so it can be used as a DrawFilter by the tree
    Implements Infragistics.Win.IUIElementDrawFilter


    Public Event Invalidate(ByVal sender As Object, ByVal e As EventArgs)
    Public Event QueryStateAllowedForNode(ByVal sender As Object, ByVal e As QueryStateAllowedForNodeEventArgs)

    'Used by the QueryStateAllowedForNode event
    Public Class QueryStateAllowedForNodeEventArgs
        Inherits EventArgs

        Public Node As UltraTreeNode
        Public DropLinePosition As DropLinePositionEnum
        Public StatesAllowed As DropLinePositionEnum
    End Class

    Public Sub New()
        'Initialize the properties to the defaults
        InitProperties()
    End Sub

    'Initialize the properties to the defaults
    Private Sub InitProperties()
        mvarDropHighLightNode = Nothing
        mvarDropLinePosition = DropLinePositionEnum.None
        mvarDropHighLightBackColor = System.Drawing.SystemColors.Highlight
        mvarDropHighLightForeColor = System.Drawing.SystemColors.HighlightText
        mvarDropLineColor = System.Drawing.SystemColors.ControlText
        mvarEdgeSensitivity = 0
        mvarDropLineWidth = 2
    End Sub

    'Clean up
    Public Sub Dispose()
        mvarDropHighLightNode = Nothing
    End Sub

    'The DropHighLightNode is a reference to the node the
    'Mouse is currently over
    Private mvarDropHighLightNode As UltraTreeNode
    Public Property DropHightLightNode() As UltraTreeNode
        Get
            Return mvarDropHighLightNode
        End Get
        Set(ByVal Value As UltraTreeNode)
            'If the Node is being set to the same value,
            ' just exit
            If mvarDropHighLightNode.Equals(Value) Then Exit Property

            mvarDropHighLightNode = Value
            'The DropNode has changed.
            PositionChanged()
        End Set
    End Property

    Private mvarDropLinePosition As DropLinePositionEnum
    Public Property DropLinePosition() As DropLinePositionEnum
        Get
            Return mvarDropLinePosition
        End Get
        Set(ByVal Value As DropLinePositionEnum)
            'If the position is the same as it was, 
            'just exit
            If mvarDropLinePosition = Value Then Exit Property

            mvarDropLinePosition = Value
            'The Drop Position has changed
            PositionChanged()
        End Set
    End Property

    'The width of the DropLine
    Private mvarDropLineWidth As Integer
    Public Property DropLineWidth() As Integer
        Get
            Return mvarDropLineWidth
        End Get
        Set(ByVal Value As Integer)
            mvarDropLineWidth = Value
        End Set
    End Property

    'The BackColor of the DropHighLight node
    'This only affect the node when it is being dropped On. 
    'Not Above or Below. 
    Private mvarDropHighLightBackColor As Color
    Public Property DropHighLightBackColor() As Color
        Get
            Return mvarDropHighLightBackColor
        End Get
        Set(ByVal Value As Color)
            mvarDropHighLightBackColor = Value
        End Set
    End Property

    'The ForeColor of the DropHighLight node
    'This only affect the node when it is being dropped On. 
    'Not Above or Below. 
    Private mvarDropHighLightForeColor As Color
    Public Property DropHighLightForeColor() As Color
        Get
            Return mvarDropHighLightForeColor
        End Get
        Set(ByVal Value As Color)
            mvarDropHighLightForeColor = Value
        End Set
    End Property

    'The color of the DropLine
    Private mvarDropLineColor As Color
    Public Property DropLineColor() As Color
        Get
            Return mvarDropLineColor
        End Get
        Set(ByVal Value As Color)
            mvarDropLineColor = Value
        End Set
    End Property

    'Determines how close to the top or bottom edge of a node
    'the mouse must be to be consider dropping Above or Below
    'respectively. 
    'By default the top 1/3 of the node is Above, the bottom 1/3
    'is Below, and the middle is On. 
    Private mvarEdgeSensitivity As Integer
    Public Property EdgeSensitivity() As Integer
        Get
            Return mvarEdgeSensitivity
        End Get
        Set(ByVal Value As Integer)
            mvarEdgeSensitivity = Value
        End Set
    End Property

    'When the DropNode or DropPosition change, we fire the
    'Invalidate event to let the program know to invalidate
    'the Tree control. 
    'This is neccessary since the DrawFilter does not have a 
    'reference to the Tree Control (although it probably could)
    Private Sub PositionChanged()
        Dim e As EventArgs = EventArgs.Empty
        RaiseEvent Invalidate(Me, e)
    End Sub

    'Set the DropNode to Nothing and the position to None. This
    'Will clear whatever Drophighlight is in the tree
    Public Sub ClearDropHighlight()
        SetDropHighlightNode(Nothing, DropLinePositionEnum.None)
    End Sub

    'Call this proc every time the DragOver event of the 
    'Tree fires. 
    'Note that the point pass in MUST be in Tree coords, not
    'form coords
    Public Sub SetDropHighlightNode(ByVal Node As UltraTreeNode, ByVal PointInTreeCoords As Point)
        'The distance from the edge of the node used to 
        'determine whether to drop Above, Below, or On a node
        Dim DistanceFromEdge As Integer
        'The new DropLinePosition
        Dim NewDropLinePosition As DropLinePositionEnum

        DistanceFromEdge = mvarEdgeSensitivity
        'Check to see if DistanceFromEdge is 0
        If DistanceFromEdge = 0 Then
            'It is, so we use the default value - one third. 
            DistanceFromEdge = Node.Bounds.Height / 3
        End If

        'Determine which part of the node the point is in
        Select Case PointInTreeCoords.Y
            Case Is < (Node.Bounds.Top + DistanceFromEdge)
                'Point is in the top of the node
                NewDropLinePosition = DropLinePositionEnum.AboveNode
            Case Is > ((Node.Bounds.Bottom - DistanceFromEdge) - 1)
                'Point is in the bottom of the node
                NewDropLinePosition = DropLinePositionEnum.BelowNode
            Case Else
                'Point is in the middle of the node
                NewDropLinePosition = DropLinePositionEnum.OnNode
        End Select

        'Now that we have the new DropLinePosition, call the
        'real proc to get things rolling
        SetDropHighlightNode(Node, NewDropLinePosition)
    End Sub

    Private Sub SetDropHighlightNode(ByVal Node As UltraTreeNode, ByVal DropLinePosition As DropLinePositionEnum)
        'Use to store whether there have been any changes in 
        'DropNode or DropLinePosition
        Dim IsPositionChanged As Boolean

        Try
            If Not mvarDropHighLightNode Is Nothing Then
                'Check to see if the nodes are equal and if 
                'the dropline position are equal
                If mvarDropHighLightNode.Equals(Node) And mvarDropLinePosition = DropLinePosition Then
                    'They are both equal. Nothing has changed. 
                    IsPositionChanged = False
                Else
                    'One or both have changed
                    IsPositionChanged = True
                End If
            End If
        Catch ex As Exception
            'If we reach this code, it means mvarDropHighLightNode 
            'is nothing, so it could not be compared
            If mvarDropHighLightNode Is Nothing Then
                'Check to see if Node is nothing, so we'll know
                'if Node = mvarDropHighLightNode
                IsPositionChanged = Not (Node Is Nothing)
            End If
        End Try

        'Set both properties without calling PositionChanged
        mvarDropHighLightNode = Node
        mvarDropLinePosition = DropLinePosition

        'Check to see if the PositionChanged
        If IsPositionChanged Then
            'Position did change.
            PositionChanged()
        End If
    End Sub

    'Only need to trap for 2 phases:
    'AfterDrawElement: for drawing the DropLine
    'BeforeDrawElement: for drawing the DropHighlight
    Public Function GetPhasesToFilter(ByRef drawParams As Infragistics.Win.UIElementDrawParams) As Infragistics.Win.DrawPhase Implements Infragistics.Win.IUIElementDrawFilter.GetPhasesToFilter
        Return Infragistics.Win.DrawPhase.AfterDrawElement Or Infragistics.Win.DrawPhase.BeforeDrawElement
    End Function

    'The actual drawing code
    Public Function DrawElement(ByVal drawPhase As Infragistics.Win.DrawPhase, ByRef drawParams As Infragistics.Win.UIElementDrawParams) As Boolean Implements Infragistics.Win.IUIElementDrawFilter.DrawElement
        Dim aUIElement As Infragistics.Win.UIElement
        Dim g As Graphics
        Dim aNode As UltraTreeNode

        'If there's no DropHighlight node or no position
        'specified, then we don't need to draw anything special. 
        'Just exit Function
        If mvarDropHighLightNode Is Nothing Or mvarDropLinePosition = DropLinePositionEnum.None Then Exit Function

        'Create a new QueryStateAllowedForNodeEventArgs object
        'to pass to the event
        Dim eArgs As New QueryStateAllowedForNodeEventArgs()

        'Initialize the object with the correct info
        eArgs.Node = mvarDropHighLightNode
        eArgs.DropLinePosition = Me.mvarDropLinePosition
        'Default to all states allowed. 
        eArgs.StatesAllowed = DropLinePositionEnum.All

        'Raise the event
        RaiseEvent QueryStateAllowedForNode(Me, eArgs)

        'Chek to see if the user allowed the current state
        'for this node. If not, exit function
        If (eArgs.StatesAllowed And mvarDropLinePosition) <> mvarDropLinePosition Then Exit Function

        'Get the element being drawn
        aUIElement = drawParams.Element
        'Determine which drawing phase we are in. 
        Select Case drawPhase
            Case Infragistics.Win.DrawPhase.BeforeDrawElement
                'We are in BeforeDrawElement, so we are only concerned with 
                'drawing the OnNode state. 
                If (mvarDropLinePosition And DropLinePositionEnum.OnNode) = DropLinePositionEnum.OnNode Then
                    'Check to see if we are drawing a NodeTextUIElement
                    If TypeOf aUIElement Is Infragistics.Win.UltraWinTree.NodeTextUIElement Then
                        'Get a reference to the node that this
                        'NodeTextUIElement is associated with
                        aNode = aUIElement.GetContext(GetType(UltraTreeNode))

                        'See if this is the DropHighlightNode
                        If aNode.Equals(mvarDropHighLightNode) Then
                            'Set the ForeColor and Backcolor of the node 
                            'to the DropHighlight colors 
                            'Note that AppearanceData only affects the
                            'node for this one paint. It will not
                            'change any properties of the node
                            drawParams.AppearanceData.BackColor = mvarDropHighLightBackColor
                            drawParams.AppearanceData.ForeColor = mvarDropHighLightForeColor
                        End If
                    End If
                End If
            Case Infragistics.Win.DrawPhase.AfterDrawElement
                'We're in AfterDrawElement
                'So the only states we are conderned with are
                'Below and Above
                'Check to see if we are drawing the Tree Element
                If TypeOf aUIElement Is Infragistics.Win.UltraWinTree.UltraTreeUIElement Then
                    'Declare a pen to us for drawing Droplines
                    Dim p As New Pen(mvarDropLineColor, mvarDropLineWidth)
                    'Get a reference to the Graphics object
                    'we are drawing to. 
                    g = drawParams.Graphics

                    'Get the NodeSelectableAreaUIElement for the 
                    'current DropNode. We will use this for
                    'positioning and sizing the DropLine
                    Dim tElement As Infragistics.Win.UltraWinTree.NodeSelectableAreaUIElement
                    tElement = drawParams.Element.GetDescendant(GetType(Infragistics.Win.UltraWinTree.NodeSelectableAreaUIElement), mvarDropHighLightNode)

                    'The left edge of the DropLine
                    Dim LeftEdge As Integer = tElement.Rect.Left - 4

                    'We need a reference to the control to 
                    'determine the right edge of the line
                    Dim aTree As UltraTree
                    aTree = tElement.GetContext(GetType(UltraTree))
                    Dim RightEdge As Integer = aTree.DisplayRectangle.Right - 4

                    'Used to store the Vertical position of the 
                    'DropLine
                    Dim LineVPosition As Integer

                    If (mvarDropLinePosition And DropLinePositionEnum.AboveNode) = DropLinePositionEnum.AboveNode Then
                        'Draw line above node
                        LineVPosition = mvarDropHighLightNode.Bounds.Top
                        g.DrawLine(p, LeftEdge, LineVPosition, RightEdge, LineVPosition)
                        p.Width = 1
                        g.DrawLine(p, LeftEdge, LineVPosition - 3, LeftEdge, LineVPosition + 2)
                        g.DrawLine(p, LeftEdge + 1, LineVPosition - 2, LeftEdge + 1, LineVPosition + 1)
                        g.DrawLine(p, RightEdge, LineVPosition - 3, RightEdge, LineVPosition + 2)
                        g.DrawLine(p, RightEdge - 1, LineVPosition - 2, RightEdge - 1, LineVPosition + 1)
                    End If
                    If (mvarDropLinePosition And DropLinePositionEnum.BelowNode) = DropLinePositionEnum.BelowNode Then
                        'Draw Line below node
                        LineVPosition = mvarDropHighLightNode.Bounds.Bottom
                        g.DrawLine(p, LeftEdge, LineVPosition, RightEdge, LineVPosition)
                        p.Width = 1
                        g.DrawLine(p, LeftEdge, LineVPosition - 3, LeftEdge, LineVPosition + 2)
                        g.DrawLine(p, LeftEdge + 1, LineVPosition - 2, LeftEdge + 1, LineVPosition + 1)
                        g.DrawLine(p, RightEdge, LineVPosition - 3, RightEdge, LineVPosition + 2)
                        g.DrawLine(p, RightEdge - 1, LineVPosition - 2, RightEdge - 1, LineVPosition + 1)
                    End If
                End If
        End Select
    End Function

End Class