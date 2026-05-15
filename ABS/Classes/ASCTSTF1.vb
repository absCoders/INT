Public Class ASCTSTF1


#Region "Class Variables"
    Private _InventoryAllocation As Boolean = False
#End Region

#Region "Properties"
    ''' <summary>
    ''' Turn Inventory Allocation features on or off for testing
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property InventoryAllocation() As Boolean
        Get
            Return _InventoryAllocation
        End Get

        Set(ByVal value As Boolean)
            _InventoryAllocation = value
        End Set
    End Property
#End Region

End Class
