Public Class WHCBASE0
    Inherits ASCBASE0

    Sub New(e As GunEnvironment)
        MyBase.New(e)

    End Sub
 
    Public Sub Dispose()
        If ASCMAIN1 IsNot Nothing Then
            Me.ASCMAIN1.MultiTask_Release()
        End If
        Me.clsASCBASE1.Dispose()
        Me.ASCDATA1.Dispose()
    End Sub
End Class