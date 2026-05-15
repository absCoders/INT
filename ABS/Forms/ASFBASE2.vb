Public Class ASFBASE2
    Public frmASFBASE1 As ASFBASE1

    Sub New()

        MENU_ID = "XX"
        MENU_ITEM_TYPE = "F"
        Dim XX As String = Split(Replace(Me.ToString, ",", "."), ".")(1)

        MENU_ITEM_OBJECT = XX
        MENU_ITEM_DESC = Me.Text
        MENU_ITEM_SECURITY = ""
        MENU_ITEM_PP = ""
        MENU_ITEM_FORM = ""
        MODULE_ID = Mid(XX, 1, 2)
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ASFBASE2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ASCMAIN1.USER_ID <> "" Then
            ASCMAIN1.Center(Me)
        End If
    End Sub

    Private Sub ASFBASE2_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub
End Class