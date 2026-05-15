Public Class TAFCRED1
    Public CUST_CODE As String
    Public rowSOTINVH1 As DataRow
    Public STATUS As String

    Public Sub New(ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & ""
        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & ""

        ASCMAIN1.sql = "Select BT.*, ST.* from " _
        & " (Select ORDR_NO, CUST_NAME, CUST_CITY, CUST_STATE, CUST_ZIP_CODE from SOTORDR5 where ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'BT') BT" _
        & ",(Select CUST_NAME CUST_SHIP_TO_NAME, CUST_CITY CUST_SHIP_TO_CITY, CUST_STATE CUST_SHIP_TO_STATE, CUST_ZIP_CODE CUST_SHIP_TO_ZIP_CODE from SOTORDR5 where ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'BT') ST"
        Create_TDA(dst.Tables.Add, "SOTORDRA", "**", 0, False, "", 1)

        Create_TDA(dst.Tables.Add, "SOTINVH1", "*", , False)
        Dim X As Int32 = dst.Tables("SOTORDRA").Rows.Count
        Fill_Records("SOTORDRA")
        Fill_Records("SOTINVH1", New String() {"I", INV_NO})

        Set_Read_Only(grpCustomer, True)
        Set_Read_Only(grpInvoice, True)

        Bind_Controls(Me, "SOTORDRA")
        TABLE_NAME = "SOTINVH1"
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
        End Select
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub cmdSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        If Not ASCMAIN1.Running_in_VS Then
            Exit Sub
        End If

        STATUS = "S"
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        STATUS = "C"
        Me.Close()
    End Sub
End Class