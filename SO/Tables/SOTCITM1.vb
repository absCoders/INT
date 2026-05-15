Public Class SOTCITM1

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                If EntryMode = "New" Then
                    ASCMAIN1.sql = "Select * from SOTCITM1 where CUST_CODE = :PARM1 and ITEM_CODE = :PARM2"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("ITEM_CODE").Text})
                    If row IsNot Nothing Then
                        EMsg &= vbCr & "Record already on File for " & Absx1.txtFor("CUST_CODE").Text & " with Item Code " & Absx1.txtFor("ITEM_CODE").Text & vbCr & " - see Customer Item Code " & row.Item("CUST_ITEM_CODE")
                        'MsgBox("Record already on File for " & Absx1.txtFor("CUST_CODE").Text & " with Item Code " & Absx1.txtFor("ITEM_CODE").Text & ": " & row.Item("CUST_ITEM_CODE"), MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                        'Exit Sub
                    End If
                End If
        End Select
    End Sub
     
     
#End Region
End Class