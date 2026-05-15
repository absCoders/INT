Public Class ICTTYPE1

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
        ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_PLAN_WASTE_PCT = " & Val(rowASFBASE1.Item("ITEM_WASTE_PCT") & "") & " where ITEM_TYPE_CODE = '" & rowASFBASE1.Item("ITEM_TYPE_CODE") & "'"
        ASCDATA1.ExecuteSQL()
    End Sub
End Class