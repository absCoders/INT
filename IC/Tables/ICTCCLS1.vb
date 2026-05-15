Public Class ICTCCLS1

#Region "Overrides"

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME

            'Case "ITEM_SUB_VARIETY_CODE"
            '    sql_where = "ITEM_VARIETY_CODE = '" & MyBase.Absx1.txtFor("ITEM_VARIETY_CODE").Text.Trim & "'"
        End Select
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Edit"

            Case "Update"

                Dim COST_BASIS As String = Absx1.optFor("COST_BASIS").Value & ""

                If COST_BASIS = "R" Then
                    Dim COST_BASE_PCT_OF_MSRP As Decimal = Val(Absx1.numFor("COST_BASE_PCT_OF_MSRP").Value & "")
                    Dim COST_BASE_PCT_OF_MSRP_FUT As Decimal = Val(Absx1.numFor("COST_BASE_PCT_OF_MSRP_FUT").Value & "")

                    If COST_BASE_PCT_OF_MSRP <= 0 Or COST_BASE_PCT_OF_MSRP <= 0 Then
                        EMsg &= vbCr & "% of Retail may not be 0 or negative"
                    End If
                End If
        End Select

    End Sub


    Overrides Sub Proceed_Update_Special_Pre()
        If EntryMode = "New" Then
            rowASFBASE1.Item("COST_BASE_PCT_OF_MSRP_FUT") = rowASFBASE1.Item("COST_BASE_PCT_OF_MSRP")

        End If
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        If EntryMode = "New" Then
            rowASFBASE1.Item("COST_BASIS") = "R"
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        'If ScreenMode Then
        '    EnforceConstraints(False)
        '    dst.Tables("ICTRETLC").Rows.Clear()
        '    EnforceConstraints(True)
        'End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        Set_Read_Only_for_ctl(Absx1.optFor("COST_BASIS"), EntryMode <> "New")
        Set_Read_Only_for_ctl(Absx1.numFor("COST_BASE_PCT_OF_MSRP"), Not tf Or EntryMode <> "New")
        Set_Read_Only_for_ctl(Absx1.numFor("COST_BASE_PCT_OF_MSRP_FUT"), Not tf Or EntryMode <> "Edit")

        lblPCT_OF_MSRP.Visible = (Not tf Or EntryMode = "Edit" Or EntryMode = "View")
        Absx1.numFor("COST_BASE_PCT_OF_MSRP_FUT").Visible = (Not tf Or EntryMode = "Edit" Or EntryMode = "View")
        lblCOST_BASE_PCT_OF_MSRP_FUT.Visible = (Not tf Or EntryMode = "Edit" Or EntryMode = "View")

        'If EntryMode = "Edit" Then
        '    If ASCMAIN1.USER_SECURITY_CODEs.Contains("RT") Then
        '    Else
        '        Set_Read_Only_for_ctl(Absx1.numFor("ITEM_RETAIL_PRICE"), True)
        '    End If

        '    Set_Read_Only_for_ctl(Absx1.numFor("ITEM_NEW_RETAIL_PRICE"), True)
        '    Set_Read_Only_for_ctl(Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE"), True)
        'End If
    End Sub

#End Region


#Region "ABSColumn Controls"

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "COST_BASIS"
                If Not Me.IsLoading Then
                    grpPCT_OF_MSRP.Visible = (Absx1.optFor("COST_BASIS").Value & "" = "R")
                End If
        End Select
    End Sub

    Public Overrides Sub chk_CheckedChanged(sender As Object, e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case ""

        End Select
    End Sub

#End Region

End Class