Public Class TATCURR1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select TATCURR2.*, GLTPARM2.LEGEND" & vbCrLf _
                & " from TATCURR2,GLTPARM2" & vbCrLf _
                & " where GLTPARM2.OPS_YYYYPP (+) = TATCURR2.OPS_YYYYPP" & vbCrLf _
                & "  and TATCURR2.CURR_CODE = :PARM1" & vbCrLf _
                & "  and TATCURR2.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "'"
            Create_TDA(.Tables.Add, "TATCURR2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select TATCURR3.*" & vbCrLf _
                & " from TATCURR3" & vbCrLf _
                & " where TATCURR3.CURR_CODE = :PARM1 " & vbCrLf _
                & "  and TATCURR3.CURR_DATE >= SYSDATE - 365 * 2"
            Create_TDA(.Tables.Add, "TATCURR3", "**", 0, True, "V", 2)

        End With

        grdTATCURR2.DataSource = dst.Tables("TATCURR2")
        grdTATCURR3.DataSource = dst.Tables("TATCURR3")

        Dim daily_rates As Boolean = True

        ' numCURR_EXCH_FUT.Visible = Not daily_rates
        numCURR_EXCH_CUR.Visible = Not daily_rates
        ' lblCURR_EXCH_FUT.Visible = Not daily_rates
        lblCURR_EXCH_CUR.Visible = Not daily_rates

        If daily_rates Then
            '  splRates.Panel1Collapsed = True
        Else
            splRates.Panel2Collapsed = True
        End If

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        '   Load_Popup_Menu(grdTATCURR2, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        ' Dim sqlDelete = " = '" & Absx1.txtFor("PRICE_LIST_CODE").Text & "'"
        If ASCMAIN1.CLIENT = "AHA" Then
            WriteAuditTrail("TATCURR3")

            'For Each row As DataRow In dst.Tables("TATCURR3").Rows
            '    If row.RowState = DataRowState.Unchanged Then
            '    Else
            '        Dim m As String = IIf(row.RowState = DataRowState.Deleted, "D", IIf(row.RowState = DataRowState.Modified, "E", "?"))
            '        Write_Audit_Trail(row, m)
            '    End If
            'Next
            Update_Record_TDA("TATCURR3")
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

        If EntryMode = "Edit" Then
            TAC.TACMAIN1.Update_Forex()
        End If

        EnforceConstraints(False)
        Fill_Records("TATCURR2", New String() {Absx1.txtFor("CURR_CODE").Text})
        Sort_grdColumns(grdTATCURR2, "OPS_YYYYPP".ToLower)
        Fill_Records("TATCURR3", New String() {Absx1.txtFor("CURR_CODE").Text})
        Sort_grdColumns(grdTATCURR3, "CURR_DATE".ToLower)
        EnforceConstraints(True)

        grdTATCURR2.Text = "Monthly Rate of Exchange History for " & Absx1.txtFor("CURR_CODE").Text
        grdTATCURR3.Text = "Daily Rate of Exchange History for " & Absx1.txtFor("CURR_CODE").Text
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("TATCURR2").Rows.Clear()
            dst.Tables("TATCURR3").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdTATCURR2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        grdTATCURR2.Visible = tf
        grdTATCURR3.Visible = tf


        If ASCMAIN1.CLIENT = "AHA" Then
            With grdTATCURR3.DisplayLayout.Override
                If EntryMode = "Edit" Then
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        End If

    End Sub

    Public Overrides Function Remote_Control( _
ByVal command As String, _
Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "View"
                If key <> "" Then
                    Absx1.txtFor("CURR_CODE").Text = key
                    Click_Command(command)
                End If
        End Select

        Return return_key
    End Function

#End Region
     
      
End Class