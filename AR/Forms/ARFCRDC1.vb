Public Class ARFCRDC1


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("ARTPARM1")

        With dst

            ASCMAIN1.sql = "SELECT ARTCCPA1.*, ARTCUST1.NO_CRC FROM ARTCCPA1, ARTCUST1 WHERE ARTCCPA1.CUST_CODE = ARTCUST1.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTCCPA1", "**", 0, False, String.Empty, 1)

        End With

        Show_Filter(grdARTCRDC1)

        grdARTCRDC1.DataSource = dst.Tables("ARTCCPA1")

        Create_Summary(grdARTCRDC1, "CCPA_NO", "Count")
        Create_Summary(grdARTCRDC1, "CCPA_AMT")

        Create_Lookup("GLTACCT1")

        ASCMAIN1.Add_Value_List(grdARTCRDC1, "CCPA_STATUS")
        ASCMAIN1.Add_Value_List(grdARTCRDC1, "CCPA_REASON")
        ASCMAIN1.Add_Value_List(grdARTCRDC1, "RESPONSE_CODE")
        ASCMAIN1.Add_Value_List(grdARTCRDC1, "CCPA_TYPE")

        grpPeriod.Top = grpDate.Top
        grpPeriod.Left = grpDate.Left

        optFilter.CheckedIndex = 0

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Select Case optFilter.Value

                    Case "P"
                        Validate_Code("OPS_YYYYPP")

                    Case "D"
                        If Not IsDate(dteFrom.Value) OrElse Not IsDate(dteTo.Value) Then
                            EMsg = "The From and To dates are required."
                        ElseIf dteFrom.Value > dteTo.Value Then
                            EMsg = "The From date must be less equal the To date."
                        End If

                    Case Else
                        EMsg = "Invalid Filter"

                End Select

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdARTCRDC1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ARTCCPA1").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP

        grdARTCRDC1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data")
        Call Save_Header_Fields(UltraGroupBox1)

        Dim dateField As String = String.Empty


        Dim sql As String = "SELECT ARTCCPA1.*,ARTCUST1.NO_CRC"
        sql &= " FROM ARTCCPA1, ARTCUST1"
        sql &= " WHERE ARTCCPA1.CUST_CODE = ARTCUST1.CUST_CODE"
        sql &= " and CCPA_STATUS <> 'X'"

        Select Case optFilter.Value
            Case "P"
                sql &= " AND ARTCCPA1.OPS_YYYYPP = '" & txtPeriod.Text & "'"
            Case "D"
                Select Case optDate.Value
                    Case "S"
                        dateField = "CCPA_DATE_SALE"
                    Case Else
                        dateField = "INIT_DATE"
                End Select

                sql &= " AND TRUNC(ARTCCPA1." & dateField & ") BETWEEN '" & CDate(dteFrom.Value).ToString("dd-MMM-yyyy") & "' AND '" & CDate(dteTo.Value).ToString("dd-MMM-yyyy") & "'"
        End Select


        Fill_Records("ARTCCPA1", String.Empty, True, sql)

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()


        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdARTGLARR, "SSSSS" _
        '                     , "Show Sales Order Types" _
        '                     , "Show Misc Charge Codes" _
        '                     , "Show Pymt Sources" _
        '                     , "Show Deduction Codes" _
        '                     , "Show GL Columns")

        Call Load_Popup_Menu(grdARTCRDC1, "SS" _
                     , "Show Filter" _
                     , "Show GroupBox")

        'Call Load_Popup_Menu(grdARTGLAR2, "SS" _
        '             , "Show Filter" _
        '             , "Show GroupBox")

        'Call Load_Popup_Menu(grdARTGLARC, "SS" _
        '     , "Show Filter" _
        '     , "Show GroupBox")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If tlb_pop.Tools.Exists("Show Sales Order Types") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Sales Order Types"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Misc Charge Codes") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Misc Charge Codes"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Pymt Sources") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Pymt Sources"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Deduction Codes") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Deduction Codes"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show GL Columns") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GL Columns"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
            '    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If ScreenMode Then
            Exit Sub
        End If

        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Call Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Call Click_Command("Load")
        End Select
    End Sub

#End Region


    Private Sub optFilter_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFilter.ValueChanged

        grpPeriod.Visible = optFilter.Value = "P"
        grpDate.Visible = optFilter.Value = "D"

    End Sub
End Class