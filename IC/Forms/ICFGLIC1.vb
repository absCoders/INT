Public Class ICFGLIC1

    Dim ICTGLICC As String

    Dim RYP As String = ""
    Dim LYP As String = ""

    Dim EVENT_CODEs As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")

        With dst

            Dim sqlx As String = ""

            ' ICTGLICC

            ASCMAIN1.sql = "Select ITEM_CODE" _
                & ", WHSE_QTY_BEG BEG_BAL" _
                & ", WHSE_QTY_SHP SHP" _
                & ", WHSE_QTY_RTN RTN" _
                & ", WHSE_QTY_REC REC" _
                & ", WHSE_QTY_ADJ ADJ" _
                & ", WHSE_QTY_CON CON" _
                & ", WHSE_QTY_BEG END_BAL" _
                & ", .01 * WHSE_QTY_BEG REVC" _
                & " from ICTSTAT1 where rownum < 1"
            ICTGLICC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTGLICC & " Add Primary Key (ITEM_CODE)")

            ASCMAIN1.sql = "" _
            & "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.COLLECTION_CODE" _
            & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.PROD_CODE" _
            & ", ICTCOLL1.BRAND_CODE" _
            & ", ICTCOSTA_PREV.ITEM_COST_TOTAL ITEM_COST_TOTAL_PREV" _
            & ", ICTCOSTA_CURR.ITEM_COST_TOTAL ITEM_COST_TOTAL_CURR" _
            & ", ICTGLICC.BEG_BAL" _
            & ", ICTGLICC.SHP" _
            & ", ICTGLICC.RTN" _
            & ", ICTGLICC.REC" _
            & ", ICTGLICC.ADJ" _
            & ", ICTGLICC.CON" _
            & ", ICTGLICC.END_BAL" _
            & ", ICTGLICC.REVC" _
            & " from " & ICTGLICC & " ICTGLICC,ICTITEM1,ICTCOLL1,ICTCOSTA ICTCOSTA_PREV,ICTCOSTA ICTCOSTA_CURR" _
            & " where ICTITEM1.ITEM_CODE = ICTGLICC.ITEM_CODE" _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
            & "   and ICTCOSTA_PREV.OPS_YYYYPP = :PARM1" _
            & "   and ICTCOSTA_PREV.ITEM_CODE (+) = ICTGLICC.ITEM_CODE" _
            & "   and ICTCOSTA_CURR.OPS_YYYYPP = :PARM2" _
            & "   and ICTCOSTA_CURR.ITEM_CODE (+) = ICTGLICC.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTGLICC", "**", 0, False, "VV", 1)

            With .Tables("ICTGLICC").Columns
                .Add("OOBAL", GetType(System.Int64), _
                    "ISNULL(BEG_BAL,0)-ISNULL(SHP,0)+ISNULL(RTN,0)+ISNULL(REC,0)+ISNULL(ADJ,0)-ISNULL(CON,0)-ISNULL(END_BAL,0)")

                .Add("BEG_BALC", GetType(System.Decimal), "ISNULL(BEG_BAL,0)*ISNULL(ITEM_COST_TOTAL_PREV,0)")
                For Each COLUMN_NAME As String In New String() {"SHP", "RTN", "REC", "ADJ", "CON", "END_BAL"}
                    .Add(COLUMN_NAME & "C", GetType(System.Decimal), "ISNULL(" & COLUMN_NAME & ",0)*ISNULL(ITEM_COST_TOTAL_CURR,0)")
                Next

                .Add("OOBALC", GetType(System.Decimal), _
                    "ISNULL(BEG_BALC,0)+ISNULL(REVC,0)-ISNULL(SHPC,0)+ISNULL(RTNC,0)+ISNULL(RECC,0)+ISNULL(ADJC,0)-ISNULL(CONC,0)-ISNULL(END_BALC,0)")
            End With

            Create_TDA(.Tables.Add, "SOTEVNT1", "*", 0, False)
            Fill_Records("SOTEVNT1")
            Create_TDA(.Tables.Add, "ICTCOST1", "*", 0, False)
            Fill_Records("ICTCOST1")

            With .Tables.Add("SOTINVHE")
                .Columns.Add("OPS_YYYYPP")
                .Columns.Add("COST_CATGY_CODE")
                .Columns.Add("COST_CATGY_DESC")
                .PrimaryKey = New DataColumn() {dst.Tables("SOTINVHE").Columns("OPS_YYYYPP"), dst.Tables("SOTINVHE").Columns("COST_CATGY_CODE")}
                .Columns.Add("TOTAL", GetType(System.Decimal))
                .Columns.Add("EVT_", GetType(System.Decimal))
                .Columns("EVT_").DefaultValue = 0
                Dim T As String = "EVT_"
                For Each row As DataRow In dst.Tables("SOTEVNT1").Select("", "EVENT_CODE")
                    Dim EVENT_CODE As String = row.Item("EVENT_CODE")
                    EVENT_CODEs.Add(EVENT_CODE, "0")
                    Dim ACCT_CODE As String = row.Item("ACCT_CODE")
                    Dim C As String = "EVT_" & EVENT_CODE
                    .Columns.Add(C, GetType(System.Decimal))
                    .Columns(C).DefaultValue = 0
                    T &= "+" & C
                Next
                .Columns("TOTAL").Expression = T
            End With
        End With

        grdICTGLICC.DataSource = dst.Tables("ICTGLICC")
        grdSOTINVHE.DataSource = dst.Tables("SOTINVHE")

        With grdSOTINVHE.DisplayLayout.Bands(0)
            .Columns("OPS_YYYYPP").Header.Caption = "YP"
            .Columns("OPS_YYYYPP").Hidden = True
            .Columns("COST_CATGY_CODE").Header.Caption = "CC"
            .Columns("COST_CATGY_CODE").Width = 50
            .Columns("COST_CATGY_DESC").Header.Caption = "CC"
            .Columns("COST_CATGY_DESC").Width = 100
            .ColHeaderLines = 2
            .Columns("TOTAL").Header.Caption = "Total"
            Create_Summary(grdSOTINVHE, "TOTAL")
            .Columns("EVT_").Header.Caption = "Non-Event"
            Create_Summary(grdSOTINVHE, "EVT_")
            For Each rowSOTEVNT1 As DataRow In dst.Tables("SOTEVNT1").Select("")
                Dim EVENT_CODE As String = rowSOTEVNT1.Item("EVENT_CODE")
                Dim ACCT_CODE As String = rowSOTEVNT1.Item("ACCT_CODE")
                .Columns("EVT_" & EVENT_CODE).Header.Caption = EVENT_CODE & vbCrLf & ACCT_CODE
                .Columns("EVT_" & EVENT_CODE).Width = 100
                Create_Summary(grdSOTINVHE, "EVT_" & EVENT_CODE)
            Next
        End With


        With grdICTGLICC.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ITEM_COST_TOTAL_PREV", "ITEM_COST_TOTAL_CURR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                    gcol.Width = 70
                    gcol.Format = "###.0000"
                End If
                If New String() {"BEG_BAL", "SHP", "RTN", "REC", "ADJ", "CON", "END_BAL", "OOBAL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    gcol.Width = 70
                    gcol.Format = "#,##0"
                End If
                If New String() {"BEG_BALC", "REVC", "SHPC", "RTNC", "RECC", "ADJC", "CONC", "END_BALC", "OOBALC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
            Next
        End With

        Create_Summary(grdICTGLICC, "ITEM_CODE", "Count")
        Create_Summary(grdICTGLICC, New String() {"BEG_BAL", "SHP", "RTN", "REC", "ADJ", "CON", "END_BAL", "OOBAL"})
        Create_Summary(grdICTGLICC, New String() {"BEG_BALC", "REVC", "SHPC", "RTNC", "RECC", "ADJC", "CONC", "END_BALC", "OOBALC"})


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("OPS_YYYYPP")

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

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

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf
        Setup_tabMain()
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTGLICC", "SOTINVHE"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        RYP = HFs("OPS_YYYYPP")
        LYP = ASCMAIN1.Period_Calc(RYP, -1)

        Load_ICTGLICC()
        Load_SOTINVHE()

        Setup_tabMain()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdICTGLICC, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")

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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
 
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
            
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")

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

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

    End Sub

    Sub Load_SOTINVHE()
        ASCMAIN1.sql = "SELECT SOTINVH2.ORDR_YYYYPP_UPDATED, ICTITEM1.COST_CATGY_CODE, SOTINVH1.EVENT_CODE, SUM(ORDR_QTY_SHIP * ITEM_UNIT_COST) CGS" _
            & " from SOTINVH2,ICTITEM1,SOTINVH1" _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & RYP & "' AND SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'" _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
            & " group by SOTINVH2.ORDR_YYYYPP_UPDATED, ICTITEM1.COST_CATGY_CODE, SOTINVH1.EVENT_CODE"

        For i As Integer = 0 To EVENT_CODEs.Count - 1
            EVENT_CODEs(i) = "0"
        Next

        dst.Tables("SOTINVHE").Rows.Clear()
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim OPS_YYYYPP As String = row.Item("ORDR_YYYYPP_UPDATED")
            Dim EVENT_CODE As String = row.Item("EVENT_CODE") & ""
            EVENT_CODEs(EVENT_CODE) = "1"

            Dim COST_CATGY_CODE As String = row.Item("COST_CATGY_CODE")
            Dim CGS As String = Val(row.Item("CGS") & "")
            Dim rowSOTINVHE As DataRow = dst.Tables("SOTINVHE").Rows.Find(New String() {OPS_YYYYPP, COST_CATGY_CODE})
            If rowSOTINVHE Is Nothing Then
                rowSOTINVHE = dst.Tables("SOTINVHE").NewRow
                rowSOTINVHE.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowSOTINVHE.Item("COST_CATGY_CODE") = COST_CATGY_CODE
                Dim rowICTCOST1 As DataRow = dst.Tables("ICTCOST1").Rows.Find(COST_CATGY_CODE)
                If rowICTCOST1 IsNot Nothing Then
                    rowSOTINVHE.Item("COST_CATGY_DESC") = rowICTCOST1.Item("COST_CATGY_DESC")
                End If

                dst.Tables("SOTINVHE").Rows.Add(rowSOTINVHE)
            End If
            rowSOTINVHE.Item("EVT_" & EVENT_CODE) = Val(rowSOTINVHE.Item("EVT_" & EVENT_CODE) & "") + CGS
        Next

    End Sub

    Sub Load_ICTGLICC()

        ASCMAIN1.sql = "Truncate Table " & ICTGLICC
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ICTGLICC & vbCrLf _
            & "Select ITEM_CODE, Sum (BEG_BAL) BEG_BAL" _
            & ", Sum (SHP) SHP, Sum (RTN) RTN, Sum (REC) REC, Sum (ADJ) ADJ, Sum (CON) CON" & vbCrLf _
            & ", Sum (END_BAL) END_BAL, Sum (REVC) REVC" & vbCrLf _
            & " from (" _
            & "(Select ITEM_CODE" & vbCrLf _
            & ", Sum (WHSE_QTY_BEG) BEG_BAL" & vbCrLf _
            & ", Sum (WHSE_QTY_SHP) SHP" & vbCrLf _
            & ", Sum (WHSE_QTY_RTN) RTN" & vbCrLf _
            & ", Sum (WHSE_QTY_REC) REC" & vbCrLf _
            & ", Sum (WHSE_QTY_ADJ) ADJ" & vbCrLf _
            & ", Sum (WHSE_QTY_CON) CON" & vbCrLf _
            & ", 0 END_BAL" & vbCrLf _
            & ", 0 REVC" & vbCrLf _
            & "from ICTSTAT1 where OPS_YYYYPP = '" & RYP & "' group by ITEM_CODE)" & vbCrLf _
            & " union " & vbCrLf _
            & "(Select ITEM_CODE" & vbCrLf _
            & ", 0 BEG_BAL" & vbCrLf _
            & ", 0 SHP" & vbCrLf _
            & ", 0 RTN" & vbCrLf _
            & ", 0 REC" & vbCrLf _
            & ", 0 ADJ" & vbCrLf _
            & ", 0 CON" & vbCrLf _
            & ", Sum (WHSE_QTY_ON_HAND) END_BAL" & vbCrLf _
            & ", 0 REVC" & vbCrLf _
            & " from " _
            & IIf(RYP = ASCMAIN1.CYP, _
                  "ICTSTAT2", _
                  "ICTSTAT5 where OPS_YYYYPP = '" & RYP & "'") & vbCrLf _
            & " group by ITEM_CODE)" & vbCrLf _
            & " union " & vbCrLf _
            & "(Select ITEM_CODE" & vbCrLf _
            & ", 0 BEG_BAL" & vbCrLf _
            & ", 0 SHP" & vbCrLf _
            & ", 0 RTN" & vbCrLf _
            & ", 0 REC" & vbCrLf _
            & ", 0 ADJ" & vbCrLf _
            & ", 0 CON" & vbCrLf _
            & ", 0 END_BAL" & vbCrLf _
            & ", RV_EXP REVC" & vbCrLf _
            & " from ICTIVAR1 where OPS_YYYYPP = '" & RYP & "')" & vbCrLf _
            & ") group by ITEM_CODE" & vbCrLf

        ASCDATA1.ExecuteSQL()

        Fill_Records("ICTGLICC", New String() {LYP, RYP})
        Sort_grdColumns(grdICTGLICC, "ITEM_CODE")

    End Sub

    Private Sub grdICTGLICC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTGLICC.InitializeRow

        Dim OOBAL As Int64 = Val(e.Row.Cells("OOBAL").Value & "")
        Dim OOBALC As Decimal = Val(e.Row.Cells("OOBALC").Value & "")

        If OOBAL <> 0 Then
            e.Row.Cells("OOBAL").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("OOBAL").Appearance.ForeColor = Drawing.Color.Empty
        End If

        If OOBALC <> 0 Then
            e.Row.Cells("OOBALC").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("OOBALC").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub
End Class