Public Class SPFCREC1

    Dim SPTCRECX As String

    Dim RYP As String = ""
    Dim LYP As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Get_PARM("SPTPARM1")

        With dst

            Create_Work_Table("", "")

            ASCMAIN1.sql = "" _
            & "Select X.*, X.BEG + X.ACC - X.PMT CALC, X.BEG + X.ACC - X.PMT - X.END DIFF from (" & vbCrLf _
            & "Select CUST_CODE, ASP_CODE, SUM (BEG) BEG, SUM (ACC) ACC, SUM (PMT) PMT, SUM (END) END" & vbCrLf _
            & " from " & SPTCRECX & " group by CUST_CODE, ASP_CODE) X"
            Create_TDA(.Tables.Add, "SPTCREC1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SPTCRECX.*" & vbCrLf _
                & ", SPTACOMC.QTY_SOLD, SPTACOMC.AMT_SOLD, SPTACOMC.AMT_COMM, SPTACOMC.AMT_COMM_ADJ" & vbCrLf _
                & ", SPTACOMC.AMT_COMM_OFFSET, SPTACOMC.AMT_COMM_PAID" & vbCrLf _
                & ", SPTACOMC.INV_NO, SPTACOMC.ACC_CTL_NO_ACCRUAL, SPTACOMC.PYMT_NO" & vbCrLf _
                & ", SPTAPMT1.PYMT_DATE, SPTAPMT1.PYMT_REF_NO" & vbCrLf _
                & ", SPTAPMT1.PYMT_BATCH_NO, SPTAPMT1.PYMT_NOTES, SPTAPMT1.PYMT_CTL_NO" & vbCrLf _
                & "  from " & SPTCRECX & " SPTCRECX,SPTACOMC,SPTAPMT1" & vbCrLf _
                & " where SPTCRECX.CUST_CODE = :PARM1 and SPTCRECX.ASP_CODE = :PARM2" & vbCrLf _
                & "   and SPTACOMC.ACC_CTL_NO = SPTCRECX.ACC_CTL_NO and SPTAPMT1.PYMT_NO (+) = SPTACOMC.PYMT_NO"
            Create_TDA(.Tables.Add, "SPTCREC2", "**", 0, False, "VV", 3)

            'With .Tables("ICTGLICC").Columns
            '    .Add("OOBAL", GetType(System.Int64), _
            '        "ISNULL(BEG_BAL,0)-ISNULL(SHP,0)+ISNULL(RTN,0)+ISNULL(REC,0)+ISNULL(ADJ,0)-ISNULL(CON,0)-ISNULL(END_BAL,0)")

            '    .Add("BEG_BALC", GetType(System.Decimal), "ISNULL(BEG_BAL,0)*ISNULL(ITEM_COST_TOTAL_PREV,0)")
            '    For Each COLUMN_NAME As String In New String() {"SHP", "RTN", "REC", "ADJ", "CON", "END_BAL"}
            '        .Add(COLUMN_NAME & "C", GetType(System.Decimal), "ISNULL(" & COLUMN_NAME & ",0)*ISNULL(ITEM_COST_TOTAL_CURR,0)")
            '    Next

            '    .Add("OOBALC", GetType(System.Decimal), _
            '        "ISNULL(BEG_BALC,0)+ISNULL(REVC,0)-ISNULL(SHPC,0)+ISNULL(RTNC,0)+ISNULL(RECC,0)+ISNULL(ADJC,0)-ISNULL(CONC,0)-ISNULL(END_BALC,0)")
            'End With

        End With

        grdSPTCREC1.DataSource = dst.Tables("SPTCREC1")
        grdSPTCREC2.DataSource = dst.Tables("SPTCREC2")
        With grdSPTCREC1.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"CUST_CODE", "ASP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                    gcol.Width = 100
                    gcol.Format = "#,##0.00"
                End If
                If New String() {"BEG", "ACC", "PMT", "END"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    gcol.Width = 100
                    gcol.Format = "#,##0.00"
                End If
                If New String() {"CALC", "DIFF"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    gcol.Width = 100
                    gcol.Format = "#,##0.00"
                End If
            Next
        End With

        Create_Summary(grdSPTCREC1, "CUST_CODE", "Count")
        Create_Summary(grdSPTCREC1, New String() {"BEG", "ACC", "PMT", "END", "CALC", "DIFF"})
        Create_Summary(grdSPTCREC2, "ACC_CTL_NO", "Count")
        Create_Summary(grdSPTCREC2, New String() {"BEG", "ACC", "PMT", "END", "CALC", "DIFF"})

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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

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
        For Each TABLE_NAME As String In New String() {"SPTCREC1"}
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
        Create_Work_Table(LYP, RYP)
        Load_SPTCREC1()

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

        Load_Popup_Menu(grdSPTCREC1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTCREC2, "BB", "Show Payment")
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

            Case "Show Payment"
                Dim PYMT_NO As String = grd.ActiveRow.Cells("PYMT_NO").Text
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                If PYMT_NO <> "" Then
                    Context_Launch("View", CUST_CODE & ":" & PYMT_NO, e.Tool.Key, "SPFAPMT1")
                End If
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
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Click_Command("Load")
        End Select
    End Sub

#End Region

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

    End Sub

    Sub Load_SPTCREC1()
        Fill_Records("SPTCREC1")
        Sort_grdColumns(grdSPTCREC1, "CUST_CODE,ASP_CODE")
    End Sub

    Sub Setup_SPTCREC2()
        Dim CUST_CODE As String = grdSPTCREC1.ActiveRow.Cells("CUST_CODE").Value
        Dim ASP_CODE As String = grdSPTCREC1.ActiveRow.Cells("ASP_CODE").Value
        Fill_Records("SPTCREC2", New String() {CUST_CODE, ASP_CODE})
        Sort_grdColumns(grdSPTCREC2, "ACC_CTL_NO")
    End Sub
    Private Sub grdSPTCREC1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTCREC1.AfterRowActivate
        SETUP_SPTCREC2()
    End Sub

    Private Sub grdSPTCREC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCREC1.InitializeRow

        Dim DIFF As Int64 = Val(e.Row.Cells("DIFF").Value & "")

        If DIFF <> 0 Then
            e.Row.Cells("DIFF").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("DIFF").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Create_Work_Table(LYP, RYP)
        ASCMAIN1.sql = "" _
            & "Select X.*, X.BEG + X.ACC - X.PMT CALC, X.BEG + X.ACC - X.PMT - X.END DIFF from (" & vbCrLf _
            & "Select CUST_CODE, ASP_CODE, ACC_CTL_NO, SUM (BEG) BEG, SUM (ACC) ACC, SUM (PMT) PMT, SUM (END) END from (" & vbCrLf _
            & "Select GLTCREC3.DETL_CVX_NO CUST_CODE, SPTACOMC.ASP_CODE, GLTCREC3.DETL_CTL_NO ACC_CTL_NO, SUM (GLTCREC3.CREC_AMT) BEG, 0 ACC, 0 PMT, 0 END" & vbCrLf _
            & " from GLTCREC3,SPTACOMC" & vbCrLf _
            & " where GLTCREC3.CREC_TYPE_CODE = 'AC' AND GLTCREC3.OPS_YYYYPP = '{0}'" & vbCrLf _
            & "   and SPTACOMC.ACC_CTL_NO = GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & " group by GLTCREC3.DETL_CVX_NO, SPTACOMC.ASP_CODE, GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select CUST_CODE, ASP_CODE, ACC_CTL_NO, 0 BEG, SUM (AMT_COMM) ACC, 0 PMT, 0 END" & vbCrLf _
            & " from SPTACOMC where OPS_YYYYPP = '{1}' and ACC_CTL_NO_ORIG IS NULL" & vbCrLf _
            & " group by CUST_CODE, ASP_CODE, ACC_CTL_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select CUST_CODE, ASP_CODE, ACC_CTL_NO, 0 BEG, 0 ACC, SUM (AMT_COMM_OFFSET) PMT, 0 END" & vbCrLf _
            & " from SPTACOMC where OPS_YYYYPP_PAID = '{1}'" & vbCrLf _
            & " group by CUST_CODE, ASP_CODE, ACC_CTL_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select GLTCREC3.DETL_CVX_NO CUST_CODE, SPTACOMC.ASP_CODE, GLTCREC3.DETL_CTL_NO ACC_CTL_NO, 0 BEG, 0 ACC, 0 PMT, SUM (GLTCREC3.CREC_AMT) END" & vbCrLf _
            & " from GLTCREC3,SPTACOMC where GLTCREC3.CREC_TYPE_CODE = 'AC' and GLTCREC3.OPS_YYYYPP = '{1}'" & vbCrLf _
            & "   and SPTACOMC.ACC_CTL_NO = GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & " group by GLTCREC3.DETL_CVX_NO, SPTACOMC.ASP_CODE, GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & ") group by CUST_CODE, ASP_CODE, ACC_CTL_NO) X"
        ASCMAIN1.sql = String.Format(ASCMAIN1.sql, LYP, RYP)

        If SPTCRECX = "" Then
            SPTCRECX = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCRECX & " Add Primary Key (CUST_CODE, ASP_CODE, ACC_CTL_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SPTCRECX)
            ASCDATA1.ExecuteSQL("Insert into " & SPTCRECX & " " & ASCMAIN1.sql)
        End If
    End Sub
End Class