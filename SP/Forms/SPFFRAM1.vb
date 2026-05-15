Imports Infragistics.Win.UltraWinGrid

Public Class SPFFRAM1

    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from SPTFRAM1"
            Create_TDA(.Tables.Add, "SPTFRAM1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select FRAME_MFG_CLEAN, COUNT (*) JOBS from SPTFRAM1 GROUP BY FRAME_MFG_CLEAN"
            Create_TDA(.Tables.Add, "SPTFRAMB", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, COUNT (*) JOBS from SPTFRAM1 WHERE FRAME_MFG_CLEAN = :PARM1 GROUP BY CUST_CODE, CUST_NAME"
            Create_TDA(.Tables.Add, "SPTFRAMBC", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select FRAME_MODEL_NO_CLEAN, MAX(INV_DATE) INV_DATE, COUNT (*) JOBS from SPTFRAM1 WHERE FRAME_MFG_CLEAN = :PARM1 and CUST_CODE = :PARM2 GROUP BY FRAME_MODEL_NO_CLEAN"
            'ASCMAIN1.sql = "Select * from SPTFRAM1 WHERE FRAME_MFG_CLEAN = :PARM1 and CUST_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SPTFRAMBCM", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, COUNT (*) JOBS from SPTFRAM1 GROUP BY CUST_CODE, CUST_NAME"
            Create_TDA(.Tables.Add, "SPTFRAMC", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select FRAME_MFG_CLEAN, COUNT (*) JOBS from SPTFRAM1 WHERE CUST_CODE = :PARM1 GROUP BY FRAME_MFG_CLEAN"
            Create_TDA(.Tables.Add, "SPTFRAMCB", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select FRAME_MODEL_NO_CLEAN, MAX(INV_DATE) INV_DATE, COUNT (*) JOBS from SPTFRAM1 WHERE FRAME_MFG_CLEAN = :PARM1 and CUST_CODE = :PARM2 GROUP BY FRAME_MODEL_NO_CLEAN"
            '  ASCMAIN1.sql = "Select * from SPTFRAM1 WHERE FRAME_MFG_CLEAN = :PARM1 and CUST_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SPTFRAMCBM", "**", 0, False, "VV", 0)


        End With
        grdSPTFRAM1.DataSource = dst.Tables("SPTFRAM1")
        Create_Summary(grdSPTFRAM1, "JOB_NO", "Count")
        '  Create_Summary(grdSPTFRAM1, "JOBS")

        grdSPTFRAMB.DataSource = dst.Tables("SPTFRAMB")
        Create_Summary(grdSPTFRAMB, "FRAME_MFG_CLEAN", "Count")
        Create_Summary(grdSPTFRAMB, "JOBS")

        grdSPTFRAMBC.DataSource = dst.Tables("SPTFRAMBC")
        Create_Summary(grdSPTFRAMBC, "CUST_CODE", "Count")
        Create_Summary(grdSPTFRAMBC, "JOBS")

        grdSPTFRAMBCM.DataSource = dst.Tables("SPTFRAMBCM")
        Create_Summary(grdSPTFRAMBCM, "FRAME_MODEL_NO_CLEAN", "Count")
        Create_Summary(grdSPTFRAMBCM, "JOBS")



        grdSPTFRAMC.DataSource = dst.Tables("SPTFRAMC")
        Create_Summary(grdSPTFRAMC, "CUST_CODE", "Count")
        Create_Summary(grdSPTFRAMC, "JOBS")

        grdSPTFRAMCB.DataSource = dst.Tables("SPTFRAMCB")
        Create_Summary(grdSPTFRAMCB, "FRAME_MFG_CLEAN", "Count")
        Create_Summary(grdSPTFRAMCB, "JOBS")

        grdSPTFRAMCBM.DataSource = dst.Tables("SPTFRAMCBM")
        Create_Summary(grdSPTFRAMCBM, "FRAME_MODEL_NO_CLEAN", "Count")
        Create_Summary(grdSPTFRAMCBM, "JOBS")

        spl.Panel1Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

            Case "Cancel"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
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
                With .Groups("Screen Control")

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not (EntryMode = "E")
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")

                End With
            End With
        End If

        UltraTabControl1.Visible = ScreenMode

        If ScreenMode Then

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl.Panel1Collapsed = ScreenMode
        splSPTATTRX.Visible = ScreenMode

        splSPTATTRX.Panel2Collapsed = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() _
            {"SPTFRAM1", "SPTFRAMB", "SPTFRAMBC", "SPTFRAMBCM", "SPTFRAMC", "SPTFRAMCB", "SPTFRAMCBM"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Fill_Records("SPTFRAM1")

        Fill_Records("SPTFRAMB")
        Sort_grdColumns(grdSPTFRAMB, "JOBS".ToLower)

        Fill_Records("SPTFRAMC")
        Sort_grdColumns(grdSPTFRAMC, "JOBS".ToLower)


        EnforceConstraints(True)
        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()


    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTFRAM1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMB, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMBC, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMBCM, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMC, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMCB, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTFRAMCBM, "SS", "Show Filter", "Show GroupBox")

    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then ' Or Not GRDs.ContainsKey(e.SourceControl.Name) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing ' GRDs(Mid(e.SourceControl.Name, 4))
        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            e.Cancel = True
            Exit Sub
        Else
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSPTATTRX"


            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim rowsSkipped As Int16 = 0

        Select Case e.Tool.Key

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE", "ATTR_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE", "ATTR_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                'If EntryMode = "" Then
                '    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                '        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                '        If cdr IsNot Nothing Then

                '        End If
                '    End If
                'End If
        End Select
    End Sub

    Private Sub grdSPTFRAMB_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTFRAMB.AfterRowActivate
        If grdSPTFRAMB.ActiveRow Is Nothing OrElse grdSPTFRAMB.ActiveRow.IsFilterRow Then Exit Sub

        Dim FRAME_MFG_CLEAN As String = grdSPTFRAMB.ActiveRow.Cells("FRAME_MFG_CLEAN").Value
        Fill_Records("SPTFRAMBC", FRAME_MFG_CLEAN)
        Sort_grdColumns(grdSPTFRAMBC, "JOBS".ToLower)
        grdSPTFRAMBC.Text = "Customers with Brand " & FRAME_MFG_CLEAN
    End Sub

    Private Sub grdSPTFRAMC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTFRAMC.AfterRowActivate
        If grdSPTFRAMC.ActiveRow Is Nothing OrElse grdSPTFRAMC.ActiveRow.IsFilterRow Then Exit Sub

        Dim CUST_CODE As String = grdSPTFRAMC.ActiveRow.Cells("CUST_CODE").Value
        Fill_Records("SPTFRAMCB", CUST_CODE)
        Sort_grdColumns(grdSPTFRAMCB, "JOBS".ToLower)
        grdSPTFRAMCB.Text = "Brands with Customer " & CUST_CODE
    End Sub


    Private Sub grdSPTFRAMBC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTFRAMBC.AfterRowActivate
        If grdSPTFRAMBC.ActiveRow Is Nothing OrElse grdSPTFRAMBC.ActiveRow.IsFilterRow Then Exit Sub

        Dim CUST_CODE As String = grdSPTFRAMBC.ActiveRow.Cells("CUST_CODE").Value
        Dim FRAME_MFG_CLEAN As String = grdSPTFRAMB.ActiveRow.Cells("FRAME_MFG_CLEAN").Value
        Fill_Records("SPTFRAMBCM", New String() {FRAME_MFG_CLEAN, CUST_CODE})
        '  Sort_grdColumns(grdSPTFRAMBCM, "JOBS".ToLower)
        grdSPTFRAMBCM.Text = FRAME_MFG_CLEAN & " Models ordered by Customer " & CUST_CODE

    End Sub

    Private Sub grdSPTFRAMCB_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTFRAMCB.AfterRowActivate
        If grdSPTFRAMCB.ActiveRow Is Nothing OrElse grdSPTFRAMCB.ActiveRow.IsFilterRow Then Exit Sub

        Dim CUST_CODE As String = grdSPTFRAMC.ActiveRow.Cells("CUST_CODE").Value
        Dim FRAME_MFG_CLEAN As String = grdSPTFRAMCB.ActiveRow.Cells("FRAME_MFG_CLEAN").Value
        Fill_Records("SPTFRAMCBM", New String() {FRAME_MFG_CLEAN, CUST_CODE})
        ' Sort_grdColumns(grdSPTFRAMCBM, "JOBS".ToLower)
        grdSPTFRAMCBM.Text = FRAME_MFG_CLEAN & " Models ordered by Customer " & CUST_CODE

    End Sub

#End Region


End Class