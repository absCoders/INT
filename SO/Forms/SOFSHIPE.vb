Public Class SOFSHIPE

    Private wkTable As String = String.Empty
    Private sqlEDT945T1 As String = String.Empty
    Private sqlARTCCPAC As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Dim sql As String = String.Empty

            sql = "Select EDT945T1.*, SOTORDR1.CUST_CODE CUST_CODE_S, SOTORDR1.CUST_NAME, SOTORDR1.FRT_TERMS, SOTORDR1.ORDR_CUST_PO"
            sql &= " FROM EDT945T1, SOTPICK1, SOTORDR1"
            sql &= " where EDT945T1.EDI_PICK_NO = SOTPICK1.PICK_NO (+)"
            sql &= " and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO (+)"
            sql &= " and NVL(EDT945T1.EDI_PROCESS_IND, '0') = '0'"
            sql &= " and SOTPICK1.PICK_STATUS <> 'D'"
            sql &= " and (EDT945T1.EDI_SHIPMENT_ID, EDT945T1.EDI_PICK_NO) IN (SELECT EDI_SHIPMENT_ID, MAX(EDI_PICK_NO) EDI_PICK_NO FROM EDT945T1 WHERE NVL(EDT945T1.EDI_PROCESS_IND, '0') = '0' GROUP BY EDI_SHIPMENT_ID)"
            sqlEDT945T1 = sql
            wkTable = ASCMAIN1.Temp_Table(sql & " and rownum < 1")

            Create_TDA(.Tables.Add, "EDT945T1", sql)

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")

            .Relations.Add("EDT945T1_SOTSHIP1", dst.Tables("EDT945T1").Columns("EDI_SHIPMENT_ID"), dst.Tables("SOTSHIP1").Columns("SHIP_BOL_NO"))
            .Relations.Add("SOTSHIP1_SOTPICK1", dst.Tables("SOTSHIP1").Columns("SHIP_BOL_NO"), dst.Tables("SOTPICK1").Columns("SHIP_BOL_NO"))

            sql = "SELECT ARTCCPAC.*, SOTINVH1.CUST_CODE, SOTINVH1.ORDR_CUST_PO, SOTINVH1.INV_DATE"
            sql &= " FROM ARTCCPAC, SOTINVH1"
            sql &= " where ARTCCPAC.INV_NO = SOTINVH1.INV_NO (+)"
            sqlARTCCPAC = sql
            Create_TDA(.Tables.Add, "ARTCCPAC", sql)
        End With

        grdEDT945T1.DataSource = dst.Tables("EDT945T1")
        grdARTCCPAC.DataSource = dst.Tables("ARTCCPAC")

        Create_Summary(grdEDT945T1, "CUST_CODE", "Count")
        Create_Summary(grdEDT945T1, "EDI_FRT_COST", "Sum")

        Create_Summary(grdARTCCPAC, "CUST_CODE", "Count")
        Create_Summary(grdARTCCPAC, "INV_TOTAL_AMOUNT", "Sum")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"


            Case "Done"

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
                EntryMode = "L"
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

        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"EDT945T1", "SOTSHIP1", "SOTPICK1", "ARTCCPAC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdEDT945T1)
        Clear_All_Filters(grdARTCCPAC)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & wkTable)
        ASCDATA1.ExecuteSQL("INSERT INTO " & wkTable & " " & sqlEDT945T1)

        Fill_Records("EDT945T1", String.Empty, True, "SELECT * from " & wkTable)
        Fill_Records("SOTSHIP1", String.Empty, True, "select * from SOTSHIP1 where ship_bol_no in (select edi_shipment_id from " & wkTable & ")")
        Fill_Records("SOTPICK1", String.Empty, True, "select * from SOTPICK1 where ship_bol_no in (select edi_shipment_id from " & wkTable & ")")

        Fill_Records("ARTCCPAC", String.Empty, True, sqlARTCCPAC)

        EnforceConstraints(True)

        Sort_grdColumns(grdEDT945T1, "CUST_CODE_S,EDI_SHIPMENT_ID")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT945T1, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Set Freight")
        Load_Popup_Menu(grdARTCCPAC, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Clear Entry")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdEDT945T1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Set Freight"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = grd.ActiveRow.Band.Index = 0

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"

            Case "De-Select All"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All for Customer"

            Case "Set Freight"
                Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                Dim EDI_FRT_COST As Decimal = Val(grd.ActiveRow.Cells("EDI_FRT_COST").Value & String.Empty)

                Dim new_EDI_FRT_COST As String = InputBox("Do you want to change the Freight Amount for this shipment?", "Change Freight")
                If Val(new_EDI_FRT_COST) = EDI_FRT_COST Then
                    Exit Sub
                End If

                If Val(new_EDI_FRT_COST) < 0 Then
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to chnage the freight on this shipment from " & EDI_FRT_COST.ToString("#,##0.00") & " to " _
                                    & Val(new_EDI_FRT_COST).ToString("#,##0.00") & "?", "Change Freight", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ASCDATA1.ExecuteSQL("Update EDT945T1 SET EDI_FRT_COST = " & new_EDI_FRT_COST & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                grd.ActiveRow.Cells("EDI_FRT_COST").Value = new_EDI_FRT_COST

            Case "Clear Entry"
                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value
                Dim INV_NO As Decimal = grd.ActiveRow.Cells("INV_NO").Value

                If MessageBox.Show("Do you want to remove the entry for Pick Ticket: " & PICK_NO & ",Invoice No: " & INV_NO & "?", "Clear CC Entry", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ASCDATA1.ExecuteSQL("DELETE FROM ARTCCPAC WHERE PICK_NO = '" & PICK_NO & "' AND INV_NO = '" & INV_NO & "'")
                grd.ActiveRow.Delete()

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region

End Class