Imports Infragistics.Win.UltraWinGrid

Public Class EDF856O1

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "EDT856O1", "*")
            Create_TDA(.Tables.Add, "EDT856O2", "*", 2)
            Create_TDA(.Tables.Add, "EDT856O3", "*", 2)
            Create_TDA(.Tables.Add, "EDT856O4", "*", 2)
            Create_TDA(.Tables.Add, "EDT856O5", "*", 2)
            Create_TDA(.Tables.Add("EDT856O5_0"), "EDT856O5", "*", 2)
            Create_TDA(.Tables.Add, "EDTSYSIH", "*")

            With .Tables("EDT856O2")
                .Columns.Add("SELECTED", GetType(String))
                .Columns("SELECTED").DefaultValue = "0"
            End With
        End With

        CascadeRelationDeleteRule()

        grdEDT856O1.DataSource = dst.Tables("EDT856O1")

        grdEDT856O1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdEDT856O1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdEDT856O1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        For Each iband As UltraGridBand In grdEDT856O1.DisplayLayout.Bands
            For Each iCol As UltraGridColumn In iband.Columns
                Select Case iCol.Key
                    Case "SELECTED", "BILL_OF_LADING_NO", "SHIP_ADDR_CODE"
                        iCol.CellActivation = Activation.AllowEdit
                        iCol.CellAppearance.BackColor = Drawing.Color.LightGreen
                        iCol.Header.VisiblePosition = 0
                    Case Else
                        iCol.CellActivation = Activation.NoEdit
                End Select
            Next
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                EDI_OUTBOUND_DOC_NO = txtEDI_OUTBOUND_DOC_NO.Text.Trim
                If EDI_OUTBOUND_DOC_NO.Length = 0 Then
                    Dim PICK_NO As String = txtPICK_NO.Text.Trim
                    If PICK_NO.Length > 0 Then
                        ASCMAIN1.sql = "SELECT * FROM EDT856O2 WHERE PICK_NO = :PARM1"
                        Dim TBL As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "EDT856O2", "V", {PICK_NO})
                        Select Case TBL.Rows.Count
                            Case 0
                                EMsg &= vbCr & $"Cannot locate an ASN for Pick Ticket No {PICK_NO}."
                                Exit Select
                            Case 1
                                txtEDI_OUTBOUND_DOC_NO.Text = TBL.Rows(0).Item("EDI_OUTBOUND_DOC_NO")
                                EDI_OUTBOUND_DOC_NO = txtEDI_OUTBOUND_DOC_NO.Text.Trim
                            Case Else
                                EMsg &= vbCr & $"There are multiple ASNs for Pick Ticket No {PICK_NO}. Please provide the EDI Outbound Doc No."
                                Exit Select
                        End Select
                    End If
                End If

                If EDI_OUTBOUND_DOC_NO.Length = 0 Then
                    EMsg &= vbCr & $"EDI Outbound Doc No is required."
                    Exit Select
                End If

                Dim rowEDT856O1 As DataRow = LookUp("EDT856O1", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
                If rowEDT856O1 Is Nothing Then
                    EMsg &= vbCr & $"EDI Outbound Doc No {EDI_OUTBOUND_DOC_NO} cannot be located."
                    Exit Select
                End If

                If Not ASCMAIN1.Logical_Lock("EDT856O1", EDI_OUTBOUND_DOC_NO) Then
                    Exit Sub
                End If

            Case "Cancel"

            Case "Update"
                grdEDT856O1.UpdateData()

                Dim numSelected As Int32 = dst.Tables("EDT856O2").Select("SELECTED = '1'").Length

                If numSelected = 0 Then
                    EMsg &= vbCr & $"There are no selected Pick Tickets."
                    Exit Select
                End If

                Dim BILL_OF_LADING_NO As String = dst.Tables("EDT856O1").Rows(0).Item("BILL_OF_LADING_NO") & String.Empty
                Dim EDI_CUSTOMER As String = dst.Tables("EDT856O1").Rows(0).Item("EDI_CUSTOMER") & String.Empty
                BILL_OF_LADING_NO = BILL_OF_LADING_NO.Trim
                If BILL_OF_LADING_NO.Length = 0 Then
                    EMsg &= vbCr & $"Bill of Lading No is Required."
                    Exit Select
                End If

                ASCMAIN1.sql = "SELECT * FROM EDT856O1 WHERE BILL_OF_LADING_NO = :PARM1 AND EDI_CUSTOMER = :PARM2"
                Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "EDT856O1", "VV", {BILL_OF_LADING_NO, EDI_CUSTOMER})
                If tbl.Rows.Count > 0 Then
                    If MessageBox.Show($"There are {tbl.Rows.Count} ASNs for this customer with Bill of Lading No {BILL_OF_LADING_NO}. Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                ElseIf MessageBox.Show($"Do you want to create a separate ASN for the {numSelected} selected Pick Tickets.", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

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
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                End With

                '  Always show the filter
                ' .Groups("Filter").Visible = tf
            End With
        End If

        If ScreenMode Then

        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5", "EDT856O5_0", "EDTSYSIH"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EDI_OUTBOUND_DOC_NO = String.Empty
        EnforceConstraints(True)

        txtEDI_OUTBOUND_DOC_NO.Clear()
        txtPICK_NO.Clear()
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")

        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        Fill_Records("EDT856O1", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
        Fill_Records("EDT856O2", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
        Fill_Records("EDT856O3", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
        Fill_Records("EDT856O4", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
        Fill_Records("EDT856O5", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})
        Fill_Records("EDTSYSIH", {ASCMAIN1.CLIENT, EDI_OUTBOUND_DOC_NO})

        For Each rowEDT856O5 As DataRow In dst.Tables("EDT856O5").Select("EDI_HL2_SEQ = 0")
            dst.Tables("EDT856O5_0").ImportRow(rowEDT856O5)
            rowEDT856O5.Delete()
        Next
        dst.Tables("EDT856O5").AcceptChanges()
        dst.Tables("EDT856O5_0").AcceptChanges()

        EnforceConstraints(True)

        Dim rowEDT856O1 As DataRow = dst.Tables("EDT856O1").Rows(0)
        Dim EDI_CUSTOMER As String = rowEDT856O1.Item("EDI_CUSTOMER") & String.Empty

        ASCMAIN1.sql = $"SELECT DISTINCT CUST_DC_NO SHIP_ADDR_CODE, CUST_DC_NO FROM ARTCUST2 WHERE CUST_CODE = '{EDI_CUSTOMER}' ORDER BY CUST_DC_NO"
        ASCMAIN1.Add_Value_List(grdEDT856O1, "SHIP_ADDR_CODE", ASCMAIN1.sql)

        EnforceConstraints(True)

        ASCMAIN1.Progress("", "")

        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

        Try
            Dim NEW_EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
            BeginTrans()
            CascadeRelationDeleteRule()

            For Each rowEDT856O2 As DataRow In dst.Tables("EDT856O2").Select("")
                If rowEDT856O2.Item("SELECTED") & String.Empty <> "1" Then
                    rowEDT856O2.Delete()
                End If
            Next

            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5", "EDT856O5_0", "EDTSYSIH"}
                dst.Tables(TABLE_NAME).AcceptChanges()
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    row.Item("EDI_OUTBOUND_DOC_NO") = NEW_EDI_OUTBOUND_DOC_NO
                Next
                dst.Tables(TABLE_NAME).AcceptChanges()
            Next
            EnforceConstraints(True)

            Dim EDI_SHIP_CNT_CARTONS As Int32 = Val(dst.Tables("EDT856O2").Compute("SUM(EDI_ORD_CNT_CARTONS)", "") & String.Empty)
            Dim CART_TOTAL_WGT_ACTUAL As Decimal = Val(dst.Tables("EDT856O3").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
            Dim SHIP_ADDR_CODE As String = dst.Tables("EDT856O1").Rows(0).Item("SHIP_ADDR_CODE")

            For Each rowEDT856O5_0 As DataRow In dst.Tables("EDT856O5_0").Select("EDI_ADDR_TYPE = 'ST'")
                rowEDT856O5_0.Item("EDI_ADDR_CODE") = SHIP_ADDR_CODE
            Next

            dst.Tables("EDT856O1").Rows(0).Item("EDI_SHIP_CNT_CARTONS") = EDI_SHIP_CNT_CARTONS
            dst.Tables("EDT856O1").Rows(0).Item("EDI_SHIP_TOTAL_WGT") = CART_TOTAL_WGT_ACTUAL
            dst.Tables("EDTSYSIH").Rows(0).Item("EDI_PROCESS_IND") = "1"

            For Each TABLE_NAME As String In New String() {"EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5", "EDT856O5_0", "EDTSYSIH"}
                dst.Tables(TABLE_NAME).AcceptChanges()
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    row.SetAdded()
                Next
                Update_Record_TDA(TABLE_NAME)
            Next

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            EnforceConstraints(True)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT856O1, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")
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
        Dim tlb_btn1 As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
            Exit Sub
        End If

        Select Case e.SourceControl.Name
            Case grdEDT856O1.Name
        End Select
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
            Case "Auto Fit Columns"
                Me.Cursor = Cursors.WaitCursor
                grd.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
                Me.Cursor = Cursors.Default
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub CascadeRelationDeleteRule()

        For Each relationship As String In New String() {"EDT856O1_EDT856O2", "EDT856O2_EDT856O3", "EDT856O2_EDT856O5", "EDT856O3_EDT856O4", "EDT856O1_EDT856O5_0"}
            If Not dst.Relations.Contains(relationship) Then
                Select Case relationship
                    Case "EDT856O1_EDT856O2"
                        Create_Relation("EDT856O1", "EDT856O2", "COMPANY_CODE,EDI_OUTBOUND_DOC_NO")
                    Case "EDT856O2_EDT856O3"
                        Create_Relation("EDT856O2", "EDT856O3", "COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_HL2_SEQ")
                    Case "EDT856O3_EDT856O4"
                        Create_Relation("EDT856O3", "EDT856O4", "COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_HL2_SEQ,EDI_HL3_SEQ")
                    Case "EDT856O2_EDT856O5"
                        Create_Relation("EDT856O2", "EDT856O5", "COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_HL2_SEQ")
                    Case "EDT856O1_EDT856O5_0"
                        Create_Relation("EDT856O1", "EDT856O5_0", "COMPANY_CODE,EDI_OUTBOUND_DOC_NO")
                End Select
            End If
        Next

        dst.Relations("EDT856O1_EDT856O2").ChildKeyConstraint.DeleteRule = Rule.Cascade
        dst.Relations("EDT856O2_EDT856O3").ChildKeyConstraint.DeleteRule = Rule.Cascade
        dst.Relations("EDT856O3_EDT856O4").ChildKeyConstraint.DeleteRule = Rule.Cascade
        dst.Relations("EDT856O2_EDT856O5").ChildKeyConstraint.DeleteRule = Rule.Cascade

    End Sub

#End Region

End Class