'Imports System.Windows.Forms
'Imports ABSolution
'Imports Infragistics.Win
'Imports System.IO
'Imports System.Data

Imports Infragistics.Win.UltraWinGrid

Public Class EDF860T1

#Region "Class Variables"
    Private tempTable As String = String.Empty
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    ''' <summary>
    ''' ** Sets up required Data Tables and intializes form controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            tempTable = ASCMAIN1.Temp_Table("Select EDI_DOC_SEQ_NO from EDT860T1 Where ROWNUM < 0")

            ASCMAIN1.sql = $"SELECT * FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tempTable})"
            Create_TDA(.Tables.Add, "EDT860T1", ASCMAIN1.sql, 0, False)
            .Tables("EDT860T1").Columns.Add("SELECTED", GetType(System.String))

            ASCMAIN1.sql = $"SELECT * FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tempTable})"
            Create_TDA(.Tables.Add, "EDT860T1X", ASCMAIN1.sql, 0, False)

            ASCMAIN1.sql = $"SELECT * FROM EDT860T2 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tempTable})"
            Create_TDA(.Tables.Add, "EDT860T2", ASCMAIN1.sql, 0, False)

            ASCMAIN1.sql = $"SELECT * FROM EDT860T3 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tempTable})"
            Create_TDA(.Tables.Add, "EDT860T3", ASCMAIN1.sql, 0, False)

            ASCMAIN1.sql = $"SELECT * FROM EDT860T4 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tempTable})"
            Create_TDA(.Tables.Add, "EDT860T4", ASCMAIN1.sql, 0, False)

            Create_Relation("EDT860T1", "EDT860T2", "EDI_DOC_SEQ_NO")
            Create_Relation("EDT860T2", "EDT860T3", "EDI_DOC_SEQ_NO,EDI_ORIG_DTL_SEQ")
            Create_Relation("EDT860T1", "EDT860T4", "EDI_DOC_SEQ_NO")

            ASCMAIN1.sql = $"SELECT * FROM SOTORDR0 WHERE CUST_CODE = :PARM1 AND ORDR_DATE >= TRUNC(SYSDATE - 360) AND ORDR_STATUS <> 'C'"
            Create_TDA(.Tables.Add, "SOTORDR0", ASCMAIN1.sql, 0, False, "V")
            .Tables("SOTORDR0").Columns.Add("SELECTED", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)

            ASCMAIN1.sql = "Select EDT855O1.*
                                FROM EDT855O1, EDTSYSIH, EDTTRPM1
                                WHERE EDT855O1.COMPANY_CODE = EDTSYSIH.COMPANY_CODE 
                                AND EDT855O1.EDI_OUTBOUND_DOC_NO = EDTSYSIH.EDI_OUTBOUND_DOC_NO
                                AND EDTSYSIH.EDI_TP_ID = EDTTRPM1.EDI_TP_ID
                                AND EDTSYSIH.EDI_OUR_ID = EDTTRPM1.EDI_OUR_ID
                                AND EDTTRPM1.EDI_DOC_NO = '855'
                                AND EDTTRPM1.CUST_CODE = :PARM1
                                AND EDT855O1.ORDR_CUST_PO IN (:PARM2)"
            Create_TDA(.Tables.Add, "EDT855O1", ASCMAIN1.sql, 0, False, "VV")

            Create_Relation("SOTORDR0", "EDT860T1X", "ORDR_CUST_PO", "EDI_PO_NO")
            Create_Relation("SOTORDR0", "EDT855O1", "ORDR_CUST_PO", "ORDR_CUST_PO")

        End With

        grdEDT860T1.DataSource = dst.Tables("EDT860T1")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")

        Create_Summary(grdEDT860T1, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT860T1, "EDI_ORIG_DTL_SEQ", "Count", grdEDT860T1.DisplayLayout.Bands(1).Key)

        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")

        Create_Summary(grdSOTORDR0, "CUST_CODE", "Count")
        Create_Summary(grdSOTORDR0, "SELECTED", "Sum")

        ASCMAIN1.sql = "Select '01' EDI_TRANS_PURP, 'Order Cancellation' EDI_TRANS_DESC FROM DUAL
                        UNION Select '04' EDI_TRANS_PURP, 'Change' EDI_TRANS_DESC FROM DUAL
                        UNION Select '17' EDI_TRANS_PURP, 'Cancel, to be Reissued' EDI_TRANS_DESC FROM DUAL
                        UNION Select '44' EDI_TRANS_PURP, 'Rejection' EDI_TRANS_DESC FROM DUAL
                        UNION Select '54' EDI_TRANS_PURP, 'Approval' EDI_TRANS_DESC FROM DUAL
                        UNION Select 'RZ' EDI_TRANS_PURP, 'Replace All Values' EDI_TRANS_DESC FROM DUAL"
        ASCMAIN1.Add_Value_List(grdEDT860T1, "EDI_TRANS_PURP", Nothing, Nothing, 0, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select 'QC' EDI_CHANGE_TYPE, 'Quantity Change' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'PQ' EDI_CHANGE_TYPE, 'Quantity Change' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'QI' EDI_CHANGE_TYPE, 'Quantity Increase' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'QD' EDI_CHANGE_TYPE, 'Quantity Decrease' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'PC' EDI_CHANGE_TYPE, 'Price Change' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'AI' EDI_CHANGE_TYPE, 'Add Item' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'DI' EDI_CHANGE_TYPE, 'Delete Item' V FROM DUAL
                        UNION Select 'CA' EDI_CHANGE_TYPE, 'Change to Line' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'RZ' EDI_CHANGE_TYPE, 'Replace All Values' EDI_CHANGE_DESC FROM DUAL
                        UNION Select 'PC' EDI_CHANGE_TYPE, 'Price Change' EDI_CHANGE_DESC FROM DUAL"
        ASCMAIN1.Add_Value_List(grdEDT860T1, "EDI_CHANGE_TYPE", Nothing, Nothing, 1, ASCMAIN1.sql)
        'ASCMAIN1.Add_Value_List(grdSOTORDR0, "EDI_CHANGE_TYPE", Nothing, Nothing, 1, ASCMAIN1.sql)

        ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDR2, "ORDR_STATUS")

        For Each dc As UltraWinGrid.UltraGridColumn In grdEDT860T1.DisplayLayout.Bands(0).Columns
            Select Case dc.Key
                Case "SELECTED"
                    dc.CellActivation = UltraWinGrid.Activation.AllowEdit
                Case Else
                    dc.CellActivation = UltraWinGrid.Activation.NoEdit
            End Select
        Next

        grdEDT860T1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdEDT860T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdEDT860T1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        grdSOTORDR0.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTORDR0.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdSOTORDR0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        splEDT860.Dock = DockStyle.Fill
        splEDT860.Visible = False

        grdSOTORDR0.Dock = DockStyle.Fill
        grdSOTORDR0.Visible = False
    End Sub

    ''' <summary>
    ''' Clear tables and controls based on the current state of the screen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRecord()

        EnforceConstraints(False)
        For Each tableName As String In New String() {"EDT860T1", "EDT860T2", "EDT860T3", "EDT860T4", "SOTORDR0", "SOTORDR1", "SOTORDR2", "EDT855O1", "EDT860T1X"}
            dst.Tables(tableName).Clear()
        Next
        EnforceConstraints(True)

        grdSOTORDR1.Text = ""
        grdSOTORDR2.Text = ""
        txtCUST_CODE.Clear()
        txtCUST_CODE.Enabled = True
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False

    End Sub

    ''' <summary>
    ''' Load up changes to go into 832 files
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadRecord()

        ASCMAIN1.Progress("Load EDI 860 Records", String.Empty)

        Try
            EnforceConstraints(False)

            ' Update 860 records
            ASCDATA1.ExecuteSQL("Update EDT860T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT860T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC where EDI_OUR_ID = EDT860T1.EDI_OUR_ID and EDI_TP_ID = EDT860T1.EDI_TP_ID) 
                                where EDI_PROCESS_IND = '0' 
                                and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Update EDT860T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1 where EDI_TP_QUAL = EDT860T1.EDI_TP_QUAL and EDI_TP_ID = EDT860T1.EDI_TP_ID and EDI_DOC_NO = 860)
                                where EDI_PROCESS_IND = '0' 
                                and CUST_CODE IS NULL 
                                and COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {tempTable}")

            ASCMAIN1.sql = String.Empty

            Absx1.txtFor("CUST_CODE").Text = Absx1.txtFor("CUST_CODE").Text.Trim.Replace("'", "")

            Dim lstPOs As New List(Of String)
            Dim lstOrderGroupNos As New List(Of String)

            If Absx1.txtFor("CUST_CODE").TextLength > 0 Then
                ASCMAIN1.sql &= $" AND CUST_CODE = '{Absx1.txtFor("CUST_CODE").Text}'"

                For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select("SELECTED = '1'")
                    lstPOs.Add(rowSOTORDR0.Item("ORDR_CUST_PO") & String.Empty)
                    lstOrderGroupNos.Add(rowSOTORDR0.Item("ORDR_GROUP_NO") & String.Empty)
                Next

                If lstPOs.Count > 0 Then
                    ASCMAIN1.sql &= $" AND EDI_PO_NO IN ('{String.Join("','", lstPOs.ToArray)}')"
                End If

                ASCMAIN1.sql = $"Select EDI_DOC_SEQ_NO 
                                from EDT860T1
                                Where EDI_PO_TYPE NOT IN ('KC', 'KB') 
                                and COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'" & ASCMAIN1.sql

            Else
                ASCMAIN1.sql = $"Select EDI_DOC_SEQ_NO 
                                from EDT860T1
                                Where NVL(EDI_PROCESS_IND, '0') = '0' 
                                and EDI_PO_TYPE NOT IN ('KC', 'KB') 
                              and COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'"
            End If

            ' Type KC, KB means Contract and are viewed/processed on another screen.
            ASCDATA1.ExecuteSQL($"INSERT INTO {tempTable} {ASCMAIN1.sql}")

            If Absx1.txtFor("CUST_CODE").TextLength > 0 Then
                For Each tableName As String In New String() {"EDT860T1X"}
                    Fill_Records(tableName)
                Next

                ASCMAIN1.sql = $"SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO IN ('{String.Join("','", lstOrderGroupNos.ToArray)}')  AND ORDR_STATUS <> 'C'"
                Fill_Records("SOTORDR0", String.Empty, True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"Select EDT855O1.*
                                FROM EDT855O1, EDTSYSIH, EDTTRPM1
                                WHERE EDT855O1.COMPANY_CODE = EDTSYSIH.COMPANY_CODE 
                                AND EDT855O1.EDI_OUTBOUND_DOC_NO = EDTSYSIH.EDI_OUTBOUND_DOC_NO
                                AND EDTSYSIH.EDI_TP_ID = EDTTRPM1.EDI_TP_ID
                                AND EDTSYSIH.EDI_OUR_ID = EDTTRPM1.EDI_OUR_ID
                                AND EDTTRPM1.EDI_DOC_NO = '855'
                                AND EDTTRPM1.CUST_CODE = '{txtCUST_CODE.Text}'
                                AND EDT855O1.ORDR_CUST_PO IN ('{ String.Join("','", lstPOs.ToArray)}')"
                Fill_Records("EDT855O1", String.Empty, True, ASCMAIN1.sql)
            Else
                For Each tableName As String In New String() {"EDT860T1", "EDT860T2", "EDT860T3", "EDT860T4"}
                    Fill_Records(tableName)
                Next
            End If

            EnforceConstraints(True)

            grdEDT860T1.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = dst.Tables("EDT860T1").Select("EDI_PROCESS_IND = '0'").Length = 0
            txtCUST_CODE.Enabled = False

            grdSOTORDR0.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True

        Catch ex As Exception
            ClearRecord()
            MessageBox.Show(ex.Message)
        Finally
            grdEDT860T1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        End Try

    End Sub

    ''' <summary>
    ''' Sets up screen based on the form modality, state and type of processing
    ''' </summary>
    ''' <param name="tf"></param>
    ''' <remarks></remarks>
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_Description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Customer").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Load POs").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("View Open 860s").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        If ScreenMode Then
            splEDT860.Visible = txtCUST_CODE.TextLength = 0
            grdSOTORDR0.Visible = txtCUST_CODE.TextLength > 0
        Else ClearRecord()
            splEDT860.Visible = False
            grdSOTORDR0.Visible = True
        End If
    End Sub

    ''' <summary>
    ''' Validates data when a user selects a menu option
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim sql As String = String.Empty

        Select Case eItemKey

            Case "View Open 860s"
                txtCUST_CODE.Clear()

            Case "Load POs"
                If dst.Tables("SOTORDR0").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select at least 1 PO."
                End If

            Case "Done"

            Case "Load Customer"
                Validate_Code("CUST_CODE")

        End Select

        If EMsg <> String.Empty Then
            MessageBox.Show(EMsg, "Cannot Proceed", MessageBoxButtons.OK)
        Else
            Call Proceed(eItemKey)
        End If

    End Sub

    ''' <summary>
    ''' When the user selects a menu option perform the action
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load Customer"
                EnforceConstraints(False)
                Fill_Records("SOTORDR0", txtCUST_CODE.Text)
                Sort_grdColumns(grdSOTORDR0, "ORDR_DATE")
                EnforceConstraints(True)
                With UltraExplorerBar1
                    .Groups("Screen Control").Items("Load Customer").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Load POs").Settings.Enabled = DefaultableBoolean.True
                    .Groups("Screen Control").Items("View Open 860s").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
                End With


            Case "View Open 860s"
                LoadRecord()
                Mode_Settings(True)
                With UltraExplorerBar1
                    .Groups("Screen Control").Items("Load Customer").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Load POs").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("View Open 860s").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
                End With

            Case "Load POs"
                LoadRecord()
                Mode_Settings(True)
                With UltraExplorerBar1
                    .Groups("Screen Control").Items("Load Customer").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Load POs").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("View Open 860s").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
                End With

            Case "Done"
                Mode_Settings(False)

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    ''' <summary>
    ''' Updates data based on the current state of the screen and the type of processing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateRecord()

        Try
            BeginTrans()

            CommitTrans("Update Successful")

        Catch ex As Exception
            Rollback(ex.Message)

        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT860T1, "SSSPBBPBBPB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "Deselect All", "Select All Customer", "Deselect All Customer", "Mark Selected as Completed")
        Load_Popup_Menu(grdSOTORDR1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Sales Order Entry")
        Load_Popup_Menu(grdSOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
        Load_Popup_Menu(grdSOTORDR0, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case grdEDT860T1.Name
                    Dim numSelected As Int32 = dst.Tables("EDT860T1").Select("SELECTED = '1' AND EDI_PROCESS_IND = '0'").Length
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text

                    For Each key As String In New String() {"Select All", "Deselect All", "Mark Selected as Completed", "Select All Customer", "Deselect All Customer"}
                        tlb_btn = DirectCast(tlb.Tools(key), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False

                        If grdEDT860T1.ActiveRow.Band.Key = grdEDT860T1.DisplayLayout.Bands(0).Key Then
                            Select Case key
                                Case "Mark Selected as Completed"
                                    tlb_btn.SharedProps.Visible = numSelected > 0

                                Case "Select All Customer", "Deselect All Customer"
                                    tlb_btn.SharedProps.Caption = tlb_btn.Key.Replace("Customer", CUST_CODE)
                                    tlb_btn.SharedProps.Visible = dst.Tables("EDT860T1").Select("EDI_PROCESS_IND = '0'").Length > 0

                                Case Else
                                    tlb_btn.SharedProps.Visible = dst.Tables("EDT860T1").Select("EDI_PROCESS_IND = '0'").Length > 0
                            End Select
                        End If
                    Next
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

            Case "Mark Selected as Completed"
                Dim numSelected As Int32 = dst.Tables("EDT860T1").Select("SELECTED = '1' AND EDI_PROCESS_IND = '0'").Length
                If numSelected <= 0 Then
                    Exit Sub
                End If

                Dim msg As String = $"Do you want to mark the {numSelected} selected EDI 860 {IIf(numSelected = 1, "record", "records")} as completed?"

                If MessageBox.Show(msg, "Mark as Completed", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    Try
                        BeginTrans()
                        For Each row As DataRow In dst.Tables("EDT860T1").Select("SELECTED = '1' AND EDI_PROCESS_IND = '0'")
                            Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO")
                            ASCMAIN1.sql = "UPDATE EDT860T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO = :PARM1"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {EDI_DOC_SEQ_NO})
                            row.Item("EDI_PROCESS_IND") = "1"
                        Next
                        CommitTrans("Update Successful")

                        Try
                            Click_Command("View Open 860s")
                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Mark Selected as Completed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                End If

            Case "Select All", "Deselect All"
                Dim selected As String = If(e.Tool.Key = "Select All", "1", "0")

                For Each grdRow As UltraGridRow In grdEDT860T1.Rows.GetFilteredInNonGroupByRows
                    Dim EDI_DOC_SEQ_NO As String = grdRow.Cells("EDI_DOC_SEQ_NO").Value
                    dst.Tables("EDT860T1").Select($"EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'")(0).Item("SELECTED") = selected
                Next
                dst.Tables("EDT860T1").AcceptChanges()

            Case "Select All Customer", "Deselect All Customer"
                Dim selected As String = If(e.Tool.Key = "Select All Customer", "1", "0")
                Dim CUST_CODE As String = grdEDT860T1.ActiveRow.Cells("CUST_CODE").Value & String.Empty

                For Each grdRow As UltraGridRow In grdEDT860T1.Rows.GetFilteredInNonGroupByRows
                    If grdRow.Cells("CUST_CODE").Value & String.Empty <> CUST_CODE Then
                        Continue For
                    End If
                    Dim EDI_DOC_SEQ_NO As String = grdRow.Cells("EDI_DOC_SEQ_NO").Value
                    dst.Tables("EDT860T1").Select($"EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'")(0).Item("SELECTED") = selected
                Next
                dst.Tables("EDT860T1").AcceptChanges()

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Sales Order Entry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Context_Launch("Edit", ORDR_NO, e.Tool.Key, "SOFORDR1")

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")

        End Select
    End Sub

#End Region

#Region "Control Events"

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub grdEDT860T1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT860T1.AfterRowActivate

        If Not grdEDT860T1.ActiveRow.IsDataRow Then
            Dim dvwX As DataView = dst.Tables("SOTORDR1").DefaultView
            dvwX.RowFilter = $"CUST_CODE = '@@@@@@@@@'"
            grdSOTORDR1.Text = String.Empty

            dvwX = dst.Tables("SOTORDR2").DefaultView
            dvwX.RowFilter = $"ORDR_NO = '@@@@@@@@@'"
            grdSOTORDR2.Text = String.Empty

            Exit Sub
        End If

        Dim EDI_PO_NO As String
        Dim CUST_CODE As String

        Select Case grdEDT860T1.ActiveRow.Band.Index
            Case 0
                EDI_PO_NO = grdEDT860T1.ActiveRow.Cells("EDI_PO_NO").Text & String.Empty
                CUST_CODE = grdEDT860T1.ActiveRow.Cells("CUST_CODE").Text & String.Empty

                If dst.Tables("SOTORDR1").Select($"CUST_CODE = '{CUST_CODE}' AND ORDR_CUST_PO = '{EDI_PO_NO}'").Length = 0 Then
                    ASCMAIN1.sql = $"SELECT * FROM SOTORDR1 WHERE CUST_CODE = '{CUST_CODE}' AND ORDR_CUST_PO = '{EDI_PO_NO}'"
                    Fill_Records("SOTORDR1", String.Empty, False, ASCMAIN1.sql)

                    ASCMAIN1.sql = $"SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE CUST_CODE = '{CUST_CODE}' AND ORDR_CUST_PO = '{EDI_PO_NO}')"
                    Fill_Records("SOTORDR2", String.Empty, False, ASCMAIN1.sql)
                End If

            Case Else
                Exit Sub
        End Select

        Dim dvw As DataView = dst.Tables("SOTORDR1").DefaultView
        dvw.RowFilter = $"CUST_CODE = '{CUST_CODE}' AND ORDR_CUST_PO = '{EDI_PO_NO}'"
        Sort_grdColumns(grdSOTORDR1, "ORDR_NO")

        grdSOTORDR1.Text = $"Sales Orders For Customer {CUST_CODE}, P.O. {EDI_PO_NO}"
        grdSOTORDR1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.VisibleRows, True)

    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR1.AfterRowActivate
        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Text
        Dim dvw As DataView = dst.Tables("SOTORDR2").DefaultView
        dvw.RowFilter = $"ORDR_NO = '{ORDR_NO}'"
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

        grdSOTORDR2.Text = $"Order Details for Sales Order {ORDR_NO}"
        grdSOTORDR2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.VisibleRows, True)

    End Sub


#End Region

End Class