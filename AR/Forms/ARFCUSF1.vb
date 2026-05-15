Public Class ARFCUSF1

#Region "Declarations"
    Dim REALIGN_NO As String
    Dim rowARTCUSF1 As DataRow
    Dim ARTCUSFX As String
    Dim REALIGN_NO_000000 As String = "000000"
    Dim SELL_CODE As String

    Dim updatedByImport As Integer = 0
    Dim importErrors As New Dictionary(Of Integer, List(Of String))


#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_InquiryMode()
        Get_PARM("SOTPARM1")

        Create_Work_Tables(" and rownum < 1")

        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)

            ASCMAIN1.sql = "Select ARTCUSF1.*" _
                & " from ARTCUSF1"
            Create_TDA(.Tables.Add, "ARTCUSFX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "ARTCUSF1", "*", 1)

            ASCMAIN1.sql = "Select ARTCUSF2.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.SELL_CODE SELL_CODE_CURR, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
                & " from ARTCUSF2,ARTCUST2" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = ARTCUSF2.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = ARTCUSF2.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "ARTCUSF2", "**", 1)
            'Create_TDA(.Tables.Add, "ARTCUSFE", "**", 0, False, "")
            'With .Tables("ARTCUSFE").Columns
            '    .Add("IMPORT_ERROR")
            'End With

            Dim importErrorTable As New DataTable("ARTCUSIE")
            dst.Tables.Add(importErrorTable)
            With .Tables("ARTCUSIE").Columns
                .Add("ROW", GetType(System.Int32))
                .Add("IMPORT_ERROR", GetType(System.String))
            End With
        End With

        grdARTCUSFX.DataSource = dst.Tables("ARTCUSFX")
        grdARTCUSF2.DataSource = dst.Tables("ARTCUSF2")
        grdARTCUSFE.DataSource = dst.Tables("ARTCUSIE")

        Create_Summary(grdARTCUSFX, "REALIGN_NO", "Count")
        Create_Summary(grdARTCUSF2, "CUST_STORE_NO", "Count")
        Create_Summary(grdARTCUSFE, "ROW", "Count")



        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdARTCUSF2}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_STORE_NO", "CUST_STORE_NAME", "CUST_STORE_STATUS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"SELL_CODE"}.Contains(gcol.Key) Then
                        gcol.CellActivation = IIf(grd.Name = "grdARTCUSF2", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                        gcol.CellAppearance.BackColor = IIf(grd.Name = "grdARTCUSF2", Drawing.Color.White, Drawing.Color.Beige)
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige

                    End If
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"CUST_CODE", "CUST_STORE_NO", "CUST_STORE_NAME", "CUST_STORE_STATUS"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    ElseIf New String() {"SELL_CODE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"SELL_CODE_CURR"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                Next

            End With
            ASCMAIN1.Add_Value_List(grd, "CUST_STORE_STATUS")
        Next

        With grdARTCUSFE.DisplayLayout
            .AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn

            With .Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"ROW"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    End If
                    If New String() {"IMPORT_ERROR"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                        gcol.CellAppearance.ForeColor = Drawing.Color.Red
                    End If

                Next
            End With


        End With

        spl.Panel1Collapsed = True
        splImportErrors.Panel2Collapsed = True
        chkImportErrors.Visible = False

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "ARFCUSFI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If LookUp("ARTCUSF1", REALIGN_NO_000000) IsNot Nothing Then
                    EMsg &= vbCr & "A Future Re-Alignment has already been started"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCUSF1", REALIGN_NO_000000) Then Exit Sub
                End If

            Case "Edit", "View"

                Absx1.txtFor("REALIGN_NO").Text = REALIGN_NO_000000

                If Absx1.txtFor("REALIGN_NO").Text = "" Then
                    EMsg &= vbCr & "No Re-Alignment No Specified"
                Else
                    REALIGN_NO = Absx1.txtFor("REALIGN_NO").Text
                    Dim row As DataRow = LookUp("ARTCUSF1", REALIGN_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Re-Alignment No " & REALIGN_NO
                    Else
                        If EMsg.Length = 0 Then
                            If Not ASCMAIN1.Logical_Lock("ARTCUSF1", REALIGN_NO) Then Exit Sub
                        End If
                    End If
                End If

            Case "Update"
                If Absx1.dteFor("REALIGN_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Re-Alignment Date is Mandatory"
                Else
                    If Format(Absx1.dteFor("REALIGN_DATE").Value, "yyyyMMdd") _
                     < ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1) & "01" Then
                        EMsg &= vbCr & "Re-Alignment Date cannot be Prior to 1st of Next Month"
                    End If
                End If

                If grdARTCUSF2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Stores in Re-Alignment Grid"
                Else
                    If Val(dst.Tables("ARTCUSF2").Compute("COUNT(CUST_STORE_NO)", "SELL_CODE <> SELL_CODE_CURR") & "") = 0 Then
                        EMsg &= vbCr & "No Stores with New AE Codes"
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                If EMsg = "" Then
                    If MsgBox("Do you want to Delete this Re-Alignment", _
                                MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Publish Re-Alignment"

                If EntryMode = "" Then
                    Exit Sub
                End If

                If EMsg = "" Then
                    If Format(Now.Date, "yyyyMMdd") < Format(Absx1.dteFor("REALIGN_DATE").Value, "yyyyMMdd") Then
                        'EMsg &= vbCr & "Should not be Publishing until on or after " & Format(Absx1.dteFor("REALIGN_DATE").Value, "MM/dd/yyyy")
                        'changed by request from Catherine on 1/25/2019
                        If MsgBox("Do you really want to publish this Re-Alignment today even though the Re-Alignment date is " _
                              & Format(Absx1.dteFor("REALIGN_DATE").Value, "MM/dd/yyyy") & "?" & vbCrLf & vbCrLf & "WARNING: This will update Store/AE Relationships",
            MsgBoxStyle.Question + MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    Else
                        If MsgBox("Do you want to Update this Re-Alignment?" & vbCrLf & vbCrLf & "WARNING: This will update Store/AE Relationships",
                                    MsgBoxStyle.Question + MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    If Not ASCMAIN1.Logical_Lock("ARTCUSF1", REALIGN_NO_000000) Then Exit Sub

                End If
            Case "XLS Upload"
                If EntryMode = "" Then
                    Exit Sub
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Publish Re-Alignment"
                Realignment_Update()
                Mode_Settings(False)

            Case "XLS Upload"
                XLS_Upload()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If True Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("XLS Upload").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)

                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("Publish Re-Alignment").Visible = (EntryMode = "V")
                .Items("XLS Upload").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
            End With
            .Groups("Display Options").Visible = ScreenMode
        End With

        grdARTCUSFX.Visible = Not tf
        tabRAMaster.Visible = Not tf
        chkImportErrors.Visible = False
        splImportErrors.Panel2Collapsed = True

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

            If EntryMode = "V" Then
                Set_Read_Only_for_ctl(Absx1.dteFor("REALIGN_DATE"), True)
                grdARTCUSF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdARTCUSF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdARTCUSF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                Set_Read_Only_for_ctl(Absx1.dteFor("REALIGN_DATE"), False)
                grdARTCUSF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdARTCUSF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdARTCUSF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

                'If EntryMode <> "E" Then
                '    grdARTCUSF2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                'Else
                '    grdARTCUSF2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                'End If
            End If
        Else
            Clear_Record()
            Set_Read_Only_for_ctl(Absx1.dteFor("REALIGN_DATE"), False)
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("REALIGN_NO").Text = ""
        REALIGN_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ARTCUSF1", "ARTCUSF2", "ARTCUSIE"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_ARTCUSFX("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            REALIGN_NO = REALIGN_NO_000000 '  ASCMAIN1.Next_Control_No("ARTCUSF1.REALIGN_NO")

            rowARTCUSF1 = dst.Tables("ARTCUSF1").NewRow
            With rowARTCUSF1
                .Item("REALIGN_NO") = REALIGN_NO
                '  .Item("REALIGN_DATE") = Now.DATE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("ARTCUSF1").Rows.Add(rowARTCUSF1)

            ASCMAIN1.sql = GET_SQL_ARTCUST2
            Fill_Records("ARTCUSF2", "", , ASCMAIN1.sql)

        Else
            rowARTCUSF1 = Fill_Record("ARTCUSF1", REALIGN_NO)

            Fill_Records("ARTCUSF2", REALIGN_NO)

            ASCMAIN1.sql = Get_sql_ARTCUST2() & vbCrLf _
                & " where (CUST_CODE, CUST_STORE_NO) in (" & vbCrLf _
                & " Select CUST_CODE, CUST_STORE_NO from ARTCUST2 minus" & vbCrLf _
                & " Select CUST_CODE, CUST_STORE_NO from ARTCUSF2 where REALIGN_NO = '" & REALIGN_NO & "'" & vbCrLf _
                & ")"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            If tbl.Rows.Count <> 0 Then
                Fill_Records("ARTCUSF2", "", False, ASCMAIN1.sql)
                'For Each row As DataRow In tbl.Rows
                '    row.SetAdded()
                'Next
                'dst.Tables("ARTCUSF2").Merge(tbl)
                'Using frm As New ASFMSGBF
                '    frm.Show_grd(tbl, Me, "New Customer Stores have been added")
                'End Using
            End If
        End If

        Sort_grdColumns(grdARTCUSF2, "CUST_CODE,CUST_STORE_NO")
        set_filter()

        'Display_Totals()

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, REALIGN_NO)
        For Each TABLE_NAME As String In New String() _
            {"ARTCUSF1", "ARTCUSF2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where REALIGN_NO = '" & REALIGN_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub
    Sub Dependent_Updates(S As Integer, REALIGN_NO As String)

    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        For Each rowARTCUSF2 As DataRow In dst.Tables("ARTCUSF2").Select("ISNULL(SELL_CODE,'') = ISNULL(SELL_CODE_CURR,'')")
            rowARTCUSF2.Delete()
        Next

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        INIT_LAST("ARTCUSF1", False, , True)
        Dim sqldelete As String = "REALIGN_NO = '" & REALIGN_NO & "'"
        Update_Record_TDA("ARTCUSF1", sqldelete)
        Update_Record_TDA("ARTCUSF2", sqldelete)
        Dependent_Updates(1, REALIGN_NO)

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CUST_STORE_NO"
                sql_where &= " and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCUSF2, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Copy AE", "Paste AE", "Select All for CUST_CODE")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then grd = GRDs(Mid(e.SourceControl.Name, 4))
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
            If grd.Name <> "grdARTUSF2" Then e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdARTCUSF2"

                    tlb_btn = DirectCast(tlb_pop.Tools("Select All for CUST_CODE"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Or EntryMode = "V" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Select All for " & grd.ActiveRow.Cells("CUST_CODE").Value
                        tlb_btn.Tag = grd.ActiveRow.Cells("CUST_CODE").Value
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Copy AE"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Or EntryMode = "V" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim SELL_CODE As String = grd.ActiveRow.Cells("SELL_CODE").Value & ""
                        If SELL_CODE = "" Then
                            tlb_btn.SharedProps.Visible = False
                        Else
                            tlb_btn.SharedProps.Visible = True
                            tlb_btn.SharedProps.Caption = "Copy AE " & SELL_CODE
                            tlb_btn.Tag = SELL_CODE
                        End If
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Paste AE"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Or EntryMode = "V" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        If SELL_CODE = "" Then
                            tlb_btn.SharedProps.Visible = False
                        Else
                            tlb_btn.SharedProps.Visible = True
                            tlb_btn.SharedProps.Caption = "Paste AE " & SELL_CODE
                            tlb_btn.Tag = SELL_CODE
                        End If
                    End If

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select All for CUST_CODE"
                grdARTCUSF2.Selected.Rows.Clear()

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        If grow.Cells("CUST_CODE").Value = e.Tool.Tag & "" Then
                            'grow.Cells("SEL").Value = "1"
                            'grow.Update()
                            grow.Selected = True
                        End If
                    End If
                Next

            Case "Paste AE"
                If grdARTCUSF2.Selected.Rows.Count = 0 AndAlso grd.ActiveRow.IsDataRow Then
                    grd.ActiveRow.Selected = True
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then

                        grow.Cells("SELL_CODE").Value = SELL_CODE
                        grow.Update()

                    End If
                Next
                grdARTCUSF2.Selected.Rows.Clear()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Copy AE"
                SELL_CODE = grd.ActiveRow.Cells("SELL_CODE").Value
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "REALIGN_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "REALIGN_NO"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_ARTCUSFX(sqlw)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Data")

        Create_Work_Tables(sqlw)

        Fill_Records("ARTCUSFX")

        Dim dvw As DataView = DirectCast(grdARTCUSFX.DataSource, DataTable).DefaultView
        dvw.RowFilter = "REALIGN_NO = '" & REALIGN_NO_000000 & "'"

        Sort_grdColumns(grdARTCUSFX, "REALIGN_NO".ToLower)
        grdARTCUSFX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Create_Work_Tables(sqlw As String)

        ASCMAIN1.sql = "Select ARTCUSF1.*" & vbCrLf _
            & " from ARTCUSF1"

        If ARTCUSFX = "" Then
            ARTCUSFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSFX & " Add Primary Key (REALIGN_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCUSFX)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCUSFX & " " & ASCMAIN1.sql)
        End If

    End Sub

    Private Sub grdARTCUSFX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCUSFX.DoubleClickRow

        If e.Row Is Nothing OrElse e.Row.IsFilterRow Then
            Exit Sub
        End If

        Absx1.txtFor("REALIGN_NO").Text = e.Row.Cells("REALIGN_NO").Value & String.Empty
        Click_Command("View")
    End Sub

    Sub Realignment_Update()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        Dim REALIGN_NO_archived As String = ASCMAIN1.Next_Control_No("ARTCUSF1.REALIGN_NO")
        Dim TABLE_NAME As String = "ARTCUST2"
        Dim COLUMN_NAME As String = "SELL_CODE"
        Dim USER_ID As String = ASCMAIN1.USER_ID
        Dim AUDIT_TRAIL_FIELDS As String = "TABLE_NAME, KEY_VALUE, COLUMN_NAME, USER_ID, INIT_DATE, OLD_VALUE, NEW_VALUE, FM_MODE, NOTES, SESSION_NO, XNO"

        BeginTrans()

        'Update customer store & record audit trail
        Dim sqlR As String = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select " & vbCrLf _
            & "  'ARTCUST2' TABLE_NAME, ARTCUSF2.CUST_CODE || ':' || ARTCUSF2.CUST_STORE_NO KEY_VALUE" & vbCrLf _
            & "  , 'SELL_CODE' COLUMN_NAME, 'wjz' USER_ID, ARTCUSF1.INIT_DATE, ARTCUST2.SELL_CODE OLD_VALUE" & vbCrLf _
            & "  , ARTCUSF2.SELL_CODE NEW_VALUE, 'E' FM_MODE, '" & ASCMAIN1.SESSION_NO & "' SESSION_NO, 'Field Re-Alignment' NOTES" & vbCrLf _
            & "  , '" & REALIGN_NO_archived & "' XNO, ARTCUSF2.CUST_CODE, ARTCUSF2.CUST_STORE_NO" & vbCrLf _
            & " from" & vbCrLf _
            & "  ARTCUSF1, ARTCUST2, ARTCUSF2" & vbCrLf _
            & " where" & vbCrLf _
            & "  ARTCUSF1.REALIGN_NO  = ARTCUSF2.REALIGN_NO" & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = ARTCUSF2.CUST_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_STORE_NO = ARTCUSF2.CUST_STORE_NO" & vbCrLf _
            & "  and ARTCUSF1.REALIGN_NO = '" & REALIGN_NO_000000 & "';" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ARTCUST2 Set " & vbCrLf _
            & "    SELL_CODE = R1.NEW_VALUE" & vbCrLf _
            & "    , LAST_OPER = R1.USER_ID" & vbCrLf _
            & "    , LAST_DATE = SYSDATE" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO;" & vbCrLf _
            & "   Insert into ASTAUDT1 (" & AUDIT_TRAIL_FIELDS & ")" & vbCrLf _
            & "    values " & vbCrLf _
            & "    (R1." & Replace(Replace(AUDIT_TRAIL_FIELDS, ", ", ", R1."), "R1.INIT_DATE", "SYSDATE") & ");" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(sqlR, "V", New String() {REALIGN_NO_000000})

        'Update realignment tables
        For Each TABLE As String In New String() {"ARTCUSF1", "ARTCUSF2"}
            Dim sqlT As String = "Update " & TABLE & " set REALIGN_NO = :PARM1 where REALIGN_NO = :PARM2"
            ASCDATA1.ExecuteSQL(sqlT, "VV", New String() {REALIGN_NO_archived, REALIGN_NO_000000})
        Next


        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


#Region "Excel Upload"

    Sub XLS_Upload()
        importErrors.Clear()
        chkImportErrors.Checked = True
        splImportErrors.Panel2Collapsed = True
        updatedByImport = 0
        Dim matchedButNotChanged As Integer = 0

        dst.Tables("ARTCUSIE").Rows.Clear()


        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing
            Dim firstDataRow As Integer = 0

            For n As Integer = 0 To oSheet.UsedRange.RowCount
                If (oSheet.Cells(n, 0).Value & "" = "Customer") Then
                    firstDataRow = n + 1
                    Exit For
                End If
            Next
            Dim rowARTCUST1 As DataRow = Nothing
            Dim r As Integer = firstDataRow
            Do While oSheet.Cells(r, 0).Value & "" <> ""
                Dim rowErrors As New List(Of String)


                Dim CUST_CODE As String = oSheet.Cells(r, 0).Value & ""
                If rowARTCUST1 IsNot Nothing Then
                    If CUST_CODE <> rowARTCUST1.Item("CUST_CODE") Then
                        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
                    End If
                Else
                    rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
                End If
                If rowARTCUST1 IsNot Nothing Then
                    'got a valid cust
                    CUST_CODE = rowARTCUST1.Item("CUST_CODE")
                    Dim CUST_STORE_NO As String = oSheet.Cells(r, 1).Value & ""
                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                    If IsNothing(rowARTCUST2) Then
                        rowErrors.Add("Invalid Customer/Store (" & CUST_CODE & ":" & CUST_STORE_NO & ")")
                    End If

                    Dim SELL_CODE As String = oSheet.Cells(r, 5).Value & ""
                    If SELL_CODE <> "" Then
                        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE)
                        If IsNothing(rowSOTSELL1) Then
                            rowErrors.Add("Invalid Future AE (" & SELL_CODE & ")")
                        End If

                        If rowErrors.Count = 0 Then
                            Dim sqlRealignTarget As String = "REALIGN_NO = '000000' and CUST_CODE = '" & CUST_CODE & "'" _
                                                             & " and CUST_STORE_NO = '" & CUST_STORE_NO & "' and CUST_STORE_STATUS = 'A'"

                            Dim rows() As DataRow = dst.Tables("ARTCUSF2").Select(sqlRealignTarget)

                            If rows.Length = 1 Then
                                Dim aeCurr As String = rows(0).Item("SELL_CODE") & ""
                                If aeCurr <> "" Then
                                    If aeCurr <> SELL_CODE Then
                                        rows(0).Item("SELL_CODE") = SELL_CODE
                                        updatedByImport += 1
                                    Else
                                        matchedButNotChanged += 1
                                    End If
                                Else
                                    Dim msg As String = "NO Current AE to realign"
                                End If

                            Else
                                rowErrors.Add("No Active Re-Alignment record for Customer/Store (" & CUST_CODE & ":" & CUST_STORE_NO & ")")
                            End If
                        End If
                    Else
                        rowErrors.Add("No Future AE specified on row")
                    End If

                Else
                    If CUST_CODE <> "Totals" Then
                        rowErrors.Add("Invalid Cust Code (" & CUST_CODE & ")")
                    End If
                End If
                If rowErrors.Count > 0 Then
                    importErrors.Add(r + 1, rowErrors)
                End If
                r += 1
            Loop

            Dim msgTitle As String = "Success"
            Dim importMsg As String = "Import Complete." & vbCrLf & vbCrLf
            Dim recordsUpdatedMsg As String = "Records Updated (" & CStr(updatedByImport) & ")"
            Dim recordsUnchangedMsg As String = "Records Unchanged (" & CStr(matchedButNotChanged) & ")"
            Dim recordsErrorMsg As String = "Records Skipped due to errors (" & CStr(importErrors.Count) & ")"
            If updatedByImport > 0 Then
                importMsg &= recordsUpdatedMsg & vbCrLf
            End If
            If matchedButNotChanged > 0 Then
                importMsg &= recordsUnchangedMsg & vbCrLf
            End If
            If importErrors.Count > 0 Then
                msgTitle &= " (with errors)"
                importMsg &= vbCrLf & recordsErrorMsg
            End If

            MsgBox(importMsg, MsgBoxStyle.OkOnly, msgTitle)
            If importErrors.Count > 0 Then
                Display_Import_Errors()
            End If

        End If

    End Sub

    Sub Display_Import_Errors()

        Dim rowErrors As KeyValuePair(Of Integer, List(Of String))
        For Each rowErrors In importErrors
            ' Display Key and Value.
            Dim errorsOnRow As List(Of String) = rowErrors.Value
            Dim rowError As String
            For Each rowError In errorsOnRow
                Dim rowARTCUSIE As DataRow = dst.Tables("ARTCUSIE").NewRow
                rowARTCUSIE.Item("ROW") = rowErrors.Key
                rowARTCUSIE.Item("IMPORT_ERROR") = rowError
                dst.Tables("ARTCUSIE").Rows.Add(rowARTCUSIE)
            Next

        Next

        chkImportErrors.Visible = True
        chkImportErrors.Checked = True
        splImportErrors.Panel2Collapsed = False

    End Sub

#End Region




#Region "grdARTCUSF2"

    Private Sub grdARTCUSF2_AfterRowCancelUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdARTCUSF2.AfterRowCancelUpdate
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdARTCUSF2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSF2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

                Case "SELL_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdARTCUSF2, sql_where)
            End Select
        End With

    End Sub

#End Region

    Function Get_sql_ARTCUST2()
        Dim sqlARTCUST2 As String = "Select '" & REALIGN_NO & "' REALIGN_NO" & vbCrLf _
        & ", ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.SELL_CODE" & vbCrLf _
        & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.SELL_CODE SELL_CODE_CURR, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
        & " from ARTCUST2"
        Return sqlARTCUST2
    End Function

    Sub set_filter()
        Dim dvw As DataView = DirectCast(grdARTCUSF2.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If chkActiveOnly.Checked Then sql &= " and CUST_STORE_STATUS = 'A'"
        If chkOnlyIfAEAssigned.Checked Then sql &= " and (ISNULL(SELL_CODE,'')<>'' or ISNULL(SELL_CODE_CURR,'')<>'')"
        If chkOnlyIfAEDifferent.Checked Then sql &= " and ISNULL(SELL_CODE,'')<>ISNULL(SELL_CODE_CURR,'')"

        dvw.RowFilter = Mid(sql, 5)
    End Sub
    Private Sub chkActiveOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkActiveOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        set_filter()
    End Sub

    Private Sub chkOnlyIfAE_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlyIfAEAssigned.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        set_filter()
    End Sub

    Private Sub chkOnlyIfAEDifferent_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlyIfAEDifferent.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        set_filter()
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdARTCUSF2.ActiveRow
            Select Case COLUMN_NAME
                Case "SELL_CODE"
                    Dim SELL_CODE As String = ""
                    If Trim(.Cells("SELL_CODE").Value & "") <> "" Then
                        SELL_CODE = Validate_SELL_CODE(.Cells("SELL_CODE").Value & "")
                    End If
                    Cancel = (SELL_CODE = "")
            End Select
        End With
    End Sub

    Function Validate_SELL_CODE(SELL_CODE_Z As String) As String
        Dim EMsg As String = ""
        If SELL_CODE_z = "" Then Return ""

        Dim SELL_CODE As String = ""
        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE_Z)

        If rowSOTSELL1 Is Nothing Then
            EMsg = "AE is Not on File" & vbCrLf
        Else
            'If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
            '    EMsg = "Item Status is not Active" & vbCrLf
            'End If
        End If

        If EMsg <> "" And grdARTCUSF2.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "AE Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                SELL_CODE = rowSOTSELL1.Item(0)
            End If
        End If
        Return SELL_CODE
    End Function

    Private Sub grdARTCUSF2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSF2.BeforeRowUpdate

        If e.Row.Cells("SELL_CODE").Value & "" <> "" Then Validate_Columns("SELL_CODE", e.Cancel)
        ASCMAIN1.Progress("")
        If e.Cancel = True Then
            ASCMAIN1.Progress("Invalid AE Code - either Fix the code or ESC to Cancel the entry")
            Exit Sub
        End If

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("RA_NO").Value = RA_NO
        '    Dim RA_LNO As Int64 = Val(dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") & "") + 1
        '    e.Row.Cells("RA_LNO").Value = RA_LNO
        'End If
    End Sub

    Private Sub grdARTCUSF2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSF2.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("SELL_CODE").Value & "" <> e.Row.Cells("SELL_CODE_CURR").Value & "" Then
                e.Row.Cells("SELL_CODE").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("SELL_CODE").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub chkImportErrors_CheckedChanged(sender As Object, e As EventArgs) Handles chkImportErrors.CheckedChanged
        splImportErrors.Panel2Collapsed = Not chkImportErrors.Checked
    End Sub

    Private Sub UltraLabel1_Click(sender As Object, e As EventArgs) Handles UltraLabel1.Click

    End Sub
End Class