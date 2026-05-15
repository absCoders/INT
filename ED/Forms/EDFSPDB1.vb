Public Class EDFSPDB1

    Dim MATSGRPX As String
    Dim sqlMATSGRP1 As String
    Dim STYLE_GROUP_DESC As String

    Dim MATSGRP1 As String
    Dim GMTCGMAX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Load_MATSGRPX(True)

        With dst
            ASCMAIN1.sql = "Select * from " & MATSGRPX
            Create_TDA(.Tables.Add, "MATSGRPX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select MATSGRP1.STYLE_GROUP_DESC, GMTSTYL1.*" _
                & " from " & MATSGRP1 & " MATSGRP1,GMTSTYL1" _
                & " where GMTSTYL1.DGC_CODE = MATSGRP1.DGC_CODE" _
                & "   and GMTSTYL1.VEND_CODE = MATSGRP1.VEND_CODE" _
                & "   and GMTSTYL1.STYLE_CODE = MATSGRP1.STYLE_CODE" _
                & "   and GMTSTYL1.COLOR_CODE = MATSGRP1.COLOR_CODE"
            sqlMATSGRP1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add("MATSGRP1"), MATSGRP1, "**", 0, True, "", 5)

            'ASCMAIN1.sql = "Select * from GMTCGMA0"
            'Create_TDA(.Tables.Add, "GMTCGMAX", "**", 0, False, "", 0)
            'With .Tables("GMTCGMAX").Columns
            '    .Add("FS")
            '    .Add("LS")
            '    .Add("FR")
            '    .Add("LR")
            '    .Add("BEG", GetType(System.Int32))
            '    .Add("SLS", GetType(System.Int32))
            '    .Add("REC", GetType(System.Int32))
            '    .Add("SHP", GetType(System.Int32))
            '    .Add("WSX", GetType(System.Int32))
            '    .Add("SSX", GetType(System.Int32))
            '    .Add("ADJ", GetType(System.Int32))
            'End With

            ASCMAIN1.sql = "Select * from " & GMTCGMAX
            Create_TDA(.Tables.Add("GMTCGMAX"), GMTCGMAX, "**", 0, False, "", 0)
            With .Tables("GMTCGMAX").Columns
                .Add("BEG", GetType(System.Int32))
                .Add("SLS", GetType(System.Int32))
                .Add("REC", GetType(System.Int32))
                .Add("SHP", GetType(System.Int32))
                .Add("WSX", GetType(System.Int32))
                .Add("SSX", GetType(System.Int32))
                .Add("ADJ", GetType(System.Int32))
            End With
        End With


        Create_Relation("MATSGRPX", "MATSGRP1", "STYLE_GROUP_DESC")

        grdMATSGRPX.DataSource = dst.Tables("MATSGRPX")
        grdMATSGRP1.DataSource = dst.Tables("MATSGRP1")
        grdGMTCGMAX.DataSource = dst.Tables("GMTCGMAX")

        Create_Summary(grdMATSGRPX, "STYLE_GROUP_DESC", "Count")
        Create_Summary(grdMATSGRP1, "DGC_CODE", "Count")

        With grdGMTCGMAX.DisplayLayout.Bands(0)
            .Columns("STORE_NO").Header.Fixed = True
            .Columns("ON_HAND_UNITS").Header.Fixed = True

            Create_Summary(grdGMTCGMAX, "STORE_NO", "Count")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SLS_UNITS_TWK", "SLS_UNITS_LWK", "SLS_UNITS_2WK", "SLS_UNITS_3WK", "SLS_UNITS_4WK", "SLS_UNITS_5WK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    gcol.CellAppearance.BackGradientStyle = GradientStyle.GlassRight50
                    Create_Summary(grdGMTCGMAX, gcol.Key)
                    gcol.Width = 50

                End If
                If New String() {"BOW_UNITS_TWK", "BOW_UNITS_LWK", "BOW_UNITS_2WK", "BOW_UNITS_3WK", "BOW_UNITS_4WK", "BOW_UNITS_5WK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    Create_Summary(grdGMTCGMAX, gcol.Key)
                    gcol.Width = 50
                End If
                If New String() {"ON_HAND_UNITS"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                End If
                If New String() {"FS", "LS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.Width = 50
                End If
                If New String() {"FR", "LR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.Width = 50
                End If
                If New String() {"BEG", "SLS", "REC", "SHP", "WSX", "SSX", "ADJ"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    Create_Summary(grdGMTCGMAX, gcol.Key)
                    gcol.Width = 50

                    .Columns(gcol.Key & "S").Hidden = True
                    .Columns(gcol.Key & "M").Hidden = True
                    .Columns(gcol.Key & "W").Hidden = True
                End If
            Next
            .Columns("BOW_UNITS_TWK").Header.Caption = "BTW"
            .Columns("BOW_UNITS_LWK").Header.Caption = "BLW"
            .Columns("SLS_UNITS_TWK").Header.Caption = "STW"
            .Columns("SLS_UNITS_LWK").Header.Caption = "SLW"
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

            Case "Edit"

                If STYLE_GROUP_DESC = "" Then
                    EMsg &= vbCr & "You must first select a Style Group"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("MATSGRP1", STYLE_GROUP_DESC) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("STYLE_GROUP_DESC").Text = "{New Style Group}" Then
                    EMsg &= "Invalid Description for a Style Group"
                End If

                If STYLE_GROUP_DESC <> Absx1.txtFor("STYLE_GROUP_DESC").Text Then
                    If Absx1.txtFor("STYLE_GROUP_DESC").Text = "" Then
                        EMsg &= vbCr & "A Style Group Description is Mandatory"
                    Else
                        ASCMAIN1.sql = "Select Count (*) from MATSGRP1 where STYLE_GROUP_DESC = :PARM1"
                        Dim C As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("STYLE_GROUP_DESC").Text})
                        If C <> 0 Then
                            EMsg &= vbCr & "Style Group '" & Absx1.txtFor("STYLE_GROUP_DESC").Text & "' is already defined"
                        End If
                    End If
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

            Case "New"
                EntryMode = "N"
                STYLE_GROUP_DESC = "{New Style Group}"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done", "Cancel"
                Mode_Settings(False)

        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '   .Items("Print").Settings.Enabled = iScreenMode

                    .Items("New").Visible = Not InquiryMode
                    .Items("Edit").Visible = Not InquiryMode

                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    ' .Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                    .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                    .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                .Groups("Show by Store for").Visible = ScreenMode
                .Groups("Add SKU").Visible = ScreenMode And (EntryMode <> "V")
                .Groups("Style Group").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("STYLE_GROUP_DESC"), EntryMode = "V")

        With grdMATSGRP1.DisplayLayout.Override
            If EntryMode = "E" Then
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowDelete = DefaultableBoolean.False
            End If
        End With


        grdMATSGRPX.Visible = Not ScreenMode
        splMATSGRP1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"MATSGRP1", "GMTCGMAX", "MATSGRPX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        ASCDATA1.ExecuteSQL("Truncate Table " & GMTCGMAX)

        Load_MATSGRPX(False)
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        'Load_MATSGRPX(False)

        ASCDATA1.ExecuteSQL("Truncate Table " & MATSGRP1)
        ASCDATA1.ExecuteSQL("Insert into " & MATSGRP1 & " Select * from MATSGRP1 where STYLE_GROUP_DESC = :PARM1", "V", New Object() {STYLE_GROUP_DESC})

        EnforceConstraints(False)

        Fill_Records("MATSGRP1")
        Sort_grdColumns(grdMATSGRP1, "DGC_CODE")

        Load_GMTCGMAX()

        Setup_xTD()

        'ASCMAIN1.sql = "Select * from GMTCGMA0 where ROWNUM < 100"
        'Fill_Records("GMTCGMAX", "", , ASCMAIN1.sql)
        'Sort_grdColumns(grdGMTCGMAX, "DGC_CODE")

        'For Each rowGMTCGMAX As DataRow In dst.Tables("GMTCGMAX").Select("")
        '    Dim FS As String = CStr(rowGMTCGMAX.Item("YYYYWW_FRST_SOLD") & "").PadLeft(6, " ")
        '    Dim LS As String = CStr(rowGMTCGMAX.Item("YYYYWW_LAST_SOLD") & "").PadLeft(6, " ")
        '    Dim FR As String = CStr(rowGMTCGMAX.Item("YYYYWW_FRST_SHPD") & "").PadLeft(6, " ")
        '    Dim LR As String = CStr(rowGMTCGMAX.Item("YYYYWW_LAST_SHPD") & "").PadLeft(6, " ")
        '    rowGMTCGMAX.Item("FS") = Mid(FS, Len(FS) - 3, 4)
        '    rowGMTCGMAX.Item("LS") = Mid(LS, Len(LS) - 3, 4)
        '    rowGMTCGMAX.Item("FR") = Mid(FR, Len(FR) - 3, 4)
        '    rowGMTCGMAX.Item("LR") = Mid(LR, Len(LR) - 3, 4)
        'Next

        If EntryMode = "N" Then
            dst.Tables("MATSGRPX").Rows.Add(New Object() {STYLE_GROUP_DESC})
        End If

        EnforceConstraints(True)

        Absx1.txtFor("STYLE_GROUP_DESC").Text = STYLE_GROUP_DESC

        Setup_grdGMTCGMAX()

        'Sort_grdColumns(grdMATSGRPX, "RECORD_TYPE,REGISTER_DATE", True)
        'grdMATSGRPX.Text = "A/R Roll Forward for " & Absx1.txtFor("LEGEND").Text

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        ASCMAIN1.sql = "Delete from MATSGRP1 where STYLE_GROUP_DESC = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {STYLE_GROUP_DESC})

        If STYLE_GROUP_DESC <> Absx1.txtFor("STYLE_GROUP_DESC").Text Then

            Dim rowMATSGRPX As DataRow = dst.Tables("MATSGRPX").Rows.Find(STYLE_GROUP_DESC)
            STYLE_GROUP_DESC = Absx1.txtFor("STYLE_GROUP_DESC").Text
            rowMATSGRPX.Item("STYLE_GROUP_DESC") = STYLE_GROUP_DESC
            'For Each row As DataRow In dst.Tables("MATSGRP1").Select("")
            '    row.Item("STYLE_GROUP_DESC") = STYLE_GROUP_DESC
            'Next
        End If
        Update_Record_TDA("MATSGRP1", "1=1")

        ASCMAIN1.sql = "Insert into MATSGRP1 Select * from " & MATSGRP1
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdMATSGRPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdMATSGRP1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "SKU Inquiry", "Add Styles from Style Master", "Add Styles from Open POs")
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


        Select Case e.SourceControl.Name
            Case "grdMATSGRP1"
                e.Tool.ToolbarsManager.Tools("Add Styles from Style Master").SharedProps.Visible = (EntryMode <> "V")
                e.Tool.ToolbarsManager.Tools("Add Styles from Open POs").SharedProps.Visible = (EntryMode <> "V")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTINVHX"
            '    e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Add Styles from Style Master"

                Dim sql_where As String = ""
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LONG_SKU", , sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = True
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Loading Styles")

                        grdMATSGRP1.Visible = False
                        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows

                            Dim DGC_CODE As String = row.Item("DGC_CODE")
                            Dim VEND_CODE As String = row.Item("VEND_CODE")
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                            Dim rowMATSGRP1 As DataRow = dst.Tables("MATSGRP1").Rows.Find(New String() {STYLE_GROUP_DESC, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
                            If rowMATSGRP1 IsNot Nothing Then
                                MsgBox("Style " & DGC_CODE & "-" & VEND_CODE & "-" & STYLE_CODE & "-" & COLOR_CODE & " is already in Style Group", MsgBoxStyle.OkOnly, "Cannot Add Style")
                            Else
                                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
                            End If
                        Next
                        grdMATSGRP1.Visible = True
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                        Setup_grdGMTCGMAX()
                    End If
                End If

            Case "Add Styles from Open POs"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "SKU_NUMBER"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim SKU_NUMBER As String = Absx1.txtFor("SKU_NUMBER").Text
                    If SKU_NUMBER <> "" Then
                        Add_SKU(SKU_NUMBER)
                        Setup_grdGMTCGMAX()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SKU_NUMBER"
                Dim SKU_NUMBER As String = Absx1.txtFor("SKU_NUMBER").Text
                If SKU_NUMBER <> "" Then
                    Add_SKU(SKU_NUMBER)
                End If
        End Select
    End Sub
#End Region

    Sub Load_MATSGRPX(initialize As Boolean)
        ASCMAIN1.sql = "Select STYLE_GROUP_DESC, Count (*) LSKUS from MATSGRP1 group by STYLE_GROUP_DESC"
        If initialize Then
            MATSGRPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & MATSGRPX & " Add Primary Key (STYLE_GROUP_DESC)")

            ASCMAIN1.sql = "Select * from MATSGRP1 where ROWNUM < 1"
            MATSGRP1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & MATSGRP1 & " Add Primary Key (STYLE_GROUP_DESC,DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE)")

            ASCMAIN1.sql = "Select GMTCGMA0.*, 'XX' SEASON_SEQ_NO" _
                & ",GMTCGMA0.YYYYWW_FRST_SHPD FR,GMTCGMA0.YYYYWW_LAST_SHPD LR,GMTCGMA0.YYYYWW_FRST_SOLD FS,GMTCGMA0.YYYYWW_LAST_SOLD LS" _
                & ",GMTCGMA0.BOS_UNITS BEGS, GMTCGMA1.SLS_UNITS SLSS, GMTCGMA1.REC_UNITS RECS, GMTCGMA1.SHP_UNITS SHPS, GMTCGMA1.WSX_UNITS WSXS, GMTCGMA1.SSX_UNITS SSXS, GMTCGMA1.ADJ_UNITS ADJS" _
                & ",GMTCGMA0.BOM_UNITS BEGM, GMTCGMA2.SLS_UNITS SLSM, GMTCGMA2.REC_UNITS RECM, GMTCGMA2.SHP_UNITS SHPM, GMTCGMA2.WSX_UNITS WSXM, GMTCGMA2.SSX_UNITS SSXM, GMTCGMA2.ADJ_UNITS ADJM" _
                & ",GMTCGMA0.BOW_UNITS BEGW, GMTCGMA3.SLS_UNITS SLSW, GMTCGMA3.REC_UNITS RECW, GMTCGMA3.SHP_UNITS SHPW, GMTCGMA3.WSX_UNITS WSXW, GMTCGMA3.SSX_UNITS SSXW, GMTCGMA3.ADJ_UNITS ADJW" _
                & " from GMTCGMA0,GMTCGMA1,GMTCGMA2,GMTCGMA3 where ROWNUM < 1"
            GMTCGMAX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GMTCGMAX & " Add Primary Key (DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE,STORE_NO)")

        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & MATSGRPX)
            ASCDATA1.ExecuteSQL("Insert into " & MATSGRPX & " " & ASCMAIN1.sql)

            EnforceConstraints(False)
            Fill_Records("MATSGRPX")

            ASCMAIN1.sql = Replace(sqlMATSGRP1, MATSGRP1 & " MATSGRP1", "MATSGRP1")
            Fill_Records("MATSGRP1", "", , ASCMAIN1.sql)

            EnforceConstraints(True)
            Sort_grdColumns(grdMATSGRPX, "STYLE_GROUP_DESC")
        End If
    End Sub

    Private Sub grdMATSGRPX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdMATSGRPX.DoubleClickRow
        STYLE_GROUP_DESC = e.Row.Cells("STYLE_GROUP_DESC").Value
        Click_Command("View")
    End Sub

    Sub Add_SKU(SKU_NUMBER As String)
        Dim rowGMTSKUF1 As DataRow = LookUp("GMTSKUF1", SKU_NUMBER)
        If rowGMTSKUF1 Is Nothing Then
            MsgBox("Invalid Value Specified for SKU (" & SKU_NUMBER & ")", MsgBoxStyle.OkOnly, "Cannot Add SKU")
        Else
            Dim DGC_CODE As String = rowGMTSKUF1.Item("DGC_CODE")
            Dim VEND_CODE As String = rowGMTSKUF1.Item("VEND_CODE")
            Dim STYLE_CODE As String = rowGMTSKUF1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowGMTSKUF1.Item("COLOR_CODE")

            Dim rowMATSGRP1 As DataRow = dst.Tables("MATSGRP1").Rows.Find(New String() {STYLE_GROUP_DESC, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
            If rowMATSGRP1 IsNot Nothing Then
                MsgBox("SKU " & SKU_NUMBER & " is already in Style Group", MsgBoxStyle.OkOnly, "Cannot Add SKU")
            Else
                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
            End If
            Absx1.txtFor("SKU_NUMBER").Text = ""
        End If
        Application.DoEvents()
        Absx1.txtFor("SKU_NUMBER").Focus()
    End Sub

    Sub Add_Style(DGC_CODE As String, VEND_CODE As String, STYLE_CODE As String, COLOR_CODE As String)
        Dim rowGMTSTYL1 As DataRow = LookUp("GMTSTYL1", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
        Dim rowMATSGRP1 As DataRow = dst.Tables("MATSGRP1").NewRow

        For Each dcol As DataColumn In dst.Tables("MATSGRP1").Columns
            If dcol.ColumnName = "STYLE_GROUP_DESC" Then
                rowMATSGRP1.Item("STYLE_GROUP_DESC") = STYLE_GROUP_DESC
            Else
                rowMATSGRP1.Item(dcol.ColumnName) = rowGMTSTYL1.Item(dcol.ColumnName)
            End If
        Next

        dst.Tables("MATSGRP1").Rows.Add(rowMATSGRP1)

        Update_Record_TDA("MATSGRP1")
        Load_GMTCGMAX_for_LSKU(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
    End Sub

    Private Sub optIC_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optIC.ValueChanged
        Setup_grdGMTCGMAX()
    End Sub

    Private Sub optWMS_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optWMS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_xTD()
        Setup_grdGMTCGMAX()
    End Sub

    Sub Setup_xTD()
        With dst.Tables("GMTCGMAX")
            For Each COL As String In New String() {"BEG", "SLS", "REC", "SHP", "WSX", "SSX", "ADJ"}
                .Columns(COL).Expression = COL & optWMS.Value
            Next
        End With
    End Sub

    Sub Render_Store_Activity()

    End Sub

    Sub Load_GMTCGMAX()

        For Each rowMATSGRP1 As DataRow In dst.Tables("MATSGRP1").Select("")
            Dim DGC_CODE As String = rowMATSGRP1.Item("DGC_CODE")
            Dim VEND_CODE As String = rowMATSGRP1.Item("VEND_CODE")
            Dim STYLE_CODE As String = rowMATSGRP1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowMATSGRP1.Item("COLOR_CODE")
            Load_GMTCGMAX_for_LSKU(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
        Next
    End Sub

    Sub Load_GMTCGMAX_for_LSKU(DGC_CODE As String, VEND_CODE As String, STYLE_CODE As String, COLOR_CODE As String)

        ASCMAIN1.sql = "Delete from " & GMTCGMAX & " where DGC_CODE = :PARM1 and VEND_CODE = :PARM2 and STYLE_CODE = :PARM3 and COLOR_CODE = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})

        Dim sqlj As String = "" _
                             & " and GMTCGMAX.DGC_CODE (+) = X.DGC_CODE" & vbCrLf _
                             & " and GMTCGMAX.VEND_CODE (+) = X.VEND_CODE" & vbCrLf _
                             & " and GMTCGMAX.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                             & " and GMTCGMAX.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                             & " and GMTCGMAX.STORE_NO (+) = X.STORE_NO" & vbCrLf

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
             & ",X.YYYYWW_FRST_SHPD FR,X.YYYYWW_LAST_SHPD LR,X.YYYYWW_FRST_SOLD FS,X.YYYYWW_LAST_SOLD LS" & vbCrLf _
             & ",X.BOS_UNITS BEGS, GMTCGMA1.SLS_UNITS SLSS, GMTCGMA1.REC_UNITS RECS, GMTCGMA1.SHP_UNITS SHPS, GMTCGMA1.WSX_UNITS WSXS, GMTCGMA1.SSX_UNITS SSXS, GMTCGMA1.ADJ_UNITS ADJS" & vbCrLf _
             & ",X.BOM_UNITS BEGM, GMTCGMA2.SLS_UNITS SLSM, GMTCGMA2.REC_UNITS RECM, GMTCGMA2.SHP_UNITS SHPM, GMTCGMA2.WSX_UNITS WSXM, GMTCGMA2.SSX_UNITS SSXM, GMTCGMA2.ADJ_UNITS ADJM" & vbCrLf _
             & ",X.BOW_UNITS BEGW, GMTCGMA3.SLS_UNITS SLSW, GMTCGMA3.REC_UNITS RECW, GMTCGMA3.SHP_UNITS SHPW, GMTCGMA3.WSX_UNITS WSXW, GMTCGMA3.SSX_UNITS SSXW, GMTCGMA3.ADJ_UNITS ADJW" & vbCrLf _
             & " from (Select GMTCGMA0.*,GMTSEAS1.SEASON_SEQ_NO from GMTCGMA0,GMTSEAS1,GMTSTYL1" & vbCrLf _
             & " where GMTCGMA0.DGC_CODE = :PARM1 and GMTCGMA0.VEND_CODE = :PARM2 and GMTCGMA0.STYLE_CODE = :PARM3 and GMTCGMA0.COLOR_CODE = :PARM4" & vbCrLf _
             & "   and GMTSTYL1.DGC_CODE = GMTCGMA0.DGC_CODE and GMTSTYL1.VEND_CODE = GMTCGMA0.VEND_CODE and GMTSTYL1.STYLE_CODE = GMTCGMA0.STYLE_CODE and GMTSTYL1.COLOR_CODE = GMTCGMA0.COLOR_CODE" & vbCrLf _
             & "   and GMTSEAS1.SEASON_CODE = GMTSTYL1.SEASON_CODE) X" & vbCrLf _
             & ",GMTCGMA1,GMTCGMA2,GMTCGMA3" & vbCrLf _
             & " where GMTCGMA1.SEASON_SEQ_NO (+) = X.SEASON_SEQ_NO" & vbCrLf _
             & Replace(sqlj, "GMTCGMAX", "GMTCGMA1") & Replace(sqlj, "GMTCGMAX", "GMTCGMA2") & Replace(sqlj, "GMTCGMAX", "GMTCGMA3")

        ASCMAIN1.sql = "Insert into " & GMTCGMAX & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})

    End Sub

    Sub Setup_grdGMTCGMAX()

        If grdMATSGRP1.ActiveRow Is Nothing OrElse Not grdMATSGRP1.ActiveRow.IsDataRow Then
            grdGMTCGMAX.Visible = False
        Else
            grdGMTCGMAX.Visible = True

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Sales & On Hand Data")

            Dim DGC_CODE As String = grdMATSGRP1.ActiveRow.Cells("DGC_CODE").Value
            Dim VEND_CODE As String = grdMATSGRP1.ActiveRow.Cells("VEND_CODE").Value
            Dim STYLE_CODE As String = grdMATSGRP1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdMATSGRP1.ActiveRow.Cells("COLOR_CODE").Value

            If optIC.Value = "A" Then
                grdGMTCGMAX.Text = "Sales & On Hand - All Styles in Group Combined"
                ASCMAIN1.sql = "Select '*' DGC_CODE, '*' VEND_CODE, '*' STYLE_CODE, '*' COLOR_CODE, STORE_NO" _
                    & ", SUM (SLS_UNITS_TWK) SLS_UNITS_TWK" _
                    & ", SUM (SLS_UNITS_LWK) SLS_UNITS_1WK" _
                    & ", SUM (SLS_UNITS_2WK) SLS_UNITS_2WK" _
                    & ", SUM (SLS_UNITS_3WK) SLS_UNITS_3WK" _
                    & ", SUM (SLS_UNITS_4WK) SLS_UNITS_4WK" _
                    & ", SUM (SLS_UNITS_5WK) SLS_UNITS_5WK" _
                    & ", SUM (BOW_UNITS_TWK) BOW_UNITS_TWK" _
                    & ", SUM (BOW_UNITS_LWK) BOW_UNITS_LWK" _
                    & ", SUM (BOW_UNITS_2WK) BOW_UNITS_2WK" _
                    & ", SUM (BOW_UNITS_3WK) BOW_UNITS_3WK" _
                    & ", SUM (BOW_UNITS_4WK) BOW_UNITS_4WK" _
                    & ", SUM (BOW_UNITS_5WK) BOW_UNITS_5WK" _
                    & ", SUM (SHP_UNITS_TWK) SHP_UNITS_TWK" _
                    & ", SUM (SHP_UNITS_LWK) SHP_UNITS_LWK" _
                    & ", SUM (SHP_UNITS_2WK) SHP_UNITS_2WK" _
                    & ", SUM (SHP_UNITS_3WK) SHP_UNITS_3WK" _
                    & ", SUM (SHP_UNITS_4WK) SHP_UNITS_4WK" _
                    & ", SUM (SHP_UNITS_5WK) SHP_UNITS_5WK" _
                    & ", MIN (YYYYWW_FRST_SOLD) YYYYWW_FRST_SOLD" _
                    & ", MAX (YYYYWW_LAST_SOLD) YYYYWW_LAST_SOLD" _
                    & ", MIN (YYYYWW_FRST_SHPD) YYYYWW_FRST_SHPD" _
                    & ", MAX (YYYYWW_LAST_SHPD) YYYYWW_LAST_SHPD" _
                    & ", MIN (YYYYWW_FRST_LDDR) YYYYWW_FRST_LDDR" _
                    & ", MAX (YYYYWW_LAST_LDDR) YYYYWW_LAST_LDDR" _
                    & ", MIN (YYYYWW_FRST_LDDX) YYYYWW_FRST_LDDX" _
                    & ", MAX (YYYYWW_LAST_LDDX) YYYYWW_LAST_LDDX" _
                    & ", MIN (DATE_FRST_SOLD) DATE_FRST_SOLD" _
                    & ", MAX (DATE_LAST_SOLD) DATE_LAST_SOLD" _
                    & ", MIN (DATE_FRST_SHPD) DATE_FRST_SHPD" _
                    & ", MAX (DATE_LAST_SHPD) DATE_LAST_SHPD" _
                    & ", MIN (DATE_FRST_LDDR) DATE_FRST_LDDR" _
                    & ", MAX (DATE_LAST_LDDR) DATE_LAST_LDDR" _
                    & ", MIN (DATE_FRST_LDDX) DATE_FRST_LDDX" _
                    & ", MAX (DATE_LAST_LDDX) DATE_LAST_LDDX" _
                    & ", SUM (ON_HAND_UNITS) ON_HAND_UNITS" _
                    & ", SUM (ON_HAND_RETL) ON_HAND_RETL" _
                    & ", SUM (ON_HAND_COST) ON_HAND_COST" _
                    & ", SUM (BOS_UNITS) BOS_UNITS" _
                    & ", SUM (BOS_RETL) BOS_RETL" _
                    & ", SUM (BOS_COST) BOS_COST" _
                    & ", SUM (BOM_UNITS) BOM_UNITS" _
                    & ", SUM (BOM_RETL) BOM_RETL" _
                    & ", SUM (BOM_COST) BOM_COST" _
                    & ", SUM (BOW_UNITS) BOW_UNITS" _
                    & ", SUM (BOW_RETL) BOW_RETL" _
                    & ", SUM (BOW_COST) BOW_COST" _
                    & ", MAX (SEASON_SEQ_NO) SEASON_SEQ_NO" _
                    & ", SUM (BEGS) BEGS" _
                    & ", SUM (SLSS) SLSS" _
                    & ", SUM (RECS) RECS" _
                    & ", SUM (SHPS) SHPS" _
                    & ", SUM (WSXS) WSXS" _
                    & ", SUM (SSXS) SSXS" _
                    & ", SUM (ADJS) ADJS" _
                    & ", SUM (BEGM) BEGM" _
                    & ", SUM (SLSM) SLSM" _
                    & ", SUM (RECM) RECM" _
                    & ", SUM (SHPM) SHPM" _
                    & ", SUM (WSXM) WSXM" _
                    & ", SUM (SSXM) SSXM" _
                    & ", SUM (ADJM) ADJM" _
                    & ", SUM (BEGW) BEGW" _
                    & ", SUM (SLSW) SLSW" _
                    & ", SUM (RECW) RECW" _
                    & ", SUM (SHPW) SHPW" _
                    & ", SUM (WSXW) WSXW" _
                    & ", SUM (SSXW) SSXW" _
                    & ", SUM (ADJW) ADJW" _
                    & " from " & GMTCGMAX _
                    & " group by STORE_NO"
            Else
                grdGMTCGMAX.Text = "Sales & On Hand - DGC " & DGC_CODE & ", Vendor " & VEND_CODE & ", Style " & STYLE_CODE & ", Color " & COLOR_CODE
                ASCMAIN1.sql = "Select * from " & GMTCGMAX _
                    & " where DGC_CODE = '" & DGC_CODE & "'" _
                    & "   and VEND_CODE = '" & VEND_CODE & "'" _
                    & "   and STYLE_CODE = '" & STYLE_CODE & "'" _
                    & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            End If

            Fill_Records("GMTCGMAX", "", , ASCMAIN1.sql)
            For Each rowGMTCGMAX As DataRow In dst.Tables("GMTCGMAX").Select("")
                Dim FS As String = CStr(rowGMTCGMAX.Item("YYYYWW_FRST_SOLD") & "").PadLeft(6, " ")
                Dim LS As String = CStr(rowGMTCGMAX.Item("YYYYWW_LAST_SOLD") & "").PadLeft(6, " ")
                Dim FR As String = CStr(rowGMTCGMAX.Item("YYYYWW_FRST_SHPD") & "").PadLeft(6, " ")
                Dim LR As String = CStr(rowGMTCGMAX.Item("YYYYWW_LAST_SHPD") & "").PadLeft(6, " ")
                rowGMTCGMAX.Item("FS") = Mid(FS, Len(FS) - 3, 4)
                rowGMTCGMAX.Item("LS") = Mid(LS, Len(LS) - 3, 4)
                rowGMTCGMAX.Item("FR") = Mid(FR, Len(FR) - 3, 4)
                rowGMTCGMAX.Item("LR") = Mid(LR, Len(LR) - 3, 4)
            Next
            Sort_grdColumns(grdGMTCGMAX, "STORE_NO")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Private Sub grdMATSGRP1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdMATSGRP1.AfterRowActivate
        Setup_grdGMTCGMAX()
    End Sub
End Class