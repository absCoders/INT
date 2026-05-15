Public Class GMFSCOL1

    Dim GMTSCOLX As String
    Dim sqlGMTSCOL2 As String
    Dim STYLE_COLLECTION_CODE As String

    Dim POTORDRX As String

    Dim GMTSCOL2 As String
    Dim GMTCGMAX As String
    Dim rowGMTSEAS1_CURR As DataRow ' SHOULD BE IN A CENTRALIZED CLASS

    Structure LSKU
        Dim DGC_CODE As String
        Dim VEND_CODE As String
        Dim STYLE_CODE As String
        Dim COLOR_CODE As String
    End Structure

    Dim WHSEs As New List(Of String)
    Dim LSKUs As New List(Of LSKU)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Load_GMTSCOLX(True)

        With dst
            ASCMAIN1.sql = "Select * from " & GMTSCOLX
            Create_TDA(.Tables.Add, "GMTSCOLX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & POTORDRX & " POTORDRX" & vbCrLf _
                & " where POTORDRX.DGC_CODE    = :PARM1" & vbCrLf _
                & "   and POTORDRX.VEND_CODE   = :PARM2" & vbCrLf _
                & "   and POTORDRX.STYLE_CODE  = :PARM3" & vbCrLf _
                & "   and POTORDRX.COLOR_CODE  = :PARM4"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "VVVV", 0)

            Create_TDA(.Tables.Add, "GMTSCOL1", "*")

            ASCMAIN1.sql = "Select GMTSCOL2.STYLE_COLLECTION_CODE, GMTSTYL1.*" _
                & " from " & GMTSCOL2 & " GMTSCOL2,GMTSTYL1" _
                & " where GMTSTYL1.DGC_CODE   = GMTSCOL2.DGC_CODE" _
                & "   and GMTSTYL1.VEND_CODE   = GMTSCOL2.VEND_CODE" _
                & "   and GMTSTYL1.STYLE_CODE   = GMTSCOL2.STYLE_CODE" _
                & "   and GMTSTYL1.COLOR_CODE   = GMTSCOL2.COLOR_CODE"
            sqlGMTSCOL2 = ASCMAIN1.sql
            Create_TDA(.Tables.Add("GMTSCOL2"), GMTSCOL2, "**", 0, True, "", 5)
            With .Tables("GMTSCOL2").Columns
                .Add("OH_STR", GetType(System.Int64))
                .Add("OH_WHS", GetType(System.Int64))
                .Add("QTY_OPN_1", GetType(System.Int64))
                .Add("QTY_OPN_2", GetType(System.Int64))
                .Add("QTY_ORD", GetType(System.Int64))
                .Add("QTY_REC", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select * from GMTSCOL3 where STYLE_COLLECTION_CODE = :PARM1"
            Create_TDA(.Tables.Add, "GMTSCOL3", "**", 0, True, "V")
            With .Tables("GMTSCOL3").Columns
                .Add("CRITERIA_SEQ", GetType(System.Int32))
                .Add("STYLE_COUNTER", GetType(System.Int32))
            End With

            With .Tables.Add("GMTSCOL3_STYLES")
                .Columns.Add("CRITERIA_SEQ", GetType(System.Int32))
                .Columns.Add("DGC_CODE")
                .Columns.Add("VEND_CODE")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("DESCRIPTION")
                .Columns.Add("SEASON_CODE")
                .Columns.Add("OP_GRP_CODE")
            End With

            Create_Relation("GMTSCOL3", "GMTSCOL3_STYLES", "CRITERIA_SEQ")

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
                .Add("BEG", GetType(System.Int64))
                .Add("SLS", GetType(System.Int64))
                .Add("REC", GetType(System.Int64))
                .Add("SHP", GetType(System.Int64))
                .Add("WSX", GetType(System.Int64))
                .Add("SSX", GetType(System.Int64))
                .Add("ADJ", GetType(System.Int64))
                .Add("QTY_OPN_1", GetType(System.Int64))
                .Add("QTY_OPN_2", GetType(System.Int64))
                .Add("SSR", GetType(System.Decimal), "IIF(ISNULL(SLS_UNITS_TWK,0)=0,0,ISNULL(BOW_UNITS_TWK,0)/ISNULL(SLS_UNITS_TWK,0))")
            End With
            With .Tables("GMTCGMAX")
                .PrimaryKey = New DataColumn() {.Columns("DGC_CODE"), .Columns("VEND_CODE"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("STORE_NO")}
            End With






            ASCMAIN1.sql = "Select * from " & GMTCGMAX & " where STORE_NO = :PARM1"
            Create_TDA(.Tables.Add, "GMTCGMA0", "**", 0, False, "V", 0)
            With .Tables("GMTCGMA0").Columns
                .Add("BEG", GetType(System.Int64))
                .Add("SLS", GetType(System.Int64))
                .Add("REC", GetType(System.Int64))
                .Add("SHP", GetType(System.Int64))
                .Add("WSX", GetType(System.Int64))
                .Add("SSX", GetType(System.Int64))
                .Add("ADJ", GetType(System.Int64))
                .Add("QTY_OPN_1", GetType(System.Int64))
                .Add("QTY_OPN_2", GetType(System.Int64))
                .Add("SSR", GetType(System.Decimal), "IIF(ISNULL(SLS_UNITS_TWK,0)=0,0,ISNULL(BOW_UNITS_TWK,0)/ISNULL(SLS_UNITS_TWK,0))")
                .Add("LSKU", GetType(System.String), "DGC_CODE + ':' + VEND_CODE + ':' + STYLE_CODE + ':' + COLOR_CODE")
            End With
            With .Tables("GMTCGMA0")
                .PrimaryKey = New DataColumn() {.Columns("STORE_NO"), .Columns("DGC_CODE"), .Columns("VEND_CODE"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With


            ASCMAIN1.sql = "Select * from GMTSTOR1"
            Create_TDA(.Tables.Add, "GMTSTOR1", "**", 0, False)

            ASCMAIN1.sql = "Select * from GMTSEAS1"
            Create_TDA(.Tables.Add, "GMTSEAS1", "**", 0, False)

            'Dim t As DataTable = .Tables("GMTCGMAX").Clone
            't.TableName = "GMTCGMA0"
            't.Columns.Add("LSKU", GetType(System.String), "DGC_CODE + ':' + VEND_CODE + ':' + STYLE_CODE + ':' + COLOR_CODE")
            '.Tables.Add(t)

            Create_Relation("GMTCGMAX", "GMTCGMA0", "STORE_NO")

        End With

        Fill_Records("GMTSTOR1")
        Fill_Records("GMTSEAS1")

        Create_Relation("GMTSCOLX", "GMTSCOL2", "STYLE_COLLECTION_CODE")

        grdGMTSCOLX.DataSource = dst.Tables("GMTSCOLX")
        grdGMTSCOL2.DataSource = dst.Tables("GMTSCOL2")
        grdGMTSCOL3.DataSource = dst.Tables("GMTSCOL3")
        grdGMTCGMAX.DataSource = dst.Tables("GMTCGMAX")

        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")

        Create_Summary(grdGMTSCOLX, "STYLE_COLLECTION_CODE", "Count")

        Create_Summary(grdGMTSCOL2, "DGC_CODE", "Count")
        Create_Summary(grdGMTSCOL2, New String() {"OH_STR", "OH_WHS", "QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC"})

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"QTY_ORD", "QTY_OPN", "QTY_REC"})

        With grdGMTCGMAX.DisplayLayout.Bands(1)
            .Override.CellAppearance.BackColor = Drawing.Color.AliceBlue
            .Columns("LSKU").Header.Fixed = True
            .Columns("ON_HAND_UNITS").Header.Fixed = True
            .Columns("SSR").Format = "#0.0" 
        End With

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
                If New String() {"SSR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    gcol.Format = "#0.0"
                    Create_Summary(grdGMTCGMAX, gcol.Key, "Custom")
                End If
                If New String() {"ON_HAND_UNITS"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                    Create_Summary(grdGMTCGMAX, gcol.Key)
                End If
                If New String() {"QTY_OPN_1", "QTY_OPN_2"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    Create_Summary(grdGMTCGMAX, gcol.Key)
                End If
                If New String() {"FS", "LS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.Width = 50
                    With grdGMTCGMAX.DisplayLayout.Bands(1).Columns(gcol.Key)
                        .CellAppearance.TextHAlign = HAlign.Center
                        .Header.Appearance.TextHAlign = HAlign.Center
                    End With
                End If
                If New String() {"FR", "LR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.Width = 50
                    With grdGMTCGMAX.DisplayLayout.Bands(1).Columns(gcol.Key)
                        .CellAppearance.TextHAlign = HAlign.Center
                        .Header.Appearance.TextHAlign = HAlign.Center
                    End With
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

        With grdGMTSCOL2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Width = 50
                End If
                If New String() {"OH_STR", "OH_WHS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 50
                End If
            Next
        End With

        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"QTY_ORD", "QTY_OPN", "QTY_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Width = 50
                End If
            Next
        End With

        ASCMAIN1.sql = "Select STORE_NO from GMTSTOR1 where STORE_TYPE_CODE = 'W'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim STORE_NO As String = row.Item("STORE_NO")
            WHSEs.Add(STORE_NO)
        Next


        Show_Filter(grdGMTSCOLX, True)

        ASCMAIN1.sql = "SELECT * FROM GMTSEAS1 WHERE SEASON_SEQ_NO = (SELECT GM_PARM_CURR_SN FROM GMTPARM1)"
        rowGMTSEAS1_CURR = ASCDATA1.GetDataRow
        chkSN_C.Text = rowGMTSEAS1_CURR.Item("SEASON_CODE")
        chkSN_C.Tag = rowGMTSEAS1_CURR.Item("SEASON_SEQ_NO")
        Dim rowSN As DataRow
        Dim CS As Integer = Val(rowGMTSEAS1_CURR.Item("SEASON_SEQ_NO"))
        rowSN = ASCDATA1.GetDataRow("Select * from GMTSEAS1 where SEASON_SEQ_NO = '" & Format(CS - 1, "00") & "'")
        chkSN_CM1.Text = rowSN.Item("SEASON_CODE")
        chkSN_CM1.Tag = rowSN.Item("SEASON_SEQ_NO")
        rowSN = ASCDATA1.GetDataRow("Select * from GMTSEAS1 where SEASON_SEQ_NO = '" & Format(CS + 1, "00") & "'")
        chkSN_CP1.Text = rowSN.Item("SEASON_CODE")
        chkSN_CP1.Tag = rowSN.Item("SEASON_SEQ_NO")
        rowSN = ASCDATA1.GetDataRow("Select * from GMTSEAS1 where SEASON_SEQ_NO = '" & Format(CS + 2, "00") & "'")
        chkSN_CP2.Text = rowSN.Item("SEASON_CODE")
        chkSN_CP2.Tag = rowSN.Item("SEASON_SEQ_NO")

        'ASCMAIN1.Add_Value_List(grdGMTCGMAX, "SEASON_SEQ_NO", "SELECT SEASON_SEQ_NO, SEASON_CODE FROM GMTSEAS1 WHERE SEASON_ACTIVE = '1'")
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "SEASON_SEQ_NO", "SELECT SEASON_SEQ_NO, SEASON_CODE FROM GMTSEAS1 WHERE SEASON_ACTIVE = '1'")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("STYLE_COLLECTION_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Supply a Description"
                End If

            Case "View"

                If STYLE_COLLECTION_CODE = "" Then
                    EMsg &= vbCr & "Invalid Value for Style Collection Code"
                End If

            Case "Edit"

                If STYLE_COLLECTION_CODE = "" Then
                    EMsg &= vbCr & "You must first select a Style Collection"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("GMTSCOL2", STYLE_COLLECTION_CODE) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("STYLE_COLLECTION_DESC").Text = "{New Style Collection}" Then
                    EMsg &= "Invalid Description for a Style Collection"
                End If

                If Absx1.txtFor("STYLE_COLLECTION_DESC").Text = "" Then
                    EMsg &= vbCr & "A Style Collection Description is Mandatory"
                Else
                    'ASCMAIN1.sql = "Select Count (*) from GMTSCOL2 where STYLE_COLLECTION_CODE = :PARM1"
                    'Dim C As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("STYLE_COLLECTION_CODE").Text})
                    'If C <> 0 Then
                    '    EMsg &= vbCr & "Style Collection '" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text & "' is already defined"
                    'End If
                End If

            Case "Cancel"

                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                         "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If MsgBox("Do you really want to Delete the entire Style Collection?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

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

            Case "Print"
                Print_Record()

            Case "Excel"
                Gembox_Excel_Export(New UltraWinGrid.UltraGrid() {grdGMTSCOL2, grdGMTCGMAX, grdPOTORDRX})

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
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

                    .Items("Excel").Visible = ScreenMode
                    .Items("New").Visible = Not InquiryMode
                    .Items("Edit").Visible = Not InquiryMode

                    .Items("Print").Visible = ScreenMode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    ' .Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                    .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                    .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                .Groups("Show by Store for").Visible = ScreenMode
                .Groups("Add SKU").Visible = ScreenMode And (EntryMode <> "V")
                '.Groups("Style Group").Visible = ScreenMode

                If Not ScreenMode Then
                    UltraExplorerBar1.Groups("Style Image").Visible = False
                End If
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("STYLE_COLLECTION_DESC"), EntryMode = "V")
        Set_Read_Only_for_ctl(Absx1.txtFor("VEND_BUYER_CODE"), EntryMode = "V")

        lblVEND_BUYER_CODE.Visible = ScreenMode
        Absx1.txtFor("VEND_BUYER_CODE").Visible = ScreenMode

        With grdGMTSCOL2.DisplayLayout.Override
            If EntryMode = "E" Or EntryMode = "N" Then
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        With grdGMTSCOL3.DisplayLayout.Override
            If EntryMode = "E" Or EntryMode = "N" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        grdGMTSCOLX.Visible = Not ScreenMode
        splGMTSCOL2.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GMTSCOL1", "GMTSCOL2", "GMTSCOL3", "GMTSCOL3_STYLES", "GMTCGMAX", "GMTCGMA0", "GMTSCOLX", "POTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optIC.Tag = ""
        ASCDATA1.ExecuteSQL("Truncate Table " & GMTCGMAX)

        Load_GMTSCOLX(False)
    End Sub

    Sub Print_Record()

        'Dim RPT_TITLE As String = "Style Collection Status & Activity"
        'Dim reportFile As String = "GMRSCOL1"

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", STYLE_COLLECTION_CODE & "-" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text)
        CR_params.Add("XTD", optWMS.Text)
        CR_params.Add("IC", optIC.Value)
        CR_params.Add("PO_WEEKS", Absx1.numFor("PO_WEEKS").Value)
        Generate_Report("GMRSCOL1", "Style Collection Members")

        'If optIC.Value = "A" Then
        '    CR_params.Add("SUBT", STYLE_COLLECTION_CODE & "-" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text)
        '    CR_params.Add("XTD", optWMS.Text)
        '    Generate_Report("GMRSCOL2", "Style Collection Status & Activity")
        'End If

        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then

            Dim rowGMTSCOL1 As DataRow = dst.Tables("GMTSCOL1").NewRow
            STYLE_COLLECTION_CODE = ASCMAIN1.Next_Control_No("GMTSCOL1.STYLE_COLLECTION_CODE")
            rowGMTSCOL1.Item("STYLE_COLLECTION_CODE") = STYLE_COLLECTION_CODE
            rowGMTSCOL1.Item("STYLE_COLLECTION_DESC") = HFs("STYLE_COLLECTION_DESC")
            rowGMTSCOL1.Item("INIT_DATE") = DATETIME_STAMP
            rowGMTSCOL1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("GMTSCOL1").Rows.Add(rowGMTSCOL1)

            dst.Tables("GMTSCOLX").Rows.Add(New Object() {STYLE_COLLECTION_CODE})

            ASCMAIN1.sql = "Delete from " & GMTSCOL2
            ASCDATA1.ExecuteSQL()
        Else
            Fill_Records("GMTSCOL1", STYLE_COLLECTION_CODE)
            ASCDATA1.ExecuteSQL("Truncate Table " & GMTSCOL2)
            ASCDATA1.ExecuteSQL("Insert into " & GMTSCOL2 & " Select * from GMTSCOL2 where STYLE_COLLECTION_CODE = :PARM1", "V", New Object() {STYLE_COLLECTION_CODE})
        End If

        Fill_Records("GMTSCOL2")
        Sort_grdColumns(grdGMTSCOL2, "DGC_CODE")

        Fill_Records("GMTSCOL3", STYLE_COLLECTION_CODE)
        Dim CRITERIA_SEQ As Int32 = 0
        For Each rowGMTSCOL3 As DataRow In dst.Tables("GMTSCOL3").Select("")
            CRITERIA_SEQ += 1
            rowGMTSCOL3.Item("CRITERIA_SEQ") = CRITERIA_SEQ
        Next

        ASCDATA1.ExecuteSQL("Delete from " & GMTCGMAX)
        ASCDATA1.ExecuteSQL("Delete from " & POTORDRX)

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

        EnforceConstraints(True)

        ' Absx1.txtFor("STYLE_COLLECTION_DESC").Text = STYLE_COLLECTION_DESC
        optIC.Tag = ""
        Setup_grdGMTCGMAX()
        Setup_Panes()
        grdGMTSCOL2.Text = "Collection " & STYLE_COLLECTION_CODE & ":" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text & "- Member Styles"
        'Sort_grdColumns(grdGMTSCOLX, "RECORD_TYPE,REGISTER_DATE", True)
        'grdGMTSCOLX.Text = "A/R Roll Forward for " & Absx1.txtFor("LEGEND").Text

        Find_New_Styles()

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
        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"GMTSCOL1", "GMTSCOL2", "GMTSCOL3"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        BeginTrans()

        ASCMAIN1.sql = "Delete from MATSGRP1 where STYLE_GROUP_DESC = " _
            & " (Select STYLE_COLLECTION_DESC from GMTSCOL1 where STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "')"
        ASCDATA1.ExecuteSQL()


        Update_Record_TDA("GMTSCOL1")
        Update_Record_TDA("GMTSCOL2", "1=1")

        ASCMAIN1.sql = "Delete from GMTSCOL2 where STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GMTSCOL2 Select * from " & GMTSCOL2
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Delete from GMTSCOL3 where STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'"
        'ASCDATA1.ExecuteSQL()
        Update_Record_TDA("GMTSCOL3", "STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'")

        ASCMAIN1.sql = "Insert into MATSGRP1" _
            & " Select GMTSCOL1.STYLE_COLLECTION_DESC, GMTSCOL2.DGC_CODE, GMTSCOL2.VEND_CODE, GMTSCOL2.STYLE_CODE, GMTSCOL2.COLOR_CODE" _
            & " from GMTSCOL1,GMTSCOL2 where GMTSCOL1.STYLE_COLLECTION_CODE = GMTSCOL2.STYLE_COLLECTION_CODE" _
            & " and GMTSCOL1.STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTORDRX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdGMTSCOLX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdGMTSCOL2, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "SKU Inquiry", "Add Styles from Style Master", "Add Styles from Open POs")
        Load_Popup_Menu(grdGMTSCOL3, "BB", "Add All", "Add Selected")
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
            Case "grdGMTSCOL2"
                e.Tool.ToolbarsManager.Tools("Add Styles from Style Master").SharedProps.Visible = (EntryMode <> "V")
                e.Tool.ToolbarsManager.Tools("Add Styles from Open POs").SharedProps.Visible = (EntryMode <> "V")
            Case "grdGMTSCOL3"
                e.Tool.ToolbarsManager.Tools("Add All").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "GMTSCOL3_GMTSCOL3_STYLES" And (EntryMode <> "V"))
                e.Tool.ToolbarsManager.Tools("Add Selected").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "GMTSCOL3_GMTSCOL3_STYLES" And (EntryMode <> "V"))
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
                If chkSelectFrom.Checked Then
                    Dim sqlSN As String = ""
                    'If chkSN_BA.Checked Then sqlSN &= " or SEASON_SEQ_NO = '00'"
                    'If chkSN_CM1.Checked Then sqlSN &= " or SEASON_SEQ_NO = '" & chkSN_CM1.Tag & "'"
                    'If chkSN_C.Checked Then sqlSN &= " or SEASON_SEQ_NO = '" & chkSN_C.Tag & "'"
                    'If chkSN_CP1.Checked Then sqlSN &= " or SEASON_SEQ_NO = '" & chkSN_CP1.Tag & "'"
                    'If chkSN_CP2.Checked Then sqlSN &= " or SEASON_SEQ_NO = '" & chkSN_CP2.Tag & "'"
                    If chkSN_BA.Checked Then sqlSN &= " or SEASON_CODE = '00'"
                    If chkSN_CM1.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CM1.Text & "'"
                    If chkSN_C.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_C.Text & "'"
                    If chkSN_CP1.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CP1.Text & "'"
                    If chkSN_CP2.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CP2.Text & "'"
                    If sqlSN <> "" Then
                        sql_where = " (" & Mid(sqlSN, 5) & ")"
                    End If
                End If
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LONG_SKU", , sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = True
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Loading Styles")

                        grdGMTSCOL2.Visible = False
                        Dim EMsg As String = ""
                        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows

                            Dim DGC_CODE As String = row.Item("DGC_CODE")
                            Dim VEND_CODE As String = row.Item("VEND_CODE")
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                            Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").Rows.Find(New String() {STYLE_COLLECTION_CODE, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
                            If rowGMTSCOL2 IsNot Nothing Then
                                EMsg &= vbCrLf & "Style " & DGC_CODE & "-" & VEND_CODE & "-" & STYLE_CODE & "-" & COLOR_CODE
                            Else
                                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
                            End If
                        Next

                        If EMsg <> "" Then
                            MsgBox(Mid(EMsg, 3), MsgBoxStyle.OkOnly, "Could Not Add Styles Shown - already in Style Collection")
                        End If

                        grdGMTSCOL2.Visible = True
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                        Setup_grdGMTCGMAX()
                    End If
                End If

            Case "Add Styles from Open POs"

                Dim sql_where As String = "POTORDR2.QTY_OPN > 0"

                If chkSelectFrom.Checked Then
                    Dim sqlSN As String = ""
                    If chkSN_BA.Checked Then sqlSN &= " or SEASON_CODE = '00'"
                    If chkSN_CM1.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CM1.Text & "'"
                    If chkSN_C.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_C.Text & "'"
                    If chkSN_CP1.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CP1.Text & "'"
                    If chkSN_CP2.Checked Then sqlSN &= " or SEASON_CODE = '" & chkSN_CP2.Text & "'"
                    If sqlSN <> "" Then
                        sql_where &= " AND (" & Mid(sqlSN, 5) & ")"
                    End If
                End If

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LONG_SKU_PO", , sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = True
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Loading Styles")

                        grdGMTSCOL2.Visible = False
                        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows

                            Dim DGC_CODE As String = row.Item("DGC_CODE")
                            Dim VEND_CODE As String = row.Item("VEND_CODE")
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                            Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").Rows.Find(New String() {STYLE_COLLECTION_CODE, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
                            If rowGMTSCOL2 IsNot Nothing Then
                                MsgBox("Style " & DGC_CODE & "-" & VEND_CODE & "-" & STYLE_CODE & "-" & COLOR_CODE & " is already in Style Collection", MsgBoxStyle.OkOnly, "Cannot Add Style")
                            Else
                                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
                            End If
                        Next
                        grdGMTSCOL2.Visible = True
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                        Setup_grdGMTCGMAX()
                    End If
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Add All", "Add Selected"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Adding Styles")

                For Each grow As UltraWinGrid.UltraGridRow In grd.ActiveRow.ParentRow.ChildBands(0).Rows
                    If e.Tool.Key = "Add All" Or grow.Selected Then
                        Add_Style(grow.Cells("DGC_CODE").Value, _
                                  grow.Cells("VEND_CODE").Value, _
                                  grow.Cells("STYLE_CODE").Value, _
                                  grow.Cells("COLOR_CODE").Value)
                    End If
                Next

                grd.ActiveRow.ParentRow.Cells("STYLE_COUNTER").Value = Get_New_Styles(Val(grd.ActiveRow.ParentRow.Cells("CRITERIA_SEQ").Value & ""))

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

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

    Public Overrides Sub num_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PO_WEEKS"
                If e.KeyCode = Keys.Enter Then
                    Dim PO_WEEKS As Integer = Absx1.numFor("PO_WEEKS").Value
                    Load_GMTCGMAX_Extended("")
                End If
        End Select
    End Sub
#End Region

    Sub Load_GMTSCOLX(initialize As Boolean)
        ASCMAIN1.sql = "Select GMTSCOL1.STYLE_COLLECTION_CODE, GMTSCOL1.STYLE_COLLECTION_DESC, Count (*) LSKUS" & vbCrLf _
            & ",GMTSCOL1.INIT_DATE,GMTSCOL1.INIT_OPER,GMTSCOL1.LAST_DATE,GMTSCOL1.LAST_OPER,GMTSCOL1.VEND_BUYER_CODE" & vbCrLf _
            & " from GMTSCOL1,GMTSCOL2 where GMTSCOL1.STYLE_COLLECTION_CODE = GMTSCOL2.STYLE_COLLECTION_CODE" & vbCrLf _
            & " group by GMTSCOL1.STYLE_COLLECTION_CODE, GMTSCOL1.STYLE_COLLECTION_DESC" & vbCrLf _
            & ",GMTSCOL1.INIT_DATE,GMTSCOL1.INIT_OPER,GMTSCOL1.LAST_DATE,GMTSCOL1.LAST_OPER,GMTSCOL1.VEND_BUYER_CODE"
        If initialize Then
            GMTSCOLX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GMTSCOLX & " Add Primary Key (STYLE_COLLECTION_CODE)")

            ASCMAIN1.sql = "Select GMTSCOL2.* from GMTSCOL2 where ROWNUM < 1"
            GMTSCOL2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GMTSCOL2 & " Add Primary Key (STYLE_COLLECTION_CODE,DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE)")

            ASCMAIN1.sql = "Select GMTCGMA0.*, 'XX' SEASON_SEQ_NO" _
                & ",GMTCGMA0.YYYYWW_FRST_SHPD FR,GMTCGMA0.YYYYWW_LAST_SHPD LR,GMTCGMA0.YYYYWW_FRST_SOLD FS,GMTCGMA0.YYYYWW_LAST_SOLD LS" _
                & ",GMTCGMA0.BOS_UNITS BEGS, GMTCGMA1.SLS_UNITS SLSS, GMTCGMA1.REC_UNITS RECS, GMTCGMA1.SHP_UNITS SHPS, GMTCGMA1.WSX_UNITS WSXS, GMTCGMA1.SSX_UNITS SSXS, GMTCGMA1.ADJ_UNITS ADJS" _
                & ",GMTCGMA0.BOM_UNITS BEGM, GMTCGMA2.SLS_UNITS SLSM, GMTCGMA2.REC_UNITS RECM, GMTCGMA2.SHP_UNITS SHPM, GMTCGMA2.WSX_UNITS WSXM, GMTCGMA2.SSX_UNITS SSXM, GMTCGMA2.ADJ_UNITS ADJM" _
                & ",GMTCGMA0.BOW_UNITS BEGW, GMTCGMA3.SLS_UNITS SLSW, GMTCGMA3.REC_UNITS RECW, GMTCGMA3.SHP_UNITS SHPW, GMTCGMA3.WSX_UNITS WSXW, GMTCGMA3.SSX_UNITS SSXW, GMTCGMA3.ADJ_UNITS ADJW" _
                & " from GMTCGMA0,GMTCGMA1,GMTCGMA2,GMTCGMA3 where ROWNUM < 1"
            GMTCGMAX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GMTCGMAX & " Add Primary Key (DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE,STORE_NO)")

            ASCMAIN1.sql = "Select POTORDR1.VEND_CODE,POTORDR1.VEND_NAME,POTORDR1.PO_CANCEL_DATE,POTORDR1.SEASON_SEQ_NO,POTORDR1.OP_DIV_CODE, POTORDR2.* from POTORDR1,POTORDR2 where ROWNUM < 1"
            POTORDRX = ASCMAIN1.Temp_Table
            '  ASCDATA1.ExecuteSQL("Alter Table " & POTORDRX & " Add Primary Key (STYLE_COLLECTION_CODE)")

        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & GMTSCOLX)
            ASCDATA1.ExecuteSQL("Insert into " & GMTSCOLX & " " & ASCMAIN1.sql)

            EnforceConstraints(False)
            Fill_Records("GMTSCOLX")

            ASCMAIN1.sql = Replace(sqlGMTSCOL2, GMTSCOL2 & " GMTSCOL2", "GMTSCOL2")
            Fill_Records("GMTSCOL2", "", , ASCMAIN1.sql)



            EnforceConstraints(True)
            Sort_grdColumns(grdGMTSCOLX, "STYLE_COLLECTION_DESC")
        End If
    End Sub

    Private Sub grdGMTSCOLX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGMTSCOLX.DoubleClickRow
        STYLE_COLLECTION_CODE = e.Row.Cells("STYLE_COLLECTION_CODE").Value & ""
        If STYLE_COLLECTION_CODE <> "" Then
            Click_Command("View")
        End If
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

            Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").Rows.Find(New String() {STYLE_COLLECTION_CODE, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
            If rowGMTSCOL2 IsNot Nothing Then
                MsgBox("SKU " & SKU_NUMBER & " is already in Style Collection", MsgBoxStyle.OkOnly, "Cannot Add SKU")
            Else
                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
            End If
            Absx1.txtFor("SKU_NUMBER").Text = ""
        End If
        Application.DoEvents()
        Absx1.txtFor("SKU_NUMBER").Focus()
    End Sub

    Sub Add_Style(DGC_CODE As String, VEND_CODE As String, STYLE_CODE As String, COLOR_CODE As String)
        Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").NewRow

        Dim rowGMTSTYL1 As DataRow = LookUp("GMTSTYL1", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
        If rowGMTSTYL1 IsNot Nothing Then
            rowGMTSCOL2 = dst.Tables("GMTSCOL2").NewRow
            For Each dcol As DataColumn In dst.Tables("GMTSCOL2").Columns
                If dcol.ColumnName = "STYLE_COLLECTION_CODE" Then
                    rowGMTSCOL2.Item("STYLE_COLLECTION_CODE") = STYLE_COLLECTION_CODE
                Else
                    If New String() {"OH_STR", "OH_WHS", "QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(dcol.ColumnName) Then
                    Else
                        rowGMTSCOL2.Item(dcol.ColumnName) = rowGMTSTYL1.Item(dcol.ColumnName)
                    End If
                End If
            Next
        Else
            ASCMAIN1.sql = "Select POTORDR2.DGC_CODE, POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & ", POTORDR2.RETAIL_PRICE CUR_RETAIL, POTORDR2.COST LST_UNIT_COST, GMTSEAS1.SEASON_CODE" & vbCrLf _
                & ", POTORDR2.STYLE_DESC DESCRIPTION, POTORDR2.SCALE_CODE" & vbCrLf _
                & " from POTORDR1,POTORDR2,GMTSEAS1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and GMTSEAS1.SEASON_SEQ_NO = POTORDR1.SEASON_SEQ_NO" & vbCrLf _
                & "   and POTORDR2.DGC_CODE = :PARM1 and POTORDR2.STYLE_CODE = :PARM2" & vbCrLf _
                & "   and POTORDR2.COLOR_CODE = :PARM3 and POTORDR1.VEND_CODE = :PARM4"

            Dim row() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVVV", New Object() {DGC_CODE, STYLE_CODE, COLOR_CODE, VEND_CODE}).Select("")

            If row.Length > 0 Then
                rowGMTSCOL2 = dst.Tables("GMTSCOL2").NewRow
                For Each dcol As DataColumn In dst.Tables("GMTSCOL2").Columns
                    If dcol.ColumnName = "STYLE_COLLECTION_CODE" Then
                        rowGMTSCOL2.Item("STYLE_COLLECTION_CODE") = STYLE_COLLECTION_CODE
                    Else
                        If row(0).Table.Columns.Contains(dcol.ColumnName) Then
                            rowGMTSCOL2.Item(dcol.ColumnName) = row(0).Item(dcol.ColumnName)
                        End If
                    End If
                Next
            End If
        End If

        If rowGMTSCOL2 IsNot Nothing AndAlso rowGMTSCOL2.Item("STYLE_COLLECTION_CODE") & "" <> "" Then
            dst.Tables("GMTSCOL2").Rows.Add(rowGMTSCOL2)
            Update_Record_TDA("GMTSCOL2")
            Load_GMTCGMAX_for_LSKU(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
            Dim sqlw As String = " and DGC_CODE = '" & DGC_CODE & "' and VEND_CODE = '" & VEND_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            Load_GMTCGMAX_Extended(sqlw)
            dst.Tables("GMTSCOL2").AcceptChanges()
            optIC.Tag = ""
        End If

    End Sub

    Private Sub optIC_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optIC.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdGMTCGMAX()
        Setup_Panes()
    End Sub

    Private Sub optWMS_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optWMS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_xTD()
        optIC.Tag = ""
        Setup_grdGMTCGMAX()
        'grdGMTCGMAX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdGMTCGMAX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
    End Sub

    Sub Setup_xTD()
        With dst.Tables("GMTCGMAX")
            For Each COL As String In New String() {"BEG", "SLS", "REC", "SHP", "WSX", "SSX", "ADJ"}
                .Columns(COL).Expression = COL & optWMS.Value
                With grdGMTCGMAX.DisplayLayout.Bands(0).Columns(COL)
                    .Header.Caption = Mid(COL, 1, 1) & Mid(COL, 2, 2).ToLower & " " & optWMS.Value & "td"
                    .Width = 65
                End With
                dst.Tables("GMTCGMA0").Columns(COL).Expression = COL & optWMS.Value
            Next
        End With
    End Sub

    Sub Render_Store_Activity()

    End Sub

    Sub Load_GMTCGMAX()

        ' Add rows to Oracle Temp Table GMTCGMAX for each Style in GMTSCOL2
        For Each rowGMTSCOL2 As DataRow In dst.Tables("GMTSCOL2").Select("")
            Dim DGC_CODE As String = rowGMTSCOL2.Item("DGC_CODE")
            Dim VEND_CODE As String = rowGMTSCOL2.Item("VEND_CODE")
            Dim STYLE_CODE As String = rowGMTSCOL2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowGMTSCOL2.Item("COLOR_CODE")
            Load_GMTCGMAX_for_LSKU(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
        Next

        Load_GMTCGMAX_Extended("")
    End Sub

    Sub Load_GMTCGMAX_Extended(sqlw As String)

        If sqlw = "" Then
            For Each rowGMTSCOL2 As DataRow In dst.Tables("GMTSCOL2").Select("")
                rowGMTSCOL2.Item("OH_STR") = 0
                rowGMTSCOL2.Item("OH_WHS") = 0
                rowGMTSCOL2.Item("QTY_OPN_1") = 0
                rowGMTSCOL2.Item("QTY_OPN_2") = 0
                rowGMTSCOL2.Item("QTY_ORD") = 0
                rowGMTSCOL2.Item("QTY_REC") = 0
            Next
        End If

        ASCMAIN1.sql = "Select GMTCGMAX.DGC_CODE, GMTCGMAX.VEND_CODE, GMTCGMAX.STYLE_CODE, GMTCGMAX.COLOR_CODE" & vbCrLf _
            & ", SUM (DECODE(GMTSTOR1.STORE_TYPE_CODE,'S',GMTCGMAX.ON_HAND_UNITS,0)) OH_STR" & vbCrLf _
            & ", SUM (DECODE(GMTSTOR1.STORE_TYPE_CODE,'W',GMTCGMAX.ON_HAND_UNITS,0)) OH_WHS" & vbCrLf _
            & " from GMTSTOR1," & GMTCGMAX & " GMTCGMAX" & vbCrLf _
            & " where GMTSTOR1.STORE_NO = GMTCGMAX.STORE_NO" & vbCrLf _
            & Replace(sqlw, "X.", "GMTCGMAX.") & vbCrLf _
            & " group by GMTCGMAX.DGC_CODE, GMTCGMAX.VEND_CODE, GMTCGMAX.STYLE_CODE, GMTCGMAX.COLOR_CODE"

        For Each rowOH As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").Rows.Find(New String() {STYLE_COLLECTION_CODE, _
                                                                                  rowOH.Item("DGC_CODE"), _
                                                                                  rowOH.Item("VEND_CODE"), _
                                                                                  rowOH.Item("STYLE_CODE"), _
                                                                                  rowOH.Item("COLOR_CODE")})
            rowGMTSCOL2.Item("OH_STR") = Val(rowGMTSCOL2.Item("OH_STR") & "") + Val(rowOH.Item("OH_STR") & "")
            rowGMTSCOL2.Item("OH_WHS") = Val(rowGMTSCOL2.Item("OH_WHS") & "") + Val(rowOH.Item("OH_WHS") & "")
        Next

        ASCMAIN1.sql = "Select * from " & POTORDRX & " POTORDRX" & ASCMAIN1.SQL_Add_WHERE(Replace(sqlw, "X.", "POTORDRX."))

        Dim PO_WEEKS As Integer = Val(Absx1.numFor("PO_WEEKS").Value & "")

        With grdGMTSCOL2.DisplayLayout.Bands(0)
            .Columns("QTY_OPN_1").Header.Caption = "PO<=" & CStr(PO_WEEKS)
            .Columns("QTY_OPN_2").Header.Caption = "PO>" & CStr(PO_WEEKS)
        End With
        With grdGMTSCOLX.DisplayLayout.Bands(1)
            .Columns("QTY_OPN_1").Header.Caption = "PO<=" & CStr(PO_WEEKS)
            .Columns("QTY_OPN_2").Header.Caption = "PO>" & CStr(PO_WEEKS)
        End With
        With grdGMTCGMAX.DisplayLayout.Bands(0)
            .Columns("QTY_OPN_1").Header.Caption = "PO<=" & CStr(PO_WEEKS)
            .Columns("QTY_OPN_2").Header.Caption = "PO>" & CStr(PO_WEEKS)
        End With

        For Each rowPOTORDRX As DataRow In ASCDATA1.GetDataTable.Select("") ' dst.Tables("POTORDRX").Select("")
            Dim rowGMTSCOL2 As DataRow = dst.Tables("GMTSCOL2").Rows.Find(New String() {STYLE_COLLECTION_CODE, _
                                                                                rowPOTORDRX.Item("DGC_CODE"), _
                                                                                rowPOTORDRX.Item("VEND_CODE"), _
                                                                                rowPOTORDRX.Item("STYLE_CODE"), _
                                                                                rowPOTORDRX.Item("COLOR_CODE")})
            Dim PO_CANCEL_DATE As Date = rowPOTORDRX.Item("PO_CANCEL_DATE")
            Dim QTY_OPN As Int64 = Val(rowPOTORDRX.Item("QTY_OPN") & "")
            If QTY_OPN <> 0 Then
                If Format(PO_CANCEL_DATE, "yyyyMMdd") <= Format(Now.Date.AddDays(PO_WEEKS * 7), "yyyyMMdd") Then
                    rowGMTSCOL2.Item("QTY_OPN_1") = Val(rowGMTSCOL2.Item("QTY_OPN_1") & "") + QTY_OPN
                Else
                    rowGMTSCOL2.Item("QTY_OPN_2") = Val(rowGMTSCOL2.Item("QTY_OPN_2") & "") + QTY_OPN
                End If
            End If

            rowGMTSCOL2.Item("QTY_ORD") = Val(rowGMTSCOL2.Item("QTY_ORD") & "") + Val(rowPOTORDRX.Item("QTY_ORD") & "")
            rowGMTSCOL2.Item("QTY_REC") = Val(rowGMTSCOL2.Item("QTY_REC") & "") + Val(rowPOTORDRX.Item("QTY_REC") & "")
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

        Load_POTORDRX_for_LSKU(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
    End Sub

    Sub Load_POTORDRX_for_LSKU(DGC_CODE As String, VEND_CODE As String, STYLE_CODE As String, COLOR_CODE As String)

        ASCMAIN1.sql = "Delete from " & POTORDRX & " where DGC_CODE = :PARM1 and VEND_CODE = :PARM2 and STYLE_CODE = :PARM3 and COLOR_CODE = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})

        ASCMAIN1.sql = "Select POTORDR1.VEND_CODE,POTORDR1.VEND_NAME,POTORDR1.PO_CANCEL_DATE,POTORDR1.SEASON_SEQ_NO,POTORDR1.OP_DIV_CODE" & vbCrLf _
            & ", POTORDR2.* from POTORDR1,POTORDR2" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.DGC_CODE = :PARM1 and POTORDR1.VEND_CODE = :PARM2 and POTORDR2.STYLE_CODE = :PARM3 and POTORDR2.COLOR_CODE = :PARM4" & vbCrLf

        ASCMAIN1.sql = "Insert into " & POTORDRX & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})

    End Sub

    Sub Setup_grdGMTCGMAX()

        If grdGMTSCOL2.ActiveRow Is Nothing OrElse (Not grdGMTSCOL2.ActiveRow.IsDataRow And Not grdGMTSCOL2.ActiveRow.IsGroupByRow) Then
            tabDetails.Visible = False

            UltraExplorerBar1.Groups("Style Image").Visible = False
        Else
            tabDetails.Visible = True

            UltraExplorerBar1.Groups("Style Image").Visible = True
            picStyleImage.Image = ""

            Dim FOLDER_NAME As String = "\\192.168.100.99\skupics\Pictures\" ' ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
            Dim imgba() As Byte = Nothing
            Dim SKU_NO As String = grdGMTSCOL2.ActiveRow.Cells("SKU_NO").Value & ""
            If SKU_NO = "" Then
                ASCMAIN1.sql = "Select SKU_NUMBER from GMTSKUF1 " _
                    & " where DGC_CODE = '" & grdGMTSCOL2.ActiveRow.Cells("DGC_CODE").Value & "'" _
                    & "   and VEND_CODE = '" & grdGMTSCOL2.ActiveRow.Cells("VEND_CODE").Value & "'" _
                    & "   and STYLE_CODE = '" & grdGMTSCOL2.ActiveRow.Cells("STYLE_CODE").Value & "'" _
                    & "   and COLOR_CODE = '" & grdGMTSCOL2.ActiveRow.Cells("COLOR_CODE").Value & "'"
                SKU_NO = ASCDATA1.GetDataValue
            End If
            picStyleImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, SKU_NO & ".JPG", False, , , imgba)
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image " & SKU_NO


            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Sales & On Hand Data")

            Dim DGC_CODE As String = ""
            Dim VEND_CODE As String = ""
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""

            Dim sqlw As String = ""

            If grdGMTSCOL2.ActiveRow.IsGroupByRow Then
                optIC.Tag = ""
                Dim gbyrow As UltraWinGrid.UltraGridGroupByRow = grdGMTSCOL2.ActiveRow
                Do
                    If gbyrow.Column.Key = "DGC_CODE" Then
                        DGC_CODE = gbyrow.Value
                        sqlw &= " and DGC_CODE = '" & DGC_CODE & "'"
                    ElseIf gbyrow.Column.Key = "VEND_CODE" Then
                        VEND_CODE = gbyrow.Value
                        sqlw &= " and VEND_CODE = '" & VEND_CODE & "'"
                    ElseIf gbyrow.Column.Key = "STYLE_CODE" Then
                        STYLE_CODE = gbyrow.Value
                        sqlw &= " and STYLE_CODE = '" & STYLE_CODE & "'"
                    ElseIf gbyrow.Column.Key = "COLOR_CODE" Then
                        COLOR_CODE = gbyrow.Value
                        sqlw &= " and COLOR_CODE = '" & COLOR_CODE & "'"
                    End If

                    gbyrow = gbyrow.ParentRow
                Loop While gbyrow IsNot Nothing

                If sqlw <> "" Then sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw)

            Else
                DGC_CODE = grdGMTSCOL2.ActiveRow.Cells("DGC_CODE").Value
                VEND_CODE = grdGMTSCOL2.ActiveRow.Cells("VEND_CODE").Value
                STYLE_CODE = grdGMTSCOL2.ActiveRow.Cells("STYLE_CODE").Value
                COLOR_CODE = grdGMTSCOL2.ActiveRow.Cells("COLOR_CODE").Value
            End If

            If optIC.Value = "A" Or grdGMTSCOL2.ActiveRow.IsGroupByRow Then
                DGC_CODE = "*"
                VEND_CODE = "*"
                STYLE_CODE = "*"
                COLOR_CODE = "*"
                grdGMTCGMAX.Text = "Sales & On Hand - All Styles in Collection " & STYLE_COLLECTION_CODE & ":" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text & ", Combined" & sqlw
                ASCMAIN1.sql = "Select '*' DGC_CODE, '*' VEND_CODE, '*' STYLE_CODE, '*' COLOR_CODE, STORE_NO" _
                    & ", SUM (SLS_UNITS_TWK) SLS_UNITS_TWK" _
                    & ", SUM (SLS_UNITS_LWK) SLS_UNITS_LWK" _
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
                    & IIf(grdGMTSCOL2.ActiveRow.IsGroupByRow And sqlw <> "", sqlw, "") _
                    & " group by STORE_NO"
            Else
                grdGMTCGMAX.Text = "Collection " & STYLE_COLLECTION_CODE & ": Sales & On Hand - DGC " & DGC_CODE & ", Vendor " & VEND_CODE & ", Style " & STYLE_CODE & ", Color " & COLOR_CODE
                ASCMAIN1.sql = "Select * from " & GMTCGMAX _
                    & " where DGC_CODE = '" & DGC_CODE & "'" _
                    & "   and VEND_CODE = '" & VEND_CODE & "'" _
                    & "   and STYLE_CODE = '" & STYLE_CODE & "'" _
                    & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            End If

            If optIC.Value <> "A" Or optIC.Tag & "" <> "A" Then

                dst.Tables("GMTCGMA0").Rows.Clear()

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

                For Each grow As UltraWinGrid.UltraGridRow In grdGMTCGMAX.Rows
                    If WHSEs.Contains(grow.Cells("STORE_NO").Value) Then
                        grow.Fixed = True
                    End If
                Next

                Sort_grdColumns(grdGMTCGMAX, "STORE_NO")

                'ASCMAIN1.sql = "Select Sum (CASE WHEN POTORDR1.PO_CANCEL_DATE <= SYSDATE + PO_WEEKS * 7 THEN QTY_OPN ELSE 0 END) QTY_OPN_1" _
                '    & " , Sum (CASE WHEN POTORDR1.PO_CANCEL_DATE > SYSDATE + 5 * 7 THEN QTY_OPN ELSE 0 END) QTY_OPN_2" _
                '    & " from POTORDR1,POTORDR2," & GMTSCOL2 & " GMTSCOL2" _
                '    & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" _
                '    & " and POTORDR2.DGC_CODE = GMTSCOL2.DGC_CODE and POTORDR2.STYLE_CODE = GMTSCOL2.STYLE_CODE and POTORDR2.COLOR_CODE = GMTSCOL2.COLOR_CODE and POTORDR1.VEND_CODE = GMTSCOL2.VEND_CODE"

                Dim PO_WEEKS As Integer = Val(Absx1.numFor("PO_WEEKS").Value & "")
                ASCMAIN1.sql = "Select Sum (CASE WHEN POTORDRX.PO_CANCEL_DATE <= SYSDATE + " & CStr(PO_WEEKS) & " * 7 THEN QTY_OPN ELSE 0 END) QTY_OPN_1" _
                    & " , Sum (CASE WHEN POTORDRX.PO_CANCEL_DATE > SYSDATE + " & CStr(PO_WEEKS) & " * 7 THEN QTY_OPN ELSE 0 END) QTY_OPN_2" _
                    & " from " & POTORDRX & " POTORDRX"

                If optIC.Value = "A" Then
                    Fill_Records("POTORDRX", "", True, "Select * from " & POTORDRX)
                Else
                    Fill_Records("POTORDRX", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
                    ASCMAIN1.sql &= " where POTORDRX.DGC_CODE = :PARM1 and POTORDRX.VEND_CODE = :PARM2 and POTORDRX.STYLE_CODE = :PARM3 and POTORDRX.COLOR_CODE = :PARM4"
                End If

                Dim ROW() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVVV", New Object() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE}).Select("")
                If ROW.Length > 0 Then
                    Dim rowGMTCGMAX As DataRow = dst.Tables("GMTCGMAX").Rows.Find(New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE, "068"})
                    If rowGMTCGMAX Is Nothing Then
                        rowGMTCGMAX = dst.Tables("GMTCGMAX").NewRow()
                        rowGMTCGMAX.Item("DGC_CODE") = DGC_CODE
                        rowGMTCGMAX.Item("VEND_CODE") = VEND_CODE
                        rowGMTCGMAX.Item("STYLE_CODE") = STYLE_CODE
                        rowGMTCGMAX.Item("COLOR_CODE") = COLOR_CODE
                        rowGMTCGMAX.Item("STORE_NO") = "068"
                        dst.Tables("GMTCGMAX").Rows.Add(rowGMTCGMAX)
                    End If
                    rowGMTCGMAX.Item("QTY_OPN_1") = ROW(0).Item("QTY_OPN_1")
                    rowGMTCGMAX.Item("QTY_OPN_2") = ROW(0).Item("QTY_OPN_2")
                End If


                If optIC.Value = "A" Then
                    ' Stop
                    Fill_Records("GMTCGMA0", "", True, "Select * from " & GMTCGMAX & sqlw)

                    For Each rowGMTCGMA0 As DataRow In dst.Tables("GMTCGMA0").Select("")
                        Dim FS As String = CStr(rowGMTCGMA0.Item("YYYYWW_FRST_SOLD") & "").PadLeft(6, " ")
                        Dim LS As String = CStr(rowGMTCGMA0.Item("YYYYWW_LAST_SOLD") & "").PadLeft(6, " ")
                        Dim FR As String = CStr(rowGMTCGMA0.Item("YYYYWW_FRST_SHPD") & "").PadLeft(6, " ")
                        Dim LR As String = CStr(rowGMTCGMA0.Item("YYYYWW_LAST_SHPD") & "").PadLeft(6, " ")
                        rowGMTCGMA0.Item("FS") = Mid(FS, Len(FS) - 3, 4)
                        rowGMTCGMA0.Item("LS") = Mid(LS, Len(LS) - 3, 4)
                        rowGMTCGMA0.Item("FR") = Mid(FR, Len(FR) - 3, 4)
                        rowGMTCGMA0.Item("LR") = Mid(LR, Len(LR) - 3, 4)
                    Next
                End If

            End If

            If grdGMTSCOL2.ActiveRow.IsGroupByRow Then
                optIC.Tag = ""
            Else
                optIC.Tag = optIC.Value
            End If


            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Private Sub grdGMTSCOL2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdGMTSCOL2.AfterRowActivate
        If Not ScreenMode Then Exit Sub
        Setup_grdGMTCGMAX()
    End Sub

    Private Sub grdGMTSCOL2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGMTSCOL2.AfterRowsDeleted

        For Each x As LSKU In LSKUs
            ASCMAIN1.sql = "Delete from " & GMTSCOL2 & " where DGC_CODE = :PARM1 and VEND_CODE = :PARM2 and STYLE_CODE = :PARM3 and COLOR_CODE = :PARM4"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {x.DGC_CODE, x.VEND_CODE, x.STYLE_CODE, x.COLOR_CODE})
            ASCMAIN1.sql = "Delete from " & POTORDRX & " where DGC_CODE = :PARM1 and VEND_CODE = :PARM2 and STYLE_CODE = :PARM3 and COLOR_CODE = :PARM4"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {x.DGC_CODE, x.VEND_CODE, x.STYLE_CODE, x.COLOR_CODE})
        Next

        dst.Tables("GMTSCOL2").AcceptChanges()
        dst.Tables("POTORDRX").AcceptChanges()

        Setup_grdGMTCGMAX()
    End Sub

    Private Sub grdGMTSCOL2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGMTSCOL2.BeforeRowsDeleted

        LSKUs.Clear()
        For Each GROW As UltraWinGrid.UltraGridRow In e.Rows
            Dim x As New LSKU
            x.DGC_CODE = GROW.Cells("DGC_CODE").Value
            x.VEND_CODE = GROW.Cells("VEND_CODE").Value
            x.STYLE_CODE = GROW.Cells("STYLE_CODE").Value
            x.COLOR_CODE = GROW.Cells("COLOR_CODE").Value
            LSKUs.Add(x)
        Next
        optIC.Tag = ""
    End Sub

    Private Sub grdGMTSCOL2_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGMTSCOL2.InitializeLayout

    End Sub


    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        '.Item("TYRBVD_P" & Format(I, "00")).Expression = Replace("ISNULL(TY_P00,0) - ISNULL(RB_P00,0)", "P00", "P" & Format(I, "00"))
        '.Item("TYRBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(RB_P00,0)=0,0,100*ISNULL(TYRBVD_P00,0)/ISNULL(RB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYWBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(WB_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(WB_P00,0))/ISNULL(WB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYLYVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(LY_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(LY_P00,0))/ISNULL(LY_P00,0))", "P00", "P" & Format(I, "00"))


        Select Case grd.Name
            Case "grdGMTCGMAX"
                Dim KEY As String = summarySettings.Key
                If KEY.StartsWith("SSR") Then
                    Dim SLS As String = "SLS_UNITS_TWK"
                    Dim STK As String = "BOW_UNITS_TWK"
                    TOTALS.Add(SLS, 0)
                    TOTALS.Add(STK, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SLS) <> 0 Then CustomValue = TOTALS(STK) / TOTALS(SLS)
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.StartsWith("SSR") Then
                    Dim STORE_NO As String = grow2.Cells("STORE_NO").Value
                    If WHSEs.Contains(STORE_NO) Then
                    Else
                        Dim SLS As String = "SLS_UNITS_TWK"
                        Dim STK As String = "BOW_UNITS_TWK"
                        TOTALS(SLS) += Val(grow2.Cells(SLS).Value & "")
                        TOTALS(STK) += Val(grow2.Cells(STK).Value & "")
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grdGMTCGMAX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGMTCGMAX.InitializeRow
        Dim STORE_NO As String = e.Row.Cells("STORE_NO").Value
        If WHSEs.Contains(STORE_NO) Then
            e.Row.Cells("STORE_NO").Appearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub

    Sub Setup_Panes()
        If SELECTION_NO = 0 Then Exit Sub

        grdGMTCGMAX.DisplayLayout.Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
        grdGMTCGMAX.DisplayLayout.Bands(1).ColHeadersVisible = False
        If optIC.Value = "A" Then
            grdGMTCGMAX.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            With grdGMTCGMAX.DisplayLayout.Bands(0).Columns("STORE_NO")
                .Width = 200
                .Header.Caption = "Store/LSKU"
            End With
        Else
            grdGMTCGMAX.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            With grdGMTCGMAX.DisplayLayout.Bands(0).Columns("STORE_NO")
                .Width = 50
                .Header.Caption = "Store"
            End With
        End If
        'If chkShowStoreStyle.Checked And optIC.Value = "A" Then
        '    splGMTSCOL2.SplitterDistance = splGMTSCOL2.Height / 3
        '    splStores.Panel2Collapsed = False
        '    splStores.SplitterDistance = splStores.Height / 2
        'Else
        '    splGMTSCOL2.SplitterDistance = splGMTSCOL2.Height / 2
        '    splStores.Panel2Collapsed = True
        'End If
    End Sub
     
    Private Sub grdGMTSCOL2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGMTSCOL2.InitializeRow
        Dim QTY_REC As Int32 = Val(e.Row.Cells("QTY_REC").Value & "")
        Dim QTY_ORD As Int32 = Val(e.Row.Cells("QTY_ORD").Value & "")
        If QTY_REC > 0 And QTY_ORD > 0 AndAlso System.Math.Abs(QTY_REC - QTY_ORD) / QTY_ORD > 0.1 Then
            e.Row.Cells("QTY_REC").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("QTY_REC").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Function Get_New_Styles(CRITERIA_SEQ As Int32) As Int32
        Dim sql As String = "CRITERIA_SEQ = " & CStr(CRITERIA_SEQ)
        Dim rowGMTSCOL3 As DataRow = dst.Tables("GMTSCOL3").Select(sql)(0)
        ASCDATA1.DeleteRows(dst.Tables("GMTSCOL3_STYLES"), sql)

        sql = ""
        For Each COLUMN_NAME As String In New String() _
            {"DGC_CODE", "VEND_CODE", "STYLE_CODE", "COLOR_CODE", "DESCRIPTION", "SEASON_CODE", "OP_GRP_CODE"}
            Dim VALUE As String = Replace(rowGMTSCOL3.Item(COLUMN_NAME) & "", "'", "")
            If VALUE <> "" Then
                If COLUMN_NAME = "DESCRIPTION" Then
                    sql &= " and " & COLUMN_NAME & " LIKE '%" & VALUE & "%'"
                    'ElseIf COLUMN_NAME = "SEASON_CODE" Then
                    '    Dim row() As DataRow = dst.Tables("GMTSEAS1").Select("SEASON_CODE = '" & VALUE & "'")
                    '    If row.Length = 1 Then
                    '        Dim SEASON_SEQ_NO As String = row(0).Item("SEASON_SEQ_NO")
                    '        sql &= " and " & COLUMN_NAME & " = '" & SEASON_SEQ_NO & "'"
                    '    End If
                Else
                    sql &= " and " & COLUMN_NAME & " = '" & VALUE & "'"
                End If
            End If
        Next

        Dim style_counter As Int32 = 0

        If sql <> "" Then
            ASCMAIN1.sql = "Select DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE from GMTSTYL1 " & ASCMAIN1.SQL_Add_WHERE(sql) _
                & " minus " _
                & " Select DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE from " & GMTSCOL2 & " GMTSCOL2"
            ASCMAIN1.sql = "Select * from GMTSTYL1 where (DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE) in (" & ASCMAIN1.sql & ")"


            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim rowGMTSCOL3_STYLES As DataRow = dst.Tables("GMTSCOL3_STYLES").NewRow
                With rowGMTSCOL3_STYLES
                    .Item("CRITERIA_SEQ") = CRITERIA_SEQ
                    .Item("DGC_CODE") = row.Item("DGC_CODE")
                    .Item("VEND_CODE") = row.Item("VEND_CODE")
                    .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                    .Item("DESCRIPTION") = row.Item("DESCRIPTION")

                    ' .Item("OP_GRP_CODE") = row.Item("OP_GRP_CODE")
                    .Item("SEASON_CODE") = row.Item("SEASON_CODE")
                End With


                dst.Tables("GMTSCOL3_STYLES").Rows.Add(rowGMTSCOL3_STYLES)
                style_counter += 1
            Next

        End If

        Return style_counter
    End Function

#Region "grdGMTSCOL3"

    Private Sub grdGMTSCOL3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGMTSCOL3.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "DGC_CODE"
                If Validate_Columns("DGC_CODE") Then
                    ' If e.Cell.Row.Cells("DGC_CODE").Value & "" <> "" Then e.Cell.Row.Cells("DGC_DESC").Value = cdr.Item("DGC_DESC") & ""
                End If
            Case "VEND_CODE"
                If Validate_Columns("VEND_CODE") Then
                    '  If e.Cell.Row.Cells("VEND_CODE").Value & "" <> "" Then e.Cell.Row.Cells("VEND_NAME").Value = cdr.Item("VEND_NAME") & ""
                End If
            Case "COLOR_CODE"
                If Validate_Columns("COLOR_CODE") Then
                    ' If e.Cell.Row.Cells("COLOR_CODE").Value & "" <> "" Then e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC") & ""
                End If
            Case "SEASON_CODE"
                If Validate_Columns("SEASON_CODE") Then
                    ' If e.Cell.Row.Cells("SEASON_CODE").Value & "" <> "" Then e.Cell.Row.Cells("DGC_DESC").Value = cdr.Item("DGC_DESC") & ""
                End If
            Case "OP_GRP_CODE"
                If Validate_Columns("OP_GRP_CODE") Then
                    ' If e.Cell.Row.Cells("OP_GRP_CODE").Value & "" <> "" Then e.Cell.Row.Cells("OP_GRP_DESC").Value = cdr.Item("OP_GRP_DESC") & ""
                End If

        End Select
    End Sub

 
    Private Sub grdGMTSCOL3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGMTSCOL3.AfterRowActivate

    End Sub

    Private Sub grdGMTSCOL3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGMTSCOL3.AfterRowsDeleted

    End Sub

    Private Sub grdGMTSCOL3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGMTSCOL3.AfterRowUpdate
     
    End Sub

    Private Sub grdGMTSCOL3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdGMTSCOL3.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdGMTSCOL3.ActiveCell
            Select Case .Column.Key

                Case "DGC_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("GMTDGCM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If

                Case "VEND_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("APTVEND1", .Value)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If

                Case "COLOR_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("GMTCOLR1", .Value)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If
                Case Else
                    ' e.Cancel = not Validate_Columns(.Column.Key)
            End Select
        End With

    End Sub

    Private Sub grdGMTSCOL3_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGMTSCOL3.BeforeRowsDeleted
       
    End Sub

    Private Sub grdGMTSCOL3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGMTSCOL3.BeforeRowUpdate

        e.Cancel = Not Validate_Columns("DGC_CODE")
        If Not e.Cancel Then e.Cancel = Not Validate_Columns("VEND_CODE")
        If Not e.Cancel Then e.Cancel = Not Validate_Columns("COLOR_CODE")
        ' If Not e.Cancel Then e.Cancel = Not Validate_Columns("SEASON_CODE")
        ' If Not e.Cancel Then e.Cancel = Not Validate_Columns("OP_GRP_CODE")

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("STYLE_COLLECTION_CODE").Value = Absx1.txtFor("STYLE_COLLECTION_CODE").Text
            Dim CRITERIA_SEQ As Int32 = Val(dst.Tables("GMTSCOL3").Compute("MAX(CRITERIA_SEQ)", "") & "") + 1
            e.Row.Cells("CRITERIA_SEQ").Value = CRITERIA_SEQ
        End If
    End Sub

    Private Sub grdGMTSCOL3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGMTSCOL3.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdGMTSCOL3.ActiveCell.Column.Key
            Case "DGC_CODE"
            Case "VEND_CODE"
            Case "COLOR_CODE"
            Case "SEASON_CODE"
            Case "OP_GRP_CODE"
        End Select

        grdClickCellButton(grdGMTSCOL3, sql_where, False)
    End Sub

    Function Validate_Columns(COLUMN_NAME As String) As Boolean
        Dim TABLE_NAME As String = ""
        Dim ok_if_null As Boolean = True
        Select Case COLUMN_NAME
            Case "DGC_CODE"
                TABLE_NAME = "GMTDGCM1"
            Case "VEND_CODE"
                TABLE_NAME = "APTVEND1"
            Case "COLOR_CODE"
                TABLE_NAME = "GMTCOLR1"
            Case "SEASON_CODE"
                TABLE_NAME = "GMTSEAS1"
            Case "OP_GRP_CODE"
                TABLE_NAME = "GMTODIV1"
        End Select

        Dim Cancel As Boolean = False
        With grdGMTSCOL3.ActiveRow
            Dim CODE_VALUE As String = .Cells(COLUMN_NAME).Value & ""
            If (CODE_VALUE = "" And ok_if_null) Then
            Else
                If CODE_VALUE = "" OrElse LookUp(TABLE_NAME, CODE_VALUE) Is Nothing Then
                    grdGMTSCOL3.ActiveCell = .Cells(COLUMN_NAME)
                    Cancel = True
                End If
            End If
        End With
        Return Not Cancel
    End Function

#End Region

    Private Sub cmdFindNewStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdFindNewStyles.Click
        Find_New_Styles()
    End Sub
     
    Sub Find_New_Styles()
        For Each rowGMTSCOL3 As DataRow In dst.Tables("GMTSCOL3").Select("")
            Dim CRITERIA_SEQ As Int32 = Val(rowGMTSCOL3.Item("CRITERIA_SEQ") & "")
            rowGMTSCOL3.Item("STYLE_COUNTER") = Get_New_Styles(CRITERIA_SEQ)
        Next
    End Sub

    Private Sub grdGMTSCOL3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGMTSCOL3.InitializeRow
        If e.Row.Band.Key = "GMTSCOL3" Then
            If Val(e.Row.Cells("STYLE_COUNTER").Value & "") > 0 Then
                e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Appearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub
End Class