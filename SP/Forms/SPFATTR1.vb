Imports Infragistics.Win.UltraWinGrid

Public Class SPFATTR1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim ATTR_CODE As String
    Dim rowSPTATTR1 As DataRow
    Dim sqlSPTATTRX As String
    Dim ABG_CODEs As New Dictionary(Of String, Integer)
    Dim i_to_ABG As New Dictionary(Of Integer, String)

    Dim Copy_From As String
    Dim Paste_To As String

    Dim appearance_opened As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGreen}
    Dim appearance_closed As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.DarkOrange}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SOFDRBRI" Then
            InquiryMode = True
        End If

        With dst
            ASCMAIN1.sql = "Select * from SPTATTR2"
            Create_TDA(.Tables.Add, "SPTATTR2", "*")

            sqlSPTATTRX = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATUS, ARTCUST2.MALL_CODE" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CUST_RANK_W, ARTCUST2.CUST_STORE_CUST_RANK_M" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from ARTCUST1, ARTCUST2" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf
            ASCMAIN1.sql = sqlSPTATTRX & " And ARTCUST2.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SPTATTRX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select * from ICTBRAN1"
            Create_TDA(.Tables.Add, "ICTBRAN1", "**", 0, False)
            .Tables("ICTBRAN1").Columns.Add("SEL")
            .Tables("ICTBRAN1").Columns("SEL").DefaultValue = "0"
            Fill_Records("ICTBRAN1")

            ASCMAIN1.sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1 where CHANNEL_CODE = '1'"
            Create_TDA(.Tables.Add, "SOTTCLS1", "**", 0, False)
            .Tables("SOTTCLS1").Columns.Add("SEL")
            .Tables("SOTTCLS1").Columns("SEL").DefaultValue = "1"
            Fill_Records("SOTTCLS1")

            ASCMAIN1.sql = "Select * from SPTATTR1"
            Create_TDA(.Tables.Add, "SPTATTR1", "**", 0, False)
            .Tables("SPTATTR1").Columns.Add("SEL")
            .Tables("SPTATTR1").Columns("SEL").DefaultValue = "0"
            Fill_Records("SPTATTR1")

            Dim I As Integer = 0
            For Each rowSPTATTR1 As DataRow In dst.Tables("SPTATTR1").Select("", "ATTR_CODE")
                Dim ATTR_CODE As String = rowSPTATTR1.Item("ATTR_CODE")
                For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("", "BRAND_CODE")
                    Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
                    For Each BRAND_GENDER As String In New String() {"M", "W"}
                        I += 1
                        Dim DC As DataColumn = .Tables("SPTATTRX").Columns.Add("C" & Format(I, "000"))
                        DC.Caption = BRAND_CODE & "_" & BRAND_GENDER
                        ' DC.DefaultValue = "0"
                        Dim ABG As String = ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER
                        ABG_CODEs.Add(ABG, I)
                        i_to_ABG.Add(I, ABG)
                    Next
                Next
            Next

        End With

        grdICTBRAN1.DataSource = dst.Tables("ICTBRAN1")
        Sort_grdColumns(grdICTBRAN1, "BRAND_CODE")

        grdSOTTCLS1.DataSource = dst.Tables("SOTTCLS1")
        Sort_grdColumns(grdSOTTCLS1, "TRADE_CLASS_CODE")

        grdSPTATTR1.DataSource = dst.Tables("SPTATTR1")
        Sort_grdColumns(grdSPTATTR1, "ATTR_CODE")

        lblATTR_CODE.Top = lblCUST_CODE.Top
        txtATTR_CODE.Top = txtCUST_CODE.Top
        txtATTR_DESC.Top = txtCUST_NAME.Top

        grdSPTATTRX.DataSource = dst.Tables("SPTATTRX")

        With grdSPTATTRX.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add()
            G.Header.Caption = "Customer / Store Information"
            G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME In New String() {"CUST_CODE", "CUST_STORE_NO", "CUST_STORE_NAME", "CUST_STORE_STATUS", "TRADE_CLASS_CODE", "MALL_CODE", "CUST_STORE_CUST_RANK_W", "CUST_STORE_CUST_RANK_M"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Hidden = False
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                    If COLUMN_NAME = "CUST_STORE_CUST_RANK_W" Or COLUMN_NAME = "CUST_STORE_CUST_RANK_M" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Chartreuse
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                End With
            Next

            With .Columns("CUST_CODE")
                .Header.Caption = "Customer"
                .Width = 90
            End With
            With .Columns("CUST_STORE_NO")
                .Header.Caption = "Store No"
                .Width = 70
            End With
            With .Columns("CUST_STORE_NAME")
                .Header.Caption = "Store Name"
                .Width = 80
            End With
            With .Columns("CUST_STORE_STATUS")
                .Header.Caption = "Sta"
                .Width = 40
            End With
            With .Columns("TRADE_CLASS_CODE")
                .Header.Caption = "TrdCls"
                .Width = 50
            End With
            With .Columns("MALL_CODE")
                .Header.Caption = "Mall"
                .Width = 90
            End With
            With .Columns("CUST_STORE_CUST_RANK_W")
                .Header.Caption = "Rank" & vbCrLf & "W"
                .Width = 50
            End With
            With .Columns("CUST_STORE_CUST_RANK_M")
                .Header.Caption = "Rank" & vbCrLf & "M"
                .Width = 50
            End With
            Dim I As Integer = 0
            Dim B As Integer = 0
            For Each rowSPTATTR1 As DataRow In dst.Tables("SPTATTR1").Select("", "ATTR_CODE")
                B += 1
                Dim ATTR_CODE As String = rowSPTATTR1.Item("ATTR_CODE")
                G = .Groups.Add()
                G.Key = ATTR_CODE
                G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                .Header.ToolTipText = rowSPTATTR1.Item("ATTR_CODE") & vbCrLf & rowSPTATTR1.Item("ATTR_DESC")

                If B Mod 2 = 0 Then
                    G.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    G.Header.Appearance.BackColor = Drawing.Color.White
                Else
                    G.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    G.Header.Appearance.BackColor = Drawing.Color.White
                End If

                For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("", "BRAND_CODE")
                    Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
                    For Each BRAND_GENDER As String In New String() {"M", "W"}
                        I += 1
                        Dim COLUMN_NAME As String = "C" & Format(I, "000")
                        With .Columns(COLUMN_NAME)
                            .Header.Caption = BRAND_CODE & "_" & BRAND_GENDER
                            .Group = G
                            .Hidden = False
                            .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                            .Header.Appearance.TextHAlign = HAlign.Center
                            .Header.Appearance.TextVAlign = VAlign.Bottom
                            .CellAppearance.TextHAlign = HAlign.Center
                            .Header.Appearance.FontData.SizeInPoints = 8
                            .Header.TextOrientation = New TextOrientationInfo(90, TextFlowDirection.Horizontal)
                            .Width = 30
                            .Header.ToolTipText = rowICTBRAN1.Item("BRAND_CODE") & "_" & BRAND_GENDER & vbCrLf & rowICTBRAN1.Item("BRAND_NAME")
                            ' .Style = UltraWinGrid.ColumnStyle.CheckBox
                            Create_Summary(grdSPTATTRX, COLUMN_NAME)
                        End With
                    Next
                Next
            Next
            .Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        End With

        Create_Summary(grdSPTATTRX, "CUST_CODE", "Count")

        With grdSPTATTRX.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"
                If optFilter.Value = "C" Then
                    If Absx1.txtFor("CUST_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Select a Customer Code"
                    Else
                        Validate_Code("CUST_CODE")
                        If EMsg = "" Then
                            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                            ATTR_CODE = ""
                            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                            If rowARTCUST1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                            End If
                        End If
                    End If
                Else
                    If Absx1.txtFor("ATTR_CODE").Text = "" Then
                        'EMsg &= vbCr & "You Must Select a Brand Code"
                        CUST_CODE = ""
                        ATTR_CODE = ""
                    Else
                        Validate_Code("ATTR_CODE")
                        If EMsg = "" Then
                            CUST_CODE = ""
                            ATTR_CODE = Absx1.txtFor("ATTR_CODE").Text
                            rowSPTATTR1 = LookUp("SPTATTR1", ATTR_CODE)
                            If rowSPTATTR1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Attribute Code " & ATTR_CODE
                            End If
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If optFilter.Value = "C" Then
                        If Not ASCMAIN1.Logical_Lock("SPTATTR1", "CUST_CODE:" & CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("SPTATTR1", "CUST_CODE:*") Then Exit Sub
                    Else
                        If Not ASCMAIN1.Logical_Lock("SPTATTR1", "CUST_CODE:*") Then Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update", "Save"

            Case "Load XLS"
                'Ask for spreadsheet – user may cancel
                'Validate that the spreadsheet reflects the same attributes and brands that are on the screen
                'Mention that stores that are in the spreadsheet, but not on the screen, will be ignored
                'Ask how to handle stores that are in the screen that are not in the spreadsheet
                '1)	They can be cleared
                '2)	They can be left alone on the screen
                ' - DON'T ASK - JUST DESCRIBE WHAT WILL HAPPEN IF STORES ARE IN SCREEN AND NOT IN SPREADSHEET

                MsgBox("IMPORTANT:" & vbCrLf _
                       & "1) Spreadsheet Attributes And Brands must exactly match the screen" & vbCrLf _
                       & "2) Stores that are in the spreadsheet, but not on the screen, will be ignored" & vbCrLf _
                       & "3) Stores that are in the screen that are not in the spreadsheet will not be changed", MsgBoxStyle.OkOnly, Me.Text)

                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                    openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME <> "" Then

                    Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME, System.Globalization.CultureInfo.CurrentCulture)
                    Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
                    Dim range As SpreadsheetGear.IRange = Nothing

                    If worksheet.Cells("B1").Value & "" <> "SPTATTRX" Then
                        MsgBox("Workbook selected does Not appear to be correct format", MsgBoxStyle.OkOnly, "Cannot Proceed")
                        Exit Sub
                    End If

                    If worksheet.Cells("A6").Value & "" <> "Customer / Store Information" Then
                        MsgBox("Workbook selected does Not appear to be correct format", MsgBoxStyle.OkOnly, "Cannot Proceed")
                        Exit Sub
                    End If

                    Dim colABG As New Dictionary(Of Integer, String)

                    Dim ATTR_CODE As String = ""
                    Dim BRAND_CODE As String = ""
                    Dim BRAND_GENDER As String = ""
                    Dim BG As String = ""
                    Dim I As Integer = 7
                    Do While worksheet.Cells(6, I + 1).Value & "" <> ""
                        I += 1
                        Dim A As String = worksheet.Cells(5, I).Value & ""
                        If A <> "" Then ATTR_CODE = A
                        BG = worksheet.Cells(6, I).Value & ""
                        BRAND_CODE = Split(BG, "_")(0)
                        BRAND_GENDER = Split(BG, "_")(1)
                        Dim ABG As String = ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER
                        If Not ABG_CODEs.ContainsKey(ABG) Then
                            MsgBox("Cannot Map Attribute - Brand: " & ABG, MsgBoxStyle.OkOnly, "Cannot Proceed")
                            Exit Sub
                        End If
                        colABG.Add(I, ABG)
                    Loop

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Loading XLS")

                    Dim matched As Integer = 0
                    Dim not_matched As Integer = 0

                    Dim r As Integer = 6
                    Do While worksheet.Cells(r + 1, 0).Value & "" <> "" And worksheet.Cells(r + 1, 1).Value & "" <> ""
                        r += 1
                        Dim CUST_CODE As String = worksheet.Cells(r, 0).Value
                        Dim CUST_STORE_NO As String = worksheet.Cells(r, 1).Value
                        Dim rowSPTATTRX As DataRow = dst.Tables("SPTATTRX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                        If rowSPTATTRX IsNot Nothing Then
                            matched += 1
                            For Each C As Integer In colABG.Keys
                                Dim ABG As String = colABG(C)
                                Dim iABG As Integer = ABG_CODEs(ABG)
                                Dim V As String = worksheet.Cells(r, C).Value & ""
                                If V <> "" Then V = "1"
                                rowSPTATTRX.Item("C" & Format(iABG, "000")) = V
                            Next
                        Else
                            not_matched += 1
                        End If
                        If r Mod 100 = 0 Then
                            ASCMAIN1.Progress("-", CUST_CODE & "_" & CUST_STORE_NO)
                        End If
                    Loop

                    MsgBox("Load Completed Successfully" & vbCrLf _
                           & CStr(matched) & " Records Matched" & vbCrLf _
                           & CStr(not_matched) & " Records Not Matched", MsgBoxStyle.OkOnly, "Success")

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
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

            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update", "Save"
                Update_Record()
                If eItemKey = "Update" Then
                    Mode_Settings(False)
                End If

            Case "Cancel", "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Save").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not (EntryMode = "E")
                    .Items("Edit").Visible = (Not InquiryMode) And Not (EntryMode = "E")
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Save").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Load XLS").Visible = ScreenMode And (Not InquiryMode And EntryMode <> "V") ' And (optFilter.Value = "A")

                End With
                .Groups("Attributes").Visible = ScreenMode And (optFilter.Value = "C" Or ATTR_CODE = "")
                .Groups("Brands").Visible = ScreenMode
                .Groups("Trade Classes").Visible = ScreenMode
            End With
        End If

        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                grdSPTATTRX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSPTATTRX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSPTATTRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdSPTATTRX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSPTATTRX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSPTATTRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl.Panel1Collapsed = ScreenMode
        splSPTATTRX.Visible = ScreenMode

        splSPTATTRX.Panel2Collapsed = True

        If ScreenMode Then
            Set_Groups()
            Set_Columns()

            Dim W As Integer = 30
            If EntryMode = "V" Then W = 30 ' 65
            For Each HC_CODE As String In ABG_CODEs.Keys
                Dim C As Integer = ABG_CODEs(HC_CODE)
                grdSPTATTRX.DisplayLayout.Bands(0).Columns("C" & Format(C, "000")).Width = W
            Next
        Else
            Clear_Record()
            grdSPTATTRX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SPTATTRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("ATTR_CODE").Text = ""

        Copy_From = ""
        Paste_To = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If optFilter.Value = "A" Then
            ASCMAIN1.sql = sqlSPTATTRX
            Fill_Records("SPTATTRX", "", , ASCMAIN1.sql)
            grdSPTATTRX.Text = "Store / Attribute Matrix for " _
                & IIf(ATTR_CODE = "", "All Attributes", ATTR_CODE) _
                & ", for All Customers"

            ASCMAIN1.sql = "Select * from SPTATTR2"
            If ATTR_CODE <> "" Then
                ASCMAIN1.sql &= " where ATTR_CODE = '" & ATTR_CODE & "'"
            End If
            Fill_Records("SPTATTR2", "", , ASCMAIN1.sql)
        Else
            Fill_Records("SPTATTRX", New String() {CUST_CODE})
            grdSPTATTRX.Text = "Store / Attribute Matrix for Customer " & CUST_CODE

            ASCMAIN1.sql = "Select * from SPTATTR2 where CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SPTATTR2", "", , ASCMAIN1.sql)
        End If

        Dim ABG_CODEs_In As New List(Of String)

        grdSPTATTRX.SuspendSummaryUpdates()
        grdSPTATTRX.SuspendLayout()

        For Each row As DataRow In dst.Tables("SPTATTR2").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim ATTR_CODE As String = row.Item("ATTR_CODE")
            Dim BRAND_CODE As String = row.Item("BRAND_CODE")
            Dim BRAND_GENDER As String = row.Item("BRAND_GENDER")
            Dim ABG As String = ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER
            Dim i As Integer = ABG_CODEs(ABG)

            Dim rowSPTATTRX As DataRow = dst.Tables("SPTATTRX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            rowSPTATTRX.Item("C" & Format(i, "000")) = "1"
            If Not ABG_CODEs_In.Contains(ABG) Then ABG_CODEs_In.Add(ABG)
        Next

        grdSPTATTRX.ResumeSummaryUpdates(True)
        grdSPTATTRX.ResumeLayout()

        Dim ATTR_CODEs As New List(Of String)
        For Each ABG As String In ABG_CODEs_In
            Dim ATTR_CODE As String = Split(ABG, "_")(0)
            If Not ATTR_CODEs.Contains(ATTR_CODE) Then
                ATTR_CODEs.Add(ATTR_CODE)
            End If
            Dim BRAND_CODE As String = Split(ABG, "_")(1)
            Dim BRAND_GENDER As String = Split(ABG, "_")(2)
            If Me.ATTR_CODE <> "" AndAlso Not ABG_CODEs_In.Contains(ABG) Then
                ABG_CODEs_In.Add(ABG)
            End If
        Next

        Sort_grdColumns(grdSPTATTRX, "CUST_CODE,CUST_STORE_NO")

        For Each rowSPTATTR1 As DataRow In dst.Tables("SPTATTR1").Select("")
            Dim ATTR_CODE As String = rowSPTATTR1.Item("ATTR_CODE")
            Dim ATTR_STATUS As String = rowSPTATTR1.Item("ATTR_STATUS") & ""
            If Me.ATTR_CODE <> "" Then
                rowSPTATTR1.Item("SEL") = IIf(ATTR_CODE = Me.ATTR_CODE, "1", "0")
            Else
                ' NOT RIGHT
                rowSPTATTR1.Item("SEL") = IIf((optFilter.Value = "A" And ATTR_STATUS = "A") Or ATTR_CODEs.Contains(ATTR_CODE), "1", "0")
            End If
        Next

        For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("")
            Dim BRAND_STATUS As String = rowICTBRAN1.Item("BRAND_STATUS")
            If BRAND_STATUS = "A" Then rowICTBRAN1.Item("SEL") = "1"
        Next

        dst.Tables("SPTATTRX").AcceptChanges()

        EnforceConstraints(True)
        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-composing Attribute Matrix")

        dst.Tables("SPTATTR2").Rows.Clear()

        For Each row As DataRow In dst.Tables("SPTATTRX").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")

            For Each ABG As String In ABG_CODEs.Keys

                Dim ATTR_CODE As String = Split(ABG, "_")(0)
                Dim BRAND_CODE As String = Split(ABG, "_")(1)
                Dim BRAND_GENDER As String = Split(ABG, "_")(2)

                Dim i As Integer = ABG_CODEs(ABG)
                Dim C000 As String = "C" & Format(i, "000")
                If row.Item(C000) & "" <> "" Then
                    Dim rowSPTATTR2 As DataRow = dst.Tables("SPTATTR2").NewRow
                    rowSPTATTR2.Item("CUST_CODE") = CUST_CODE
                    rowSPTATTR2.Item("CUST_STORE_NO") = CUST_STORE_NO
                    rowSPTATTR2.Item("ATTR_CODE") = ATTR_CODE
                    rowSPTATTR2.Item("BRAND_CODE") = BRAND_CODE
                    rowSPTATTR2.Item("BRAND_GENDER") = BRAND_GENDER
                    dst.Tables("SPTATTR2").Rows.Add(rowSPTATTR2)
                End If
            Next
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Try
            BeginTrans()

            ASCMAIN1.sql = "Delete from SPTATTR2"
            If optFilter.Value = "C" Then
                ASCMAIN1.sql &= " where CUST_CODE = '" & CUST_CODE & "'"
            Else
                If ATTR_CODE <> "" Then
                    ASCMAIN1.sql &= " where ATTR_CODE = '" & ATTR_CODE & "'"
                End If
            End If
            ASCDATA1.ExecuteSQL()

            Update_Record_TDA("SPTATTR2")

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTATTRX, "SSBBBB", "Show Filter", "Show GroupBox", "Copy From", "Paste To", "Select All Stores for Attribute", "De-Select All Stores for Attribute")
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

                    tlb_btn = tlb_pop.Tools("Copy From")
                    If grd.ActiveCell Is Nothing Or (EntryMode <> "E") Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length <> 4 Or Not C.StartsWith("C") Then '
                            tlb_btn.SharedProps.Visible = False
                        Else
                            Dim ATTR_CODE As String = grd.ActiveCell.Column.Group.Header.Caption
                            Dim BG As String = grd.ActiveCell.Column.Header.Caption
                            Dim BRAND_CODE As String = Split(BG, "_")(0)
                            Dim BRAND_GENDER As String = Split(BG, "_")(1)
                            tlb_btn.SharedProps.Caption = "Set Copy From to " & ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER
                            tlb_btn.Tag = ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER
                            tlb_btn.SharedProps.Visible = True
                        End If

                    End If


                    tlb_btn = tlb_pop.Tools("Paste To")
                    If grd.ActiveCell Is Nothing Or (EntryMode <> "E") Or Copy_From = "" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim ATTR_CODE = grd.ActiveCell.Column.Group.Header.Caption
                        Dim BRAND_CODE = grd.ActiveCell.Column.Header.Caption
                        tlb_btn.SharedProps.Caption = "Paste " & Copy_From & " To " & ATTR_CODE & "_" & BRAND_CODE
                        tlb_btn.Tag = ATTR_CODE & "_" & BRAND_CODE
                        tlb_btn.SharedProps.Visible = True
                    End If

                    tlb_btn = tlb_pop.Tools("Select All Stores for Attribute")
                    If grd.ActiveCell Is Nothing Or EntryMode <> "E" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length = 4 And C.StartsWith("C") And Mid(C, 2, 3) >= "001" And Mid(C, 2, 3) <= "999" Then
                            tlb_btn.SharedProps.Visible = True
                            Dim AB As String = i_to_ABG(Val(Mid(C, 2, 3)))
                            tlb_btn.Tag = AB
                            tlb_btn.SharedProps.Caption = "Select All Stores for " & AB
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If

                    tlb_btn = tlb_pop.Tools("De-Select All Stores for Attribute")
                    If grd.ActiveCell Is Nothing Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length = 4 And C.StartsWith("C") And Mid(C, 2, 3) >= "001" And Mid(C, 2, 3) <= "999" Then
                            tlb_btn.SharedProps.Visible = True
                            Dim AB As String = i_to_ABG(Val(Mid(C, 2, 3)))
                            tlb_btn.Tag = AB
                            tlb_btn.SharedProps.Caption = "De-Select All Stores for " & AB
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If
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
            Case "Select All Stores for Attribute", "De-Select All Stores for Attribute"
                If MsgBox("Are you sure that you want to " & vbCrLf & e.Tool.SharedProps.Caption & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Updating")

                Dim SEL As String = "1"
                If e.Tool.Key.StartsWith("De-Select") Then SEL = ""
                Dim tlb_btn As UltraWinToolbars.ButtonTool = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim AB As String = tlb_btn.Tag
                Dim C As String = "C" & Format(ABG_CODEs(AB), "000")
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells(C).Value = SEL
                    grow.Update()
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Copy From"
                Dim tlb_btn As UltraWinToolbars.ButtonTool = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Copy_From = tlb_btn.Tag

            Case "Paste To"
                Dim tlb_btn As UltraWinToolbars.ButtonTool = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Paste_To = tlb_btn.Tag

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Pasting")

                Dim iFrom As Integer = ABG_CODEs(Copy_From)
                Dim iTo As Integer = ABG_CODEs(Paste_To)
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("C" & Format(iTo, "000")).Value = grow.Cells("C" & Format(iFrom, "000")).Value
                    grow.Update()
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
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

#End Region

    Private Sub grdICTBRAN1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdICTBRAN1.AfterRowUpdate
        Set_Columns()
    End Sub

    Sub Set_Columns()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Screen")
        For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("")
            Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
            For Each rowSPTATTR1 As DataRow In dst.Tables("SPTATTR1").Select("")
                Dim ATTR_CODE As String = rowSPTATTR1.Item("ATTR_CODE")
                For Each BRAND_GENDER As String In New String() {"M", "W"}
                    Dim I As Integer = ABG_CODEs(ATTR_CODE & "_" & BRAND_CODE & "_" & BRAND_GENDER)
                    grdSPTATTRX.DisplayLayout.Bands(0).Columns("C" & Format(I, "000")).Hidden = (rowICTBRAN1.Item("SEL") & "" <> "1")
                Next
            Next
        Next

        Dim all_selected As Boolean = True
        Dim sqlTC As String = ""
        Dim TCs As New List(Of String)
        For Each row As DataRow In dst.Tables("SOTTCLS1").Select("")
            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
            Dim SEL As String = row.Item("SEL")
            If SEL = "1" Then
                TCs.Add(TRADE_CLASS_CODE)
                sqlTC &= " or TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"
            Else
                all_selected = False
            End If
        Next

        Dim dvw As DataView = DirectCast(grdSPTATTRX.DataSource, DataTable).DefaultView

        If Not all_selected Then
            dvw.RowFilter = Mid(sqlTC, 5)
        Else
            dvw.RowFilter = ""
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSPTATTR1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTATTR1.AfterRowUpdate
        Set_Groups()
    End Sub

    Sub Set_Groups()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Screen")
        For Each rowSPTATTR1 As DataRow In dst.Tables("SPTATTR1").Select("")
            Dim ATTR_CODE As String = rowSPTATTR1.Item("ATTR_CODE")
            grdSPTATTRX.DisplayLayout.Bands(0).Groups(ATTR_CODE).Hidden = (rowSPTATTR1.Item("SEL") & "" <> "1")
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSPTATTRX_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSPTATTRX.ClickCell

    End Sub

    Private Sub grdSPTATTRX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTATTRX.InitializeRow
        For Each ABG As String In ABG_CODEs.Keys
            Dim I As Integer = ABG_CODEs(ABG)
            If e.Row.Cells("C" & Format(I, "000")).Value & "" = "" Then
                e.Row.Cells("C" & Format(I, "000")).Appearance = Nothing
            Else
                e.Row.Cells("C" & Format(I, "000")).Appearance = appearance_opened
            End If
        Next
    End Sub

    Private Sub optFilter_ValueChanged(sender As Object, e As EventArgs) Handles optFilter.ValueChanged
        lblCUST_CODE.Visible = (optFilter.Value = "C")
        txtCUST_CODE.Visible = (optFilter.Value = "C")
        txtCUST_NAME.Visible = (optFilter.Value = "C")

        lblATTR_CODE.Visible = (optFilter.Value = "A")
        txtATTR_CODE.Visible = (optFilter.Value = "A")
        txtATTR_DESC.Visible = (optFilter.Value = "A")

        lblBRAND_CODE_ALL.Visible = (optFilter.Value = "A")
    End Sub

    Private Sub grdSPTATTRX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSPTATTRX.BeforeCellUpdate

    End Sub

    Private Sub grdSPTATTRX_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTATTRX.BeforeExitEditMode
        With grdSPTATTRX.ActiveCell
            Dim COLUMN_NAME As String = .Column.Key
            If COLUMN_NAME.StartsWith("C") And COLUMN_NAME.Length = 4 And Mid(COLUMN_NAME, 2, 4) >= "001" And Mid(COLUMN_NAME, 2, 4) < "999" Then
                If .EditorResolved.Value & "" <> "" Then
                    If .EditorResolved.Value & "" <> "1" Then
                        .EditorResolved.Value = "1"
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grdSOTTCLS1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdSOTTCLS1.AfterRowUpdate
        Set_Columns()
    End Sub
End Class