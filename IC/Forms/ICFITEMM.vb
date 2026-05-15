Imports Infragistics.Win.UltraWinGrid

Public Class ICFITEMM

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private lstOperations As New List(Of String)(New String() {"ITEM_CODE", "ITEM_EAN_CODE", "ITEM_UPC_CODE", "ITEM_SO_QTY_MIN", "ITEM_SO_QTY_MULT", "ITEM_ALT_SORT"})
    Private lstCheckBoxes As New List(Of String)(New String() {"ITEM_HIDE_FROM_CAT", "ITEM_CONTAINS_ALCOHOL", "ITEM_LOT_CONTROL", "ITEM_WEIGHT_CHECK", "ITEM_APPR_1ST_REC", "ITEM_CRITICAL_TO_SHIP", "COMPLETED", "APPROVED", "REQUIRES_EAN"})
    Private SO_PARM_UPC_VENDOR_ID As String = String.Empty

    Private lstRequiredManualEntry As New List(Of String)(New String() {"PROD_CODE", "COLLECTION_CODE", "ITEM_TYPE_CODE", "ITEM_CLASS_CODE", "ITEM_DESC", "ITEM_RETAIL_PRICE"})

    Private verifyUpdateRow As Boolean = False
    Private lstIctitemmNoEdit As New List(Of String)
    Private lstTatipic1AllowEdit As New List(Of String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ICTITEMM", "*")
            With .Tables("ICTITEMM")
                .Columns.Add("PROD_DESC", GetType(String), "PROD_CODE")
                .Columns.Add("COLLECTION_NAME", GetType(String), "COLLECTION_CODE")
                .Columns.Add("ITEM_TYPE_DESC", GetType(String), "ITEM_TYPE_CODE")
                .Columns.Add("ITEM_CLASS_DESC", GetType(String), "ITEM_CLASS_CODE")
                .Columns.Add("HC_NAME", GetType(String), "HC_CODE")
                .Columns.Add("SEASON_DESC", GetType(String), "SEASON_CODE")
                .Columns.Add("VEND_NAME", GetType(String), "VEND_CODE")
                .Columns.Add("COST_CATGY_DESC", GetType(String), "COST_CATGY_CODE")
                .Columns.Add("HMAT_DESC", GetType(String), "HMAT_CODE")
            End With

            ' LAst 90 Days
            ASCMAIN1.sql = "Select * FROM ICTITEMM 
                                WHERE NVL(COMPLETED, '0') = '1' 
                                AND NVL(APPROVED, '0') = '1' 
                                AND STATUS = 'U' 
                                AND LAST_DATE >= SYSDATE - 91
                            UNION
                                SELECT 
                                'I' || ROWNUM ITEM_CODE_CTL, 
                                TATIPIC1.ITEM_CODE, 
                                TATIPIC1.ITEM_DESC, 
                                TATIPIC1.ITEM_EAN_CODE, 
                                NULL ITEM_UPC_CODE, 
                                TATIPIC1.ITEM_RETAIL_PRICE, 
                                TATIPIC1.LAUNCH_DATE, 
                                TATIPIC1.PROD_CODE, 
                                TATIPIC1.COLLECTION_CODE, 
                                ICTITEM1.ITEM_TYPE_CODE, 
                                ICTITEM1.ITEM_HIDE_FROM_CAT, 
                                TATIPIC1.ITEM_CLASS_CODE, 
                                ICTCOLL1.HC_CODE, 
                                TATIPIC1.COST_CATGY_CODE, 
                                TATIPIC1.SEASON_CODE, 
                                TATIPIC1.ITEM_BASIC_PROMO, 
                                TATIPIC1.COUNTRY_CODE, 
                                TATIPIC1.ITEM_CONTAINS_ALCOHOL, 
                                TATIPIC1.ITEM_LOT_CONTROL, 
                                TATIPIC1.ITEM_WEIGHT_CHECK, 
                                TATIPIC1.ITEM_APPR_1ST_REC, 
                                TATIPIC1.ITEM_CRITICAL_TO_SHIP, 
                                TATIPIC1.ITEM_SHELF_LIFE_YRS, 
                                ICTITEM1.VEND_CODE, 
                                TATIPIC1.ITEM_SO_QTY_MIN, 
                                TATIPIC1.ITEM_SO_QTY_MULT, 
                                TATIPIC1.COMPLETED, 
                                TATIPIC1.COMPLETED_BY, 
                                TATIPIC1.INIT_OPER, 
                                TATIPIC1.LAST_OPER, 
                                TATIPIC1.INIT_DATE, 
                                TATIPIC1.LAST_DATE, 
                                ICTITEM1.HMAT_CODE, 
                                TATIPIC1.ITEM_ALT_SORT, 
                                TATIPIC1.IPSA_STATUS_CODE STATUS, 
                                TATIPIC1.APPROVED, 
                                TATIPIC1.APPROVED_BY, 
                                TATIPIC1.COMPLETED_ON, 
                                TATIPIC1.APPROVED_ON,
                                '1' REQUIRES_EAN,
                                TATIPIC1.MARKETING_NOTES
                                FROM TATIPIC1, ICTITEM1, ICTCOLL1
                                WHERE TATIPIC1.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                AND TATIPIC1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE (+)
                                AND TATIPIC1.IPSA_STATUS_CODE = 'U' 
                                AND TATIPIC1.LAST_DATE >= SYSDATE - 91
                                AND NVL(TATIPIC1.COMPLETED, '0') = '1' 
                                AND NVL(TATIPIC1.APPROVED, '0') = '1'"
            Create_TDA(.Tables.Add, "ICTITEMH", ASCMAIN1.sql, 0, False)
            With .Tables("ICTITEMH")
                .Columns.Add("PROD_DESC", GetType(String), "PROD_CODE")
                .Columns.Add("COLLECTION_NAME", GetType(String), "COLLECTION_CODE")
                .Columns.Add("ITEM_TYPE_DESC", GetType(String), "ITEM_TYPE_CODE")
                .Columns.Add("ITEM_CLASS_DESC", GetType(String), "ITEM_CLASS_CODE")
                .Columns.Add("HC_NAME", GetType(String), "HC_CODE")
                .Columns.Add("SEASON_DESC", GetType(String), "SEASON_CODE")
                .Columns.Add("VEND_NAME", GetType(String), "VEND_CODE")
                .Columns.Add("COST_CATGY_DESC", GetType(String), "COST_CATGY_CODE")
                .Columns.Add("HMAT_DESC", GetType(String), "HMAT_CODE")
            End With

            For Each tableName As String In {"ICTCOLL1", "ICTPROD1", "ICTTYPE1", "ICTCLAS1", "ICTCOLL0", "ICTSEAS1", "TATCNTRY", "APTVEND1", "ICTBRAN1", "ICTHMAT1"}
                Create_TDA(.Tables.Add, tableName, "*")
                Fill_Records(tableName, String.Empty, True, $"SELECT * FROM {tableName}")
            Next

            Create_TDA(.Tables.Add, "ICTITEM1", "*")
            Create_TDA(.Tables.Add("ICTITEMM_UPDATE"), "ICTITEMM", "*")
            Create_TDA(.Tables.Add, "TATIPIC1", "*")
        End With

        grdICTITEMM.DataSource = dst.Tables("ICTITEMM")
        grdICTITEMH.DataSource = dst.Tables("ICTITEMH")
        grdTATIPIC1.DataSource = dst.Tables("TATIPIC1")

        For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In {grdICTITEMM, grdICTITEMH}
            ASCMAIN1.Add_Value_List(grd, "PROD_DESC", "SELECT PROD_CODE, PROD_DESC FROM ICTPROD1")
            ASCMAIN1.Add_Value_List(grd, "COLLECTION_NAME", "SELECT COLLECTION_CODE, COLLECTION_NAME FROM ICTCOLL1")
            ASCMAIN1.Add_Value_List(grd, "ITEM_TYPE_DESC", "SELECT ITEM_TYPE_CODE, ITEM_TYPE_DESC FROM ICTTYPE1")
            ASCMAIN1.Add_Value_List(grd, "ITEM_CLASS_DESC", "SELECT ITEM_CLASS_CODE, ITEM_CLASS_DESC FROM ICTCLAS1")
            ASCMAIN1.Add_Value_List(grd, "HC_NAME", "SELECT HC_CODE, HC_NAME FROM ICTCOLL0")
            ASCMAIN1.Add_Value_List(grd, "SEASON_DESC", "SELECT SEASON_CODE, SEASON_DESC FROM ICTSEAS1")
            ASCMAIN1.Add_Value_List(grd, "VEND_NAME", "SELECT VEND_CODE, VEND_NAME FROM APTVEND1")
            ASCMAIN1.Add_Value_List(grd, "HMAT_DESC", "SELECT HMAT_CODE, HMAT_DESC FROM ICTHMAT1")
            ASCMAIN1.Add_Value_List(grd, "ITEM_BASIC_PROMO")
            ASCMAIN1.Add_Value_List(grd, "COST_CATGY_DESC", "SELECT COST_CATGY_CODE, COST_CATGY_DESC FROM ICTCOST1")
            ASCMAIN1.Add_Value_List(grd, "INIT_OPER", "SELECT USER_ID, USER_NAME FROM ASTUSER1")
            ASCMAIN1.Add_Value_List(grd, "COMPLETED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")
            ASCMAIN1.Add_Value_List(grd, "APPROVED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")

            Create_Summary(grd, "ITEM_CODE", "Count")
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grd.Visible = True

            If grd.Name = grdICTITEMH.Name Then
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If

            For Each grdCol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                grdCol.CellAppearance.BackColor = Nothing
                grdCol.Header.Appearance.TextHAlign = HAlign.Left
                grdCol.CellAppearance.TextHAlign = HAlign.Left
                If grdCol.CellActivation = UltraWinGrid.Activation.NoEdit Then
                    If grd.Name = grdICTITEMM.Name Then
                        lstIctitemmNoEdit.Add(grdCol.Header.Column.Key)
                    End If
                    Select Case grdCol.Header.Column.Key
                        Case "COMPLETED_ON", "COMPLETED_BY"
                            'grdCol.CellAppearance.BackColor = Drawing.Color.LightPink
                            grdCol.Header.Appearance.BackColor = Drawing.Color.LightPink
                        Case "APPROVED_ON", "APPROVED_BY"
                            'grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen
                            grdCol.Header.Appearance.BackColor = Drawing.Color.LightGreen
                        Case "INIT_DATE", "INIT_OPER"
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightYellow
                            grdCol.Header.Appearance.BackColor = Drawing.Color.LightYellow
                        Case Else
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightGray
                    End Select
                End If
            Next
        Next

        Select Case MENU_ITEM_OBJECT
            Case "ICFITEMO"
                For Each grdCol As UltraWinGrid.UltraGridColumn In grdICTITEMM.DisplayLayout.Bands(0).Columns
                    Dim key As String = grdCol.Header.Column.Key
                    If lstIctitemmNoEdit.Contains(key) Then
                        ' leave alone
                    ElseIf lstOperations.Contains(key) Then
                        grdCol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        grdCol.CellAppearance.BackColor = Drawing.Color.LightBlue
                    Else
                        grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
                        grdCol.CellAppearance.BackColor = Drawing.Color.LightGray
                    End If

                    Select Case key
                        Case "COMPLETED_ON", "COMPLETED_BY"
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightPink
                        Case "APPROVED_ON", "APPROVED_BY"
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    End Select
                Next
                grdICTITEMM.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
                grdICTITEMM.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdICTITEMM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                splICTITEMM.Panel2Collapsed = True

            Case "ICFITEMM"
                For Each grdCol As UltraWinGrid.UltraGridColumn In grdICTITEMM.DisplayLayout.Bands(0).Columns
                    Dim key As String = grdCol.Header.Column.Key
                    If lstIctitemmNoEdit.Contains(key) Then
                        ' leave alone
                    ElseIf Not lstOperations.Contains(key) Then
                        grdCol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        grdCol.CellAppearance.BackColor = Drawing.Color.LightBlue
                    Else
                        grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
                        grdCol.CellAppearance.BackColor = Drawing.Color.LightGray
                    End If

                    Select Case key
                        Case "COMPLETED_ON", "COMPLETED_BY"
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightPink
                        Case "APPROVED_ON", "APPROVED_BY"
                            grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen
                        Case "ITEM_LOT_CONTROL", "ITEM_SHELF_LIFE_YRS"
                            grdCol.Header.Appearance.BackColor = Color.DarkViolet
                            grdCol.Header.Appearance.BackColor2 = Color.White
                            grdCol.Header.Caption &= " **"
                        Case "ITEM_CONTAINS_ALCOHOL", "HMAT_CODE"
                            grdCol.Header.Appearance.BackColor = Color.Gold
                            grdCol.Header.Appearance.BackColor2 = Color.DarkViolet
                            grdCol.Header.Caption &= " ***"
                    End Select

                    If lstRequiredManualEntry.Contains(key) Then
                        grdCol.Header.Appearance.BackColor = Color.OrangeRed
                        grdCol.Header.Appearance.BackColor2 = Color.White
                        grdCol.Header.Caption &= " *"
                    End If
                Next

                grdICTITEMM.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
                grdICTITEMM.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdICTITEMM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End Select

        SetupIPSAGrid()
        Create_Summary(grdTATIPIC1, "ITEM_CODE", "Count")

        grdICTITEMH.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdICTITEMH.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdICTITEMH.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").Header.Appearance.BackColor = Drawing.Color.LightPink
        grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").CellAppearance.BackColor = Drawing.Color.LightPink

        grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellAppearance.BackColor = Drawing.Color.LightGreen

        grdICTITEMH.DisplayLayout.Bands(0).Columns("COMPLETED").Header.Appearance.BackColor = Drawing.Color.LightPink
        grdICTITEMH.DisplayLayout.Bands(0).Columns("COMPLETED").CellAppearance.BackColor = Drawing.Color.LightPink

        grdICTITEMH.DisplayLayout.Bands(0).Columns("APPROVED").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdICTITEMH.DisplayLayout.Bands(0).Columns("APPROVED").CellAppearance.BackColor = Drawing.Color.LightGreen

        Get_PARM("SOTPARM1")
        SO_PARM_UPC_VENDOR_ID = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""

        SetupSecurity()
        Mode_Settings(False, "")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Update"
                If dst.Tables("ICTITEMM").Rows.Count = 0 AndAlso MENU_ITEM_OBJECT = "ICTITEMM" Then
                    EMsg &= vbCr & "There are no items to update."
                    Exit Select
                ElseIf dst.Tables("ICTITEMM").Rows.Count = 0 And MENU_ITEM_OBJECT = "ICTITEMO" AndAlso dst.Tables("TATIPIC1").Rows.Count = 0 Then
                    EMsg &= vbCr & "There are no items to update."
                    Exit Select
                End If

            Case "Cancel"
                If dst.Tables("ICTITEMM").Rows.Count > 0 OrElse dst.Tables("TATIPIC1").Rows.Count > 0 Then
                    If MessageBox.Show("Do you want to cancel any changes?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Card View"
                If grdICTITEMM.DisplayLayout.Bands(0).CardView = False Then
                    grdICTITEMM.DisplayLayout.Bands(0).CardView = True
                    UltraExplorerBar1.Groups("Screen Control").Items("Card View").Text = "Grid View"
                Else
                    grdICTITEMM.DisplayLayout.Bands(0).CardView = False
                    UltraExplorerBar1.Groups("Screen Control").Items("Card View").Text = "Card View"
                End If

                grdICTITEMM.Refresh()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Card View").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabICTITEMM.Visible = ScreenMode
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each tableName As String In {"ICTITEMM", "ICTITEMH", "TATIPIC1"}
            dst.Tables(tableName).Rows.Clear()
        Next
        EnforceConstraints(True)

        UltraExplorerBar1.Groups("Screen Control").Items("Card View").Text = "Card View"
        grdICTITEMM.DisplayLayout.Bands(0).CardView = False

        SetupSecurity()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Select Case MENU_ITEM_OBJECT
            Case "ICFITEMO" ' Operations
                ASCMAIN1.sql = "Select * FROM ICTITEMM WHERE NVL(COMPLETED, '0') = '1' AND NVL(APPROVED, '0') = '1' AND STATUS = 'E'"

            Case "ICFITEMM" ' Marketing
                ASCMAIN1.sql = "Select * FROM ICTITEMM WHERE STATUS = 'E'"
        End Select

        Fill_Records("ICTITEMM", String.Empty, True, ASCMAIN1.sql)

        Select Case MENU_ITEM_OBJECT
            Case "ICFITEMM" ' Marketing
                Fill_Records("ICTITEMH")
        End Select

        Fill_Records("TATIPIC1", String.Empty, True, "SELECT * FROM TATIPIC1 WHERE IPSA_STATUS_CODE = 'P'")
        grdTATIPIC1.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

        EnforceConstraints(True)

        Sort_grdColumns(grdICTITEMM, "ITEM_CODE")
        Sort_grdColumns(grdICTITEMH, "ITEM_CODE")
        Sort_grdColumns(grdTATIPIC1, "ITEM_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("ICTITEMM")

            If MENU_ITEM_OBJECT = "ICFITEMM" Then
                For Each rowTATIPIC1 As DataRow In dst.Tables("TATIPIC1").Select("")
                    If rowTATIPIC1.RowState = DataRowState.Modified Then
                        rowTATIPIC1.Item("LAST_DATE") = DateTime.Now
                        rowTATIPIC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    End If
                Next
                Update_Record_TDA("TATIPIC1")
            End If

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        If MENU_ITEM_OBJECT = "ICFITEMO" Then
            Load_Popup_Menu(grdICTITEMM, "SSSPBB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns", "Update To Master")
        Else
            Load_Popup_Menu(grdICTITEMM, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns")
        End If

        Load_Popup_Menu(grdICTITEMH, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns")
        Load_Popup_Menu(grdTATIPIC1, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns")
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
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case grdICTITEMM.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Auto Fit Columns"
                grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        End Select

        Select Case grd.Name
            Case grdICTITEMM.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Update To Master"
                Try

                    Dim COMPLETED As String = grd.ActiveRow.Cells("COMPLETED").Value & String.Empty
                    If COMPLETED <> "1" Then
                        MessageBox.Show("Item is not marked Completed.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim APPROVED As String = grd.ActiveRow.Cells("APPROVED").Value & String.Empty
                    If APPROVED <> "1" Then
                        MessageBox.Show("Item is not marked Approved", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    verifyUpdateRow = True
                    grdICTITEMM.PerformAction(UltraGridAction.CommitRow)
                    If Not verifyUpdateRow Then
                        Exit Sub
                    End If

                    Dim lstErrors As New List(Of String)

                    If Not ValidateEntry(grd.ActiveRow, lstErrors, True) Then
                        MessageBox.Show(String.Join(Environment.NewLine, lstErrors.ToArray), "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value & String.Empty
                    Dim ITEM_CODE_CTL As String = grd.ActiveRow.Cells("ITEM_CODE_CTL").Value & String.Empty
                    Dim rowICTITEMM As DataRow = dst.Tables("ICTITEMM").Rows.Find(ITEM_CODE_CTL)
                    If MessageBox.Show($"Do you want to Update item {ITEM_CODE} to the Item Master?", e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    Dim ITEM_DESC As String = rowICTITEMM.Item("ITEM_DESC") & String.Empty
                    Dim ITEM_EAN_CODE As String = rowICTITEMM.Item("ITEM_EAN_CODE") & String.Empty
                    Dim ITEM_UPC_CODE As String = rowICTITEMM.Item("ITEM_UPC_CODE") & String.Empty
                    Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEMM.Item("ITEM_RETAIL_PRICE") & String.Empty)
                    Dim ITEM_ALT_SORT As String = rowICTITEMM.Item("ITEM_ALT_SORT") & String.Empty

                    Dim LAUNCH_DATE As String = rowICTITEMM.Item("LAUNCH_DATE") & String.Empty
                    If IsDate(LAUNCH_DATE) Then
                        LAUNCH_DATE = CDate(LAUNCH_DATE).ToString("dd-MMM-yyyy")
                    Else
                        LAUNCH_DATE = ""
                    End If
                    Dim PROD_CODE As String = rowICTITEMM.Item("PROD_CODE") & String.Empty
                    Dim HC_CODE As String = rowICTITEMM.Item("HC_CODE") & String.Empty

                    Dim COLLECTION_CODE As String = rowICTITEMM.Item("COLLECTION_CODE") & String.Empty
                    Dim ITEM_TYPE_CODE As String = rowICTITEMM.Item("ITEM_TYPE_CODE") & String.Empty
                    Dim ITEM_HIDE_FROM_CAT As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_HIDE_FROM_CAT") & String.Empty))
                    Dim ITEM_CLASS_CODE As String = rowICTITEMM.Item("ITEM_CLASS_CODE") & String.Empty

                    Dim COST_CATGY_CODE As String = rowICTITEMM.Item("COST_CATGY_CODE") & String.Empty
                    Dim ITEM_SNU_CODE As String = COST_CATGY_CODE
                    Dim SEASON_CODE As String = rowICTITEMM.Item("SEASON_CODE") & String.Empty
                    Dim ITEM_BASIC_PROMO As String = rowICTITEMM.Item("ITEM_BASIC_PROMO") & String.Empty
                    Dim COUNTRY_CODE As String = rowICTITEMM.Item("COUNTRY_CODE") & String.Empty
                    Dim HMAT_CODE As String = rowICTITEMM.Item("HMAT_CODE") & String.Empty

                    Dim ITEM_CONTAINS_ALCOHOL As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_CONTAINS_ALCOHOL") & String.Empty))
                    Dim ITEM_LOT_CONTROL As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_LOT_CONTROL") & String.Empty))
                    Dim ITEM_WEIGHT_CHECK As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_WEIGHT_CHECK") & String.Empty))
                    Dim ITEM_APPR_1ST_REC As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_APPR_1ST_REC") & String.Empty))
                    Dim ITEM_CRITICAL_TO_SHIP As String = Math.Abs(Val(rowICTITEMM.Item("ITEM_CRITICAL_TO_SHIP") & String.Empty))

                    Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(rowICTITEMM.Item("ITEM_SHELF_LIFE_YRS") & String.Empty)
                    Dim VEND_CODE As String = rowICTITEMM.Item("VEND_CODE") & String.Empty
                    Dim ITEM_SO_QTY_MIN As Int32 = Val(rowICTITEMM.Item("ITEM_SO_QTY_MIN") & String.Empty)
                    Dim ITEM_SO_QTY_MULT As Int32 = Val(rowICTITEMM.Item("ITEM_SO_QTY_MULT") & String.Empty)
                    Dim INIT_OPER As String = rowICTITEMM.Item("INIT_OPER") & String.Empty
                    Dim LAST_OPER As String = rowICTITEMM.Item("LAST_OPER") & String.Empty
                    Dim INIT_DATE As Date = rowICTITEMM.Item("INIT_DATE")
                    Dim LAST_DATE As Date = rowICTITEMM.Item("LAST_DATE")

                    Dim rowICTPROD1 As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                    Dim ITEM_POS_MAX As Decimal = 0
                    Dim ITEM_POS_MIN As Decimal = 0
                    Dim ITEM_MIN_DAYS_SUPPLY As Decimal = 0
                    Dim ITEM_STD_PACK_SLS As Decimal = 1

                    Dim NRF_SIZE_CODE As String = "00000"
                    Dim NRF_COLOR_CODE As String = "000"
                    Dim ITEM_DESC_CAT As String = String.Empty

                    If rowICTPROD1 IsNot Nothing Then
                        COST_CATGY_CODE = rowICTPROD1.Item("COST_CATGY_CODE") & ""
                        If COST_CATGY_CODE = "N" AndAlso ITEM_MIN_DAYS_SUPPLY >= 0 AndAlso ITEM_POS_MAX > 0 AndAlso ITEM_POS_MIN > 0 AndAlso ITEM_POS_MAX > ITEM_POS_MIN Then
                            ITEM_POS_MAX = Val(rowICTPROD1.Item("PROD_MAX_POS") & "")
                            ITEM_POS_MIN = Val(rowICTPROD1.Item("PROD_MIN_POS") & "")
                            ITEM_MIN_DAYS_SUPPLY = Val(rowICTPROD1.Item("PROD_MIN_DAYS_SUPPLY") & "")
                        End If
                    End If

                    Dim ITEM_PLAN_WASTE_PCT As Decimal = 0
                    Dim rowICTTYPE1 As DataRow = dst.Tables("ICTTYPE1").Rows.Find(ITEM_TYPE_CODE)
                    If rowICTTYPE1 IsNot Nothing Then
                        ITEM_PLAN_WASTE_PCT = Val(rowICTTYPE1.Item("ITEM_WASTE_PCT") & String.Empty)
                    End If

                    Dim SALES_DIVISION_CODE As String = String.Empty
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                    If rowICTCOLL1 IsNot Nothing AndAlso rowICTCOLL1.Item("BRAND_CODE") & String.Empty <> String.Empty Then
                        Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(rowICTCOLL1.Item("BRAND_CODE") & String.Empty)
                        If rowICTBRAN1 IsNot Nothing Then
                            SALES_DIVISION_CODE = rowICTBRAN1.Item("SALES_DIVISION_CODE") & String.Empty
                        End If
                    End If

                    Try
                        BeginTrans()
                        ASCMAIN1.sql = $"Insert into ICTITEM1
                                    (
                                    ITEM_CODE,ITEM_DESC,ITEM_EAN_CODE,ITEM_UPC_CODE,ITEM_RETAIL_PRICE,ITEM_STATUS,
                                    LAUNCH_DATE,PROD_CODE,COLLECTION_CODE,ITEM_TYPE_CODE,ITEM_HIDE_FROM_CAT,ITEM_CLASS_CODE,
                                    COST_CATGY_CODE,ITEM_SNU_CODE,SEASON_CODE,ITEM_BASIC_PROMO,COUNTRY_CODE,
                                    ITEM_CONTAINS_ALCOHOL,ITEM_LOT_CONTROL,ITEM_WEIGHT_CHECK,ITEM_APPR_1ST_REC,ITEM_CRITICAL_TO_SHIP,
                                    ITEM_SHELF_LIFE_YRS,VEND_CODE,ITEM_SO_QTY_MIN,ITEM_SO_QTY_MULT,INIT_OPER,LAST_OPER,INIT_DATE,LAST_DATE,
                                    ITEM_UOM,ITEM_CATGY_CODE,ITEM_STATUS_DATE,ITEM_POS_MAX,ITEM_POS_MIN,ITEM_MIN_DAYS_SUPPLY,ITEM_COST_STATUS,
                                    ITEM_PLAN_WASTE_PCT,SALES_DIVISION_CODE,ITEM_STD_PACK_SLS,ITEM_COST_CURR_CODE,
                                    NRF_SIZE_CODE,NRF_COLOR_CODE,ITEM_DESC_CAT,ITEM_ALT_SORT,HMAT_CODE
                                    ) 
                                values 
                                    (
                                    '{ITEM_CODE}','{ITEM_DESC}','{ITEM_EAN_CODE}','{ITEM_UPC_CODE}',{ITEM_RETAIL_PRICE}, 'A',
                                    '{LAUNCH_DATE}','{PROD_CODE}','{COLLECTION_CODE}','{ITEM_TYPE_CODE}','{ITEM_HIDE_FROM_CAT}','{ITEM_CLASS_CODE}',
                                    '{COST_CATGY_CODE}','{ITEM_SNU_CODE}','{SEASON_CODE}','{ITEM_BASIC_PROMO}','{COUNTRY_CODE}',
                                    '{ITEM_CONTAINS_ALCOHOL}','{ITEM_LOT_CONTROL}','{ITEM_WEIGHT_CHECK}','{ITEM_APPR_1ST_REC}','{ITEM_CRITICAL_TO_SHIP}',
                                    {ITEM_SHELF_LIFE_YRS},'{VEND_CODE}',{ITEM_SO_QTY_MIN},{ITEM_SO_QTY_MULT},'{INIT_OPER}','{LAST_OPER}',:PARM1,:PARM2,
                                    'EA','E',SYSDATE,{ITEM_POS_MAX},{ITEM_POS_MIN},{ITEM_MIN_DAYS_SUPPLY},'P',
                                    {ITEM_PLAN_WASTE_PCT},'{SALES_DIVISION_CODE}',{ITEM_STD_PACK_SLS}, 'USD',
                                    '{NRF_SIZE_CODE}','{NRF_COLOR_CODE}','{ITEM_DESC_CAT}','{ITEM_ALT_SORT}','{HMAT_CODE}'
                                    )"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DD", {INIT_DATE, LAST_DATE})

                        ' Need to update oracle with the values provided by the Operations User
                        ' 04/21/2025 meeting with IPLB, the original code was deleting from ICTITEMM
                        Fill_Records("ICTITEMM_UPDATE", ITEM_CODE_CTL)
                        Dim rowICTITEMM_UPDATE As DataRow = dst.Tables("ICTITEMM_UPDATE").Rows(0)
                        For Each dCol As DataColumn In dst.Tables("ICTITEMM_UPDATE").Columns
                            rowICTITEMM_UPDATE.Item(dCol.ColumnName) = rowICTITEMM.Item(dCol.ColumnName)
                        Next
                        rowICTITEMM_UPDATE.Item("STATUS") = "U"
                        rowICTITEMM_UPDATE.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowICTITEMM_UPDATE.Item("LAST_DATE") = DateTime.Now

                        ASCMAIN1.sql = "DELETE FROM ICTITEMM WHERE ITEM_CODE_CTL = :PARM1"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {ITEM_CODE_CTL})

                        rowICTITEMM_UPDATE.AcceptChanges()
                        rowICTITEMM_UPDATE.SetAdded()
                        Update_Record_TDA("ICTITEMM_UPDATE")

                        ' Write Audit trail for new item - 04/21/2025 meeting with IPLB
                        Fill_Records("ICTITEM1", ITEM_CODE)
                        If dst.Tables("ICTITEM1").Rows.Count = 1 Then
                            dst.Tables("ICTITEM1").Rows(0).SetAdded()
                            Write_Audit_Trail(dst.Tables("ICTITEM1").Rows(0), "N")
                        End If

                        CommitTrans($"Item {ITEM_CODE} {ITEM_DESC} updated to Item Master.")
                        rowICTITEMM.Delete()
                        rowICTITEMM.AcceptChanges()
                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                Catch ex As Exception
                    MessageBox.Show(ex.Message, e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Case Else

        End Select
    End Sub

#End Region

#Region "grdICTITEMX"

    Private Sub grdICTITEMX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTITEMM.BeforeRowUpdate

        ' This prevents clicking the checkboxes from updating the grid
        If grdICTITEMM.ActiveRow IsNot Nothing Then
            If grdICTITEMM.ActiveCell IsNot Nothing Then
                If grdICTITEMM.ActiveCell.Column.Style = ColumnStyle.CheckBox Then
                    If grdICTITEMM.ActiveCell.Column.Key <> "COMPLETED" AndAlso grdICTITEMM.ActiveCell.Column.Key <> "APPROVED" Then
                        e.Cancel = True
                        Exit Sub
                    Else
                        If grdICTITEMM.ActiveCell.IsInEditMode Then
                            grdICTITEMM.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
                        End If
                    End If
                End If
            End If
        End If

        Dim lstErrors As New List(Of String)
        Dim COMPLETED As Boolean = e.Row.Cells("COMPLETED").Value IsNot Null AndAlso CBool(e.Row.Cells("COMPLETED").Value)
        Dim APPROVED As Boolean = e.Row.Cells("APPROVED").Value IsNot Null AndAlso CBool(e.Row.Cells("APPROVED").Value)

        If Not ValidateEntry(e.Row, lstErrors, False, COMPLETED) Then
            If COMPLETED Then
                MessageBox.Show("Item cannot be Marked Complete for the following reasons." & Environment.NewLine & String.Join(Environment.NewLine, lstErrors.ToArray), "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Row.Cells("COMPLETED").Value = "0"
                e.Row.Cells("COMPLETED_BY").Value = String.Empty
                e.Row.Cells("COMPLETED_ON").Value = DBNull.Value

                e.Row.Cells("APPROVED").Value = "0"
                e.Row.Cells("APPROVED_BY").Value = String.Empty
                e.Row.Cells("APPROVED_ON").Value = DBNull.Value
            Else
                MessageBox.Show(String.Join(Environment.NewLine, lstErrors.ToArray), "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            e.Cancel = True
            verifyUpdateRow = False
            Exit Sub
        End If

        Dim ITEM_CODE_CTL As String = e.Row.Cells("ITEM_CODE_CTL").Value & String.Empty
        If ITEM_CODE_CTL.Length = 0 Then
            ITEM_CODE_CTL = ASCMAIN1.Next_Control_No("ICTITEMM.ITEM_CODE_CTL")
            e.Row.Cells("ITEM_CODE_CTL").Value = ITEM_CODE_CTL
        End If

        ' Make Sure Item Code is upper case
        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & String.Empty
        ITEM_CODE = ITEM_CODE.ToUpper
        e.Row.Cells("ITEM_CODE").Value = ITEM_CODE

        If APPROVED Then
            If Not COMPLETED Then
                e.Row.Cells("APPROVED").Value = "0"
                e.Row.Cells("APPROVED_BY").Value = String.Empty
                e.Row.Cells("APPROVED_ON").Value = DBNull.Value
            End If
        Else
            e.Row.Cells("APPROVED_BY").Value = String.Empty
            e.Row.Cells("APPROVED_ON").Value = DBNull.Value
        End If

        If Not COMPLETED Then
            e.Row.Cells("COMPLETED_BY").Value = String.Empty
            e.Row.Cells("COMPLETED_ON").Value = DBNull.Value
        End If

    End Sub

    Private Sub grdICTITEMM_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTITEMM.AfterRowActivate
        'Dim Hidden As Boolean = grdICTITEMM.ActiveRow IsNot Nothing AndAlso grdICTITEMM.ActiveRow.IsAddRow

        'grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").Hidden = Hidden
        'grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED_ON").Hidden = Hidden
        'grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED_BY").Hidden = Hidden
        'grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").Hidden = Hidden
        'grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED_BY").Hidden = Hidden
        'grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED_ON").Hidden = Hidden
    End Sub

    Private Sub grdICTITEMX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMM.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdICTITEMM, sql_where, False)
    End Sub

    Private Sub grdICTITEMX_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTITEMM.Error
        grdICTITEMM.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdICTITEMM_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdICTITEMM.AfterCellUpdate

        Try
            Select Case e.Cell.Column.Key
                Case "PROD_CODE"
                    Dim PROD_CODE As String = e.Cell.Row.Cells("PROD_CODE").Value & String.Empty

                    If dst.Tables("ICTPROD1").Rows.Find(PROD_CODE) Is Nothing Then
                        e.Cell.Row.Cells("COST_CATGY_CODE").Value = String.Empty
                    Else
                        Dim rowICTPROD1 As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                        Dim COST_CATGY_CODE As String = rowICTPROD1.Item("COST_CATGY_CODE") & String.Empty
                        e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE

                        Dim PROD_BASIC_PROMO As String = rowICTPROD1.Item("PROD_BASIC_PROMO") & String.Empty
                        If PROD_BASIC_PROMO = "B" Or PROD_BASIC_PROMO = "P" Then
                            e.Cell.Row.Cells("ITEM_BASIC_PROMO").Value = PROD_BASIC_PROMO
                        End If
                    End If

                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = e.Cell.Row.Cells("COLLECTION_CODE").Value & String.Empty

                    If dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) Is Nothing Then
                        e.Cell.Row.Cells("HC_CODE").Value = String.Empty
                    Else
                        Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                        Dim HC_CODE As String = rowICTCOLL1.Item("HC_CODE") & String.Empty
                        e.Cell.Row.Cells("HC_CODE").Value = HC_CODE
                    End If


                Case "ITEM_EAN_CODE"
                    Dim ITEM_ALT_SORT As String = e.Cell.Row.Cells("ITEM_ALT_SORT").Value & String.Empty
                    Dim ITEM_EAN_CODE As String = e.Cell.Row.Cells("ITEM_EAN_CODE").Value & String.Empty
                    If ITEM_ALT_SORT.Length = 0 And ITEM_EAN_CODE.Length = 13 Then
                        ITEM_ALT_SORT = Mid(ITEM_EAN_CODE, 7, 6)
                        e.Cell.Row.Cells("ITEM_ALT_SORT").Value = ITEM_ALT_SORT
                    End If

                Case "COMPLETED", "APPROVED"
                    If e.Cell.IsInEditMode Then
                        grdICTITEMM.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
                    End If
                    ToggleEditableColumns(grdICTITEMM, e.Cell.Row, False)

            End Select

        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdICTITEMM_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdICTITEMM.InitializeLayout
        With e.Layout
            .Bands(0).Columns("ITEM_SHELF_LIFE_YRS").MinValue = 0
            .Bands(0).Columns("ITEM_SHELF_LIFE_YRS").MaxValue = 999

            .Bands(0).Columns("ITEM_SO_QTY_MIN").MinValue = 1
            .Bands(0).Columns("ITEM_SO_QTY_MIN").MaxValue = 99999

            .Bands(0).Columns("ITEM_SO_QTY_MULT").MinValue = 1
            .Bands(0).Columns("ITEM_SO_QTY_MULT").MaxValue = 99999

            .Bands(0).Columns("ITEM_RETAIL_PRICE").MinValue = 0
            .Bands(0).Columns("ITEM_RETAIL_PRICE").MaxValue = 999999

            .Bands(0).Columns("LAUNCH_DATE").MinValue = CDate("01/01/2023")
            .Bands(0).Columns("LAUNCH_DATE").MaxValue = DateAdd(DateInterval.Year, 3, DateTime.Now)

        End With
    End Sub

    Private Sub grdICTITEMM_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdICTITEMM.AfterRowInsert
        'e.Row.Cells("REQUIRES_EAN").Value = "1"
    End Sub

    Private Sub grdICTITEMM_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdICTITEMM.BeforeRowActivate
        ToggleEditableColumns(grdICTITEMM, e.Row, True)
    End Sub

    Private Sub grdICTITEMM_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdICTITEMM.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "COMPLETED"
        End Select
    End Sub

    Private Sub grdTATIPIC1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdTATIPIC1.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdTATIPIC1, sql_where, False)
    End Sub

    Private Sub grdTATIPIC1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdTATIPIC1.BeforeRowUpdate

        ' If the item is going to be ignored then do not validate.
        Dim IPSA_STATUS_CODE As String = e.Row.Cells("IPSA_STATUS_CODE").Value & String.Empty
        If IPSA_STATUS_CODE = "I" Then
            Exit Sub
        End If

        Try
            If grdTATIPIC1.ActiveCell IsNot Nothing AndAlso grdTATIPIC1.ActiveCell.IsInEditMode Then
                grdTATIPIC1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
            End If
        Catch ex As Exception
        End Try

        Dim COMPLETED As Boolean = e.Row.Cells("COMPLETED").Value IsNot Null AndAlso CBool(e.Row.Cells("COMPLETED").Value)
        Dim APPROVED As Boolean = e.Row.Cells("APPROVED").Value IsNot Null AndAlso CBool(e.Row.Cells("APPROVED").Value)

        Dim zMsg As String = String.Empty

        ' As per Stepahnie only Validate when the Item is Marked Complete
        If COMPLETED Then
            Dim ITEM_DESC As String = e.Row.Cells("ITEM_DESC").Value & String.Empty
            Dim PROD_CODE As String = e.Row.Cells("PROD_CODE").Value & String.Empty
            Dim COLLECTION_CODE As String = e.Row.Cells("COLLECTION_CODE").Value & String.Empty
            Dim ITEM_RETAIL_PRICE As Decimal = Val(e.Row.Cells("ITEM_RETAIL_PRICE").Value & String.Empty)
            Dim ITEM_VALUE As Decimal = Val(e.Row.Cells("ITEM_VALUE").Value & String.Empty)
            Dim ITEM_COST_STD As Decimal = Val(e.Row.Cells("ITEM_COST_STD").Value & String.Empty)

            e.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD
            e.Row.Cells("ITEM_VALUE").Value = ITEM_VALUE

            If ITEM_DESC.Length = 0 Then
                zMsg &= vbCr & "Item Description is required."
            ElseIf ITEM_DESC.Length > 30 Then
                zMsg &= vbCr & "Item Description may have a maximum of 30 characters."
            End If

            If PROD_CODE.Length > 0 Then
                Dim rowICTPROD1 As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                If rowICTPROD1 Is Nothing Then
                    zMsg &= vbCr & "Invalid or missing Product Code."
                End If
            End If

            If COLLECTION_CODE.Length > 0 Then
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                If rowICTCOLL1 Is Nothing Then
                    zMsg &= vbCr & "Invalid or missing Collection Code."
                End If
            Else
                zMsg &= vbCr & "Collection Code is required."
            End If

            If ITEM_RETAIL_PRICE <0 Then
                zMsg &= vbCr & "Retail Price must be greater/equal 0."
            End If

            If ITEM_COST_STD < 0 Then
                zMsg &= vbCr & "Item Cost Std must bet greater/equal 0"
            End If

            If ITEM_VALUE < 0 Then
                zMsg &= vbCr & "Item Value must be greater/equal 0."
            End If

            Dim COST_CATGY_CODE As String = e.Row.Cells("COST_CATGY_CODE").Value & String.Empty
            If COST_CATGY_CODE = "S" Then
                If ITEM_RETAIL_PRICE <= 0 Then
                    zMsg &= vbCr & "Saleable Items require a Retail Price greater than 0."
                End If
            Else
                If ITEM_RETAIL_PRICE > 0 Then
                    zMsg &= vbCr & "Non-Saleable Items require a Retail Price = 0."
                End If
            End If

            If zMsg.Length > 0 Then
                zMsg = "Item cannot be Marked Complete for the following reasons." & Environment.NewLine & zMsg
                e.Row.Cells("COMPLETED").Value = 0
                e.Row.Cells("APPROVED").Value = 0

                MessageBox.Show(zMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If

        If APPROVED Then
            If Not COMPLETED Then
                zMsg &= vbCr & "An Item cannot be marked Approved before it is marked Completed."
                e.Row.Cells("APPROVED").Value = "0"

                MessageBox.Show(zMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If

        If COMPLETED Then
            If e.Row.Cells("COMPLETED_BY").Value & String.Empty = String.Empty Then
                e.Row.Cells("COMPLETED_BY").Value = ASCMAIN1.USER_ID
            End If
            If e.Row.Cells("COMPLETED_ON").Value & String.Empty = String.Empty Then
                e.Row.Cells("COMPLETED_ON").Value = DateTime.Now
            End If
        Else
            e.Row.Cells("COMPLETED_BY").Value = DBNull.Value
            e.Row.Cells("COMPLETED_ON").Value = DBNull.Value
        End If

        If APPROVED Then
            If e.Row.Cells("APPROVED_BY").Value & String.Empty = String.Empty Then
                e.Row.Cells("APPROVED_BY").Value = ASCMAIN1.USER_ID
            End If

            If e.Row.Cells("APPROVED_ON").Value & String.Empty = String.Empty Then
                e.Row.Cells("APPROVED_ON").Value = DateTime.Now
            End If
        Else
            e.Row.Cells("APPROVED_BY").Value = DBNull.Value
            e.Row.Cells("APPROVED_ON").Value = DBNull.Value
        End If

        If COMPLETED Then
            If APPROVED Then
                If e.Row.Cells("COMPLETED_BY").Value = e.Row.Cells("APPROVED_BY").Value Then
                    e.Row.Cells("APPROVED").Value = 0
                    e.Row.Cells("APPROVED_BY").Value = DBNull.Value
                    e.Row.Cells("APPROVED_ON").Value = DBNull.Value
                End If
            End If
        End If
    End Sub

    Private Sub grdTATIPIC1_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdTATIPIC1.BeforeRowActivate
        ToggleEditableColumns(grdTATIPIC1, e.Row, True)
    End Sub

    Private Sub grdTATIPIC1_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdTATIPIC1.AfterCellUpdate
        ToggleEditableColumns(grdTATIPIC1, e.Cell.Row, False)

        Try
            Select Case e.Cell.Column.Key
                Case "PROD_CODE"
                    Dim PROD_CODE As String = e.Cell.Row.Cells("PROD_CODE").Value & String.Empty

                    If dst.Tables("ICTPROD1").Rows.Find(PROD_CODE) Is Nothing Then
                        e.Cell.Row.Cells("COST_CATGY_CODE").Value = String.Empty
                    Else
                        Dim rowICTPROD1 As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                        Dim COST_CATGY_CODE As String = rowICTPROD1.Item("COST_CATGY_CODE") & String.Empty
                        e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE

                        Dim PROD_BASIC_PROMO As String = rowICTPROD1.Item("PROD_BASIC_PROMO") & String.Empty
                        If PROD_BASIC_PROMO = "B" Or PROD_BASIC_PROMO = "P" Then
                            e.Cell.Row.Cells("ITEM_BASIC_PROMO").Value = PROD_BASIC_PROMO
                        End If
                    End If

                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = e.Cell.Row.Cells("COLLECTION_CODE").Value & String.Empty

                    If dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) Is Nothing Then
                        e.Cell.Row.Cells("HC_CODE").Value = String.Empty
                    Else
                        'Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                        'Dim HC_CODE As String = rowICTCOLL1.Item("HC_CODE") & String.Empty
                        'e.Cell.Row.Cells("HC_CODE").Value = HC_CODE
                    End If

                Case "COMPLETED", "APPROVED"
                    If e.Cell.IsInEditMode Then
                        grdTATIPIC1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
                    End If
                    ToggleEditableColumns(grdICTITEMM, e.Cell.Row, False)

            End Select

        Catch ex As Exception

        End Try


    End Sub

    Private Sub grdTATIPIC1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdTATIPIC1.InitializeRow
        If (e.Row.Cells("ITEM_DESC").Value & String.Empty).ToString.Length > 30 Then
            e.Row.Cells("ITEM_DESC").Appearance.BackColor = Color.LightSalmon
        Else
            e.Row.Cells("ITEM_DESC").Appearance.BackColor = Color.LightBlue
        End If
    End Sub


#End Region

#Region "Form Procedures"

    Private Sub SetupSecurity()
        'Security codes
        'Should have to have ME to check/uncheck completed. And to edit anything else that can be edited on the Marketing screen.
        Select Case MENU_ITEM_OBJECT
            Case "ICFITEMM"
                If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                    With grdICTITEMM.DisplayLayout.Override
                        .AllowAddNew = AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With

                    For Each grdcol As UltraGridColumn In grdICTITEMM.DisplayLayout.Bands(0).Columns
                        grdcol.CellActivation = Activation.NoEdit
                    Next

                    With grdTATIPIC1.DisplayLayout.Override
                        .AllowAddNew = AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With

                    For Each grdcol As UltraGridColumn In grdTATIPIC1.DisplayLayout.Bands(0).Columns
                        grdcol.CellActivation = Activation.NoEdit
                    Next
                End If

                'Should have to have MA to check/uncheck the Approved box (unless you uncheck “Completed” then it should clear approved)
                If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("MA") Then
                    grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.NoEdit
                    grdTATIPIC1.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.NoEdit
                Else
                    grdICTITEMM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit

                    grdTATIPIC1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdTATIPIC1.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit
                End If
        End Select

    End Sub

    Private Sub ToggleEditableColumns(ByRef grd As UltraGrid, ByRef grdrow As UltraGridRow, ByVal BeforeRowActivate As Boolean)

        If MENU_ITEM_OBJECT <> "ICFITEMM" Then
            Exit Sub
        End If

        Try
            If grdrow.Cells("COMPLETED").IsInEditMode Then
                grd.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
            End If
        Catch ex As Exception
        End Try

        Try
            If grdrow.Cells("APPROVED").IsInEditMode Then
                grd.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
            End If
        Catch ex As Exception
        End Try

        Select Case grd.Name
            Case grdICTITEMM.Name
                If Not BeforeRowActivate Then
                    If grdrow.Cells("COMPLETED").Value & String.Empty <> "1" Then
                        grdrow.Cells("COMPLETED_ON").Value = DBNull.Value
                        grdrow.Cells("COMPLETED_BY").Value = DBNull.Value
                        If grdrow.Cells("APPROVED").Value & String.Empty <> "0" Then
                            grdrow.Cells("APPROVED").Value = "0"
                        End If
                    End If

                    If grdrow.Cells("APPROVED").Value & String.Empty <> "1" Then
                        grdrow.Cells("APPROVED_ON").Value = DBNull.Value
                        grdrow.Cells("APPROVED_BY").Value = DBNull.Value
                    Else
                        If grdrow.Cells("APPROVED_ON").Value & String.Empty = String.Empty Then
                            grdrow.Cells("APPROVED_ON").Value = DateTime.Now
                            grdrow.Cells("APPROVED_BY").Value = ASCMAIN1.USER_ID
                        End If
                    End If
                End If

                If (grdrow.Cells("COMPLETED").Value & String.Empty = "1" AndAlso grdrow.Cells("APPROVED").Value & String.Empty = "1") Then
                    If grdrow.Cells("COMPLETED_BY").Value & String.Empty = grdrow.Cells("APPROVED_BY").Value & String.Empty Then
                        grdrow.Cells("APPROVED").Value = "0"
                        grdrow.Cells("APPROVED_ON").Value = DBNull.Value
                        grdrow.Cells("APPROVED_BY").Value = DBNull.Value
                        MessageBox.Show("The same person may not Complete and Approve an Item.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If

                If (grdrow.Cells("COMPLETED").Value & String.Empty = "1" OrElse grdrow.Cells("APPROVED").Value & String.Empty = "1") Then
                    For Each col As UltraGridColumn In grdICTITEMM.DisplayLayout.Bands(0).Columns
                        If Not lstIctitemmNoEdit.Contains(col.Key) Then
                            grdICTITEMM.DisplayLayout.Bands(0).Columns(col.Key).CellActivation = Activation.NoEdit
                        End If
                    Next

                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.AllowEdit
                    End If

                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("MA") Then
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit
                    End If
                Else
                    For Each col As UltraGridColumn In grdICTITEMM.DisplayLayout.Bands(0).Columns
                        If Not lstIctitemmNoEdit.Contains(col.Key) Then
                            If ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                                grdICTITEMM.DisplayLayout.Bands(0).Columns(col.Key).CellActivation = Activation.AllowEdit
                            End If
                        End If
                    Next

                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.NoEdit
                    Else
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.AllowEdit
                    End If

                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("MA") Then
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.NoEdit
                    Else
                        grdICTITEMM.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit
                    End If
                End If

            Case grdTATIPIC1.Name
                If Not BeforeRowActivate Then
                    If grdrow.Cells("COMPLETED").Value & String.Empty <> "1" Then
                        grdrow.Cells("COMPLETED_ON").Value = DBNull.Value
                        grdrow.Cells("COMPLETED_BY").Value = DBNull.Value
                        If grdrow.Cells("APPROVED").Value & String.Empty <> "0" Then
                            grdrow.Cells("APPROVED").Value = "0"
                        End If
                    End If

                    If grdrow.Cells("APPROVED").Value & String.Empty <> "1" Then
                        grdrow.Cells("APPROVED_ON").Value = DBNull.Value
                        grdrow.Cells("APPROVED_BY").Value = DBNull.Value
                    Else
                        If grdrow.Cells("APPROVED_ON").Value & String.Empty = String.Empty Then
                            grdrow.Cells("APPROVED_ON").Value = DateTime.Now
                            grdrow.Cells("APPROVED_BY").Value = ASCMAIN1.USER_ID
                        End If
                    End If
                End If

                If (grdrow.Cells("COMPLETED").Value & String.Empty = "1" AndAlso grdrow.Cells("APPROVED").Value & String.Empty = "1") Then
                    If grdrow.Cells("COMPLETED_BY").Value & String.Empty = grdrow.Cells("APPROVED_BY").Value & String.Empty Then
                        grdrow.Cells("APPROVED").Value = "0"
                        grdrow.Cells("APPROVED_ON").Value = DateTime.Now
                        grdrow.Cells("APPROVED_BY").Value = ASCMAIN1.USER_ID
                        MessageBox.Show("The same person may not Complete and Approve an Item.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If

                If (grdrow.Cells("COMPLETED").Value & String.Empty = "1" OrElse grdrow.Cells("APPROVED").Value & String.Empty = "1") Then
                    For Each col As UltraGridColumn In grdTATIPIC1.DisplayLayout.Bands(0).Columns
                        If lstTatipic1AllowEdit.Contains(col.Key) Then
                            grdTATIPIC1.DisplayLayout.Bands(0).Columns(col.Key).CellActivation = Activation.NoEdit
                        End If
                    Next

                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.AllowEdit
                    End If

                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("MA") Then
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit
                    End If
                Else
                    For Each col As UltraGridColumn In grdTATIPIC1.DisplayLayout.Bands(0).Columns
                        If lstTatipic1AllowEdit.Contains(col.Key) Then
                            If ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                                grdTATIPIC1.DisplayLayout.Bands(0).Columns(col.Key).CellActivation = Activation.AllowEdit
                            End If
                        End If
                    Next

                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("ME") Then
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.NoEdit
                    Else
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("COMPLETED").CellActivation = Activation.AllowEdit
                    End If

                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("MA") Then
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.NoEdit
                    Else
                        grdTATIPIC1.DisplayLayout.Bands(0).Columns("APPROVED").CellActivation = Activation.AllowEdit
                    End If
                End If
        End Select
    End Sub

    Private Function ValidateEntry(ByRef grdRow As UltraWinGrid.UltraGridRow,
                                   ByRef lstErrors As List(Of String),
                                   Optional CreateItemMaster As Boolean = False,
                                   Optional MarkedComplete As Boolean = False) As Boolean

        Dim ITEM_DESC As String = grdRow.Cells("ITEM_DESC").Value & String.Empty
        If ITEM_DESC.Length = 0 Then
            lstErrors.Add("Item Description is required")
        Else
            Dim rx As String = "[^a-zA-Z0-9 .$]" ' Allow Upper/Lower case, numbers, space, dot
            Dim r As New System.Text.RegularExpressions.Regex(rx)
            If r.IsMatch(ITEM_DESC) Then
                lstErrors.Add("Item Description has Special Characters which are not allowed")
            ElseIf CreateItemMaster AndAlso ITEM_DESC.Length > 30 Then
                lstErrors.Add("Item Description may have a maximum of 30 characters")
            End If
        End If

        Dim ITEM_RETAIL_PRICE As Decimal = Val(grdRow.Cells("ITEM_RETAIL_PRICE").Value & String.Empty)
        If ITEM_RETAIL_PRICE < 0 Then
            lstErrors.Add("Item Retail Price must be greater/equal 0")
        End If

        Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(grdRow.Cells("ITEM_SHELF_LIFE_YRS").Value & String.Empty)
        If ITEM_SHELF_LIFE_YRS < 0 Then
            lstErrors.Add("Life (years) must be greater/equal 0")
        End If

        Dim ITEM_SO_QTY_MIN As Int32 = Val(grdRow.Cells("ITEM_SO_QTY_MIN").Value & String.Empty)
        If ITEM_SO_QTY_MIN <= 0 Then
            If CreateItemMaster Then
                lstErrors.Add("SO Min Qty must be greater than 0")
            End If
        End If

        Dim ITEM_SO_QTY_MULT As Int32 = Val(grdRow.Cells("ITEM_SO_QTY_MULT").Value & String.Empty)
        If ITEM_SO_QTY_MULT <= 0 Then
            If CreateItemMaster Then
                lstErrors.Add("SO Qty Mult must be greater than 0")
            End If
        End If

        Dim PROD_CODE As String = grdRow.Cells("PROD_CODE").Value & String.Empty
        Dim COST_CATGY_CODE As String = String.Empty
        If PROD_CODE.Length > 0 Then
            If dst.Tables("ICTPROD1").Rows.Find(PROD_CODE) Is Nothing Then
                lstErrors.Add("Invalid Product Code")
            Else
                Dim rowICTPROD1 As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                COST_CATGY_CODE = rowICTPROD1.Item("COST_CATGY_CODE") & String.Empty
                grdRow.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE

                ' COST_CATGY_CODE is also the ITEM_SNU_CODE
                Select Case COST_CATGY_CODE
                    Case "S"
                        If ITEM_RETAIL_PRICE = 0 Then
                            lstErrors.Add("Retail Price is required (non-zero) for a Saleable Item")
                        End If

                    Case "N", "U"
                        If ITEM_RETAIL_PRICE <> 0 Then
                            lstErrors.Add("Retail Price must be zero for a Non-Saleable Item")
                        End If

                    Case Else
                        lstErrors.Add("Unable to determine SNU")
                End Select

                ' Not sure if this should force the value 
                Dim PROD_BASIC_PROMO As String = rowICTPROD1.Item("PROD_BASIC_PROMO") & String.Empty
                If PROD_BASIC_PROMO = "B" Or PROD_BASIC_PROMO = "P" Then
                    If grdRow.Cells("ITEM_BASIC_PROMO").Value & String.Empty <> String.Empty Then
                        grdRow.Cells("ITEM_BASIC_PROMO").Value = PROD_BASIC_PROMO
                    End If
                End If
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            lstErrors.Add("Product Code is required")
        End If

        Dim COLLECTION_CODE As String = grdRow.Cells("COLLECTION_CODE").Value & String.Empty
        If COLLECTION_CODE.Length > 0 Then
            If dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) Is Nothing Then
                lstErrors.Add("Invalid Collection")
            Else
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                Dim HC_CODE As String = rowICTCOLL1.Item("HC_CODE") & String.Empty
                If HC_CODE.Length > 0 Then
                    grdRow.Cells("HC_CODE").Value = HC_CODE
                End If
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            lstErrors.Add("Collection Code is required")
        End If

        Dim ITEM_TYPE_CODE As String = grdRow.Cells("ITEM_TYPE_CODE").Value & String.Empty
        If ITEM_TYPE_CODE.Length > 0 Then
            If dst.Tables("ICTTYPE1").Rows.Find(ITEM_TYPE_CODE) Is Nothing Then
                lstErrors.Add("Invalid Type")
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            lstErrors.Add("Item Type is required")
        End If

        Dim ITEM_CLASS_CODE As String = grdRow.Cells("ITEM_CLASS_CODE").Value & String.Empty
        If ITEM_CLASS_CODE.Length > 0 Then
            If dst.Tables("ICTCLAS1").Rows.Find(ITEM_CLASS_CODE) Is Nothing Then
                lstErrors.Add("Invalid Class")
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            lstErrors.Add("Class Code is required")
        End If

        Dim SEASON_CODE As String = grdRow.Cells("SEASON_CODE").Value & String.Empty
        If SEASON_CODE.Length > 0 Then
            If dst.Tables("ICTSEAS1").Rows.Find(SEASON_CODE) Is Nothing Then
                lstErrors.Add("Invalid Season")
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            'lstErrors.Add("Season Code is required")
        End If

        Dim COUNTRY_CODE As String = grdRow.Cells("COUNTRY_CODE").Value & String.Empty
        If COUNTRY_CODE.Length > 0 Then
            If dst.Tables("TATCNTRY").Rows.Find(COUNTRY_CODE) Is Nothing Then
                lstErrors.Add("Invalid Origin")
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            ' lstErrors.Add("Country Code is required")
        End If

        Dim VEND_CODE As String = grdRow.Cells("VEND_CODE").Value & String.Empty
        If VEND_CODE.Length > 0 Then
            If dst.Tables("APTVEND1").Rows.Find(VEND_CODE) Is Nothing Then
                lstErrors.Add("Invalid Supplier")
            End If
        ElseIf CreateItemMaster OrElse MarkedComplete Then
            ' lstErrors.Add("Supplier Code is required")
        End If

        ' Does this need to be validated
        Dim LAUNCH_DATE As String = grdRow.Cells("LAUNCH_DATE").Value & String.Empty

        If ITEM_SO_QTY_MULT > 1 AndAlso ITEM_SO_QTY_MIN > 1 Then
            If ITEM_SO_QTY_MIN Mod ITEM_SO_QTY_MULT <> 0 Then
                lstErrors.Add("SO Min Qty must be a multiple of SO Multiple")
            End If
        End If

        Dim ITEM_LOT_CONTROL As String = Math.Abs(Val(grdRow.Cells("ITEM_LOT_CONTROL").Value & ""))
        If ITEM_LOT_CONTROL = "1" AndAlso ITEM_SHELF_LIFE_YRS <= 0 Then
            lstErrors.Add("Cannot have Lot Control if Shelf Life is 0")
        End If

        Dim ITEM_CONTAINS_ALCOHOL As String = Math.Abs(Val(grdRow.Cells("ITEM_CONTAINS_ALCOHOL").Value & String.Empty))
        If ITEM_LOT_CONTROL <> "1" AndAlso ITEM_CONTAINS_ALCOHOL = "1" Then
            lstErrors.Add("Cannot set Hazardous (Contains Alcohol) without Lot Control")
        End If

        Dim HMAT_CODE As String = grdRow.Cells("HMAT_CODE").Value & ""
        If HMAT_CODE.Length > 0 Then
            If dst.Tables("ICTHMAT1").Rows.Find(HMAT_CODE) Is Nothing Then
                lstErrors.Add("Invalid Hazmat Code")
            ElseIf ITEM_CONTAINS_ALCOHOL = "0" Then
                lstErrors.Add("Cannot set Hazmat code when Hazardous (Contains Alcohol) is not checked")
            End If
        ElseIf ITEM_CONTAINS_ALCOHOL = "1" Then
            Dim warningMessage As String = String.Empty
            warningMessage = "Warning: Item Contains Alcohol is checked and you did provide a Hazardous Material Code. Item can be saved without a Hazmat code."
            MessageBox.Show(warningMessage, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        If grdRow.Cells("APPROVED").Value & String.Empty = "1" Then
            If grdRow.Cells("COMPLETED").Value & String.Empty <> "1" Then
                lstErrors.Add("Cannot mark Item Approved if the Item is not marked Completed.")
            End If
        End If

        If CreateItemMaster Then
            Dim rx As String = "[^A-Z0-9]" ' Allow Upper case, numbers
            Dim r As New System.Text.RegularExpressions.Regex(rx)

            Dim ITEM_CODE As String = (grdRow.Cells("ITEM_CODE").Value & String.Empty).ToString.ToUpper.Trim
            grdRow.Cells("ITEM_CODE").Value = ITEM_CODE

            If ITEM_CODE.Length = 0 Then
                lstErrors.Add("Item Code is required")
            ElseIf r.IsMatch(ITEM_CODE) Then
                lstErrors.Add("Item Code has Special Characters which are not allowed. A-Z and 0-9")
            End If

            Dim ITEM_EAN_CODE As String = grdRow.Cells("ITEM_EAN_CODE").Value & String.Empty
            Dim ITEM_UPC_CODE As String = grdRow.Cells("ITEM_UPC_CODE").Value & String.Empty
            If ITEM_EAN_CODE <> "" And ITEM_UPC_CODE <> "" Then
                lstErrors.Add("An Item Cannot have both a UPC and an EAN")
            End If

            If ITEM_EAN_CODE <> "" Then
                Dim rowEAN As DataRow = ASCDATA1.GetDataRow("Select ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE = :PARM1 and ITEM_CODE <> :PARM2", "VV", New Object() {ITEM_EAN_CODE, ITEM_CODE})
                If rowEAN IsNot Nothing Then
                    lstErrors.Add("EAN " & ITEM_EAN_CODE & " already used for Item " & rowEAN.Item("ITEM_CODE"))
                End If
            End If

            If ITEM_UPC_CODE <> "" Then
                Dim rowUPC As DataRow = ASCDATA1.GetDataRow("Select ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE = :PARM1 and ITEM_CODE <> :PARM2", "VV", New Object() {ITEM_UPC_CODE, ITEM_CODE})
                If rowUPC IsNot Nothing Then
                    lstErrors.Add("UPC " & ITEM_UPC_CODE & " already used for Item " & rowUPC.Item("ITEM_CODE"))
                End If
            End If

            If ITEM_EAN_CODE <> "" Then
                If ITEM_EAN_CODE.Length <> 13 Or Format(Val(ITEM_EAN_CODE), "0000000000000") <> ITEM_EAN_CODE Then
                    lstErrors.Add("EAN " & ITEM_EAN_CODE & " must be 13 numeric digits")
                End If
            End If

            If ITEM_UPC_CODE <> "" Then
                If ITEM_UPC_CODE.Length <> 12 Or Format(Val(ITEM_UPC_CODE), "000000000000") <> ITEM_UPC_CODE Then
                    lstErrors.Add("UPC " & ITEM_UPC_CODE & " must be 12 numeric digits")
                End If
            End If

            If grdRow.Cells("REQUIRES_EAN").Value & String.Empty = "1" Then
                If ITEM_UPC_CODE.Length = 0 AndAlso ITEM_EAN_CODE.Length = 0 Then
                    lstErrors.Add("EAN is required")
                End If
            End If

            If COST_CATGY_CODE = "S" Then
                If ITEM_EAN_CODE.Length = 0 AndAlso ITEM_UPC_CODE.Length = 0 Then
                    lstErrors.Add("Saleable Items require an Item EAN or UPC.")
                End If
            End If

            Dim rowICTTIEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTTIEM1 IsNot Nothing Then
                lstErrors.Add("Item Code is already in use.")
            End If

            If lstErrors.Count = 0 Then
                Dim ITEM_ALT_SORT As String = grdRow.Cells("ITEM_ALT_SORT").Value & String.Empty
                If ITEM_ALT_SORT.Length = 0 Then
                    If ITEM_EAN_CODE.Length = 13 Then
                        ITEM_ALT_SORT = Mid(ITEM_EAN_CODE, 7, 6)
                        grdRow.Cells("ITEM_ALT_SORT").Value = ITEM_ALT_SORT
                    Else
                        ITEM_ALT_SORT = Mid(ITEM_CODE, 1, 6)
                        grdRow.Cells("ITEM_ALT_SORT").Value = ITEM_ALT_SORT
                    End If
                End If

                ITEM_ALT_SORT = ITEM_ALT_SORT.ToUpper
                grdRow.Cells("ITEM_ALT_SORT").Value = ITEM_ALT_SORT

                If ITEM_ALT_SORT.Length > 0 Then
                    If r.IsMatch(ITEM_ALT_SORT) Then
                        lstErrors.Add("3PL Item Code has Special Characters which are not allowed")
                    Else
                        ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1
                                        where ICTITEM1.ITEM_CODE <> :PARM1
                                        and ICTITEM1.ITEM_ALT_SORT = :PARM2"
                        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", {ITEM_CODE, ITEM_ALT_SORT})
                        If row IsNot Nothing Then
                            lstErrors.Add($"3PL Item Code {ITEM_ALT_SORT} already on file for Item {row.Item("ITEM_CODE")}")
                        End If
                    End If
                End If
            End If
        End If

        Dim INIT_OPER As String = grdRow.Cells("INIT_OPER").Value & String.Empty
        If INIT_OPER.Length = 0 Then
            grdRow.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            grdRow.Cells("INIT_DATE").Value = DateTime.Now
            grdRow.Cells("STATUS").Value = "E" ' Entry mode
        End If

        If grdRow.Cells("COMPLETED").Value & String.Empty = "1" Then
            If grdRow.Cells("COMPLETED_BY").Value & String.Empty = String.Empty Then
                grdRow.Cells("COMPLETED_BY").Value = ASCMAIN1.USER_ID
                grdRow.Cells("COMPLETED_ON").Value = DateTime.Now
            End If
        Else
            grdRow.Cells("COMPLETED_BY").Value = String.Empty
            grdRow.Cells("COMPLETED_ON").Value = DBNull.Value
        End If

        ' 04/21/2025 meeting. Wants a second marketing person to verify the information
        If grdRow.Cells("APPROVED").Value & String.Empty = "1" Then
            If grdRow.Cells("APPROVED_BY").Value & String.Empty = String.Empty Then
                grdRow.Cells("APPROVED_BY").Value = ASCMAIN1.USER_ID
                grdRow.Cells("APPROVED_ON").Value = DateTime.Now
            End If
        Else
            grdRow.Cells("APPROVED_BY").Value = String.Empty
            grdRow.Cells("APPROVED_ON").Value = DBNull.Value
        End If

        If grdRow.Cells("COMPLETED").Value & String.Empty = "1" Then
            If grdRow.Cells("APPROVED").Value & String.Empty = "1" Then
                If grdRow.Cells("COMPLETED_BY").Value & String.Empty = grdRow.Cells("APPROVED_BY").Value & String.Empty Then
                    Dim zmsg As String = $"The same person may not Complete and Approve an Item - Item {(grdRow.Cells("ITEM_CODE").Value & String.Empty).ToString.ToUpper.Trim}"
                    MessageBox.Show(zmsg, "Validate", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    grdRow.Cells("APPROVED").Value = "0"
                    grdRow.Cells("APPROVED_BY").Value = String.Empty
                    grdRow.Cells("APPROVED_ON").Value = DBNull.Value
                End If
            End If
        End If

        grdRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        grdRow.Cells("LAST_DATE").Value = DateTime.Now

        ValidateEntry = lstErrors.Count = 0

    End Function

    Private Function GenerateUPCCode() As String
        If SO_PARM_UPC_VENDOR_ID = "" Then
            MessageBox.Show("No UPC Vendor Prefix (see SO Paramters Table)", "Generate UPC Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End If

        Dim UPC_CODE As String = ASCMAIN1.Next_Control_No("ICTITEM1.ITEM_UPC_CODE")
        Return TAC.SOCMAIN1.UPC(Me, UPC_CODE, SO_PARM_UPC_VENDOR_ID, True)
    End Function

    'Private Function GenerateItemCode(ByVal BRAND_CODE As String)
    '    Return TAC.SOCMAIN1.GenerateItemCode(BRAND_CODE)
    'End Function

    Private Sub SetupIPSAGrid()

        With grdTATIPIC1.DisplayLayout.Bands(0)

            Dim P As Integer = -1
            grdfmt(.Columns("ITEM_CODE"), P, 100, "Item Code", , , True, True)
            grdfmt(.Columns("ITEM_DESC"), P, 170, "Description", Color.LightBlue, , True)
            grdfmt(.Columns("ITEM_ALLOW_HALF_PACK"), P, 60, "Half?", Color.Orange, , True)
            grdfmt(.Columns("PROD_CODE"), P, 70, "Prod", Color.LightBlue, , , , "PROD_CODE")
            grdfmt(.Columns("COLLECTION_CODE"), P, 100, "Coll", Color.LightBlue, , , , "COLLECTION_CODE")
            grdfmt(.Columns("ITEM_RETAIL_PRICE"), P, 70, "Retail", Color.LightBlue, "#.00", )
            grdfmt(.Columns("ITEM_VALUE"), P, 70, "Value", Color.LightBlue, "#.00")
            grdfmt(.Columns("ITEM_COST_STD"), P, 70, "Cost", Color.LightBlue, "#.00")
            grdfmt(.Columns("IPSA_STATUS_CODE"), P, 70, "Status", Color.LightBlue)
            grdfmt(.Columns("MARKETING_NOTES"), P, 170, "Marketing Notes", Color.LightBlue)

            grdfmt(.Columns("COMPLETED"), P, 80, "Completed", Color.LightPink)
            grdfmt(.Columns("COMPLETED_BY"), P, 80, "Completed By", Color.LightPink)
            grdfmt(.Columns("COMPLETED_ON"), P, 80, "Completed On", Color.LightPink, "n")

            grdfmt(.Columns("APPROVED"), P, 80, "Approved", Color.LightGreen)
            grdfmt(.Columns("APPROVED_BY"), P, 80, "Approved By", Color.LightGreen)
            grdfmt(.Columns("APPROVED_ON"), P, 80, "Approved On", Color.LightGreen, "n")

            grdfmt(.Columns("ITEM_ALT_SORT"), P, 80, "3PL Item")
            grdfmt(.Columns("ITEM_EAN_CODE"), P, 120, "EAN")
            grdfmt(.Columns("SIZE_CODE"), P, 80, "Size", Color.LightBlue, , , , "SIZE_CODE")
            grdfmt(.Columns("ITEM_CLASS_CODE"), P, 60, "Class", Color.LightBlue, , , , "ITEM_CLASS_CODE")
            grdfmt(.Columns("COST_CATGY_CODE"), P, 60, "CostCat", Color.LightBlue, , , True, "COST_CATGY_CODE")
            grdfmt(.Columns("ITEM_CATGY_CODE"), P, 60, "Catgy", Color.LightBlue, , , , "ITEM_CATGY_CODE")
            grdfmt(.Columns("SEASON_CODE"), P, 90, "Season", Color.LightBlue, , , , "SEASON_CODE")
            grdfmt(.Columns("LAUNCH_DATE"), P, 100, "Launch", Color.LightBlue, "MM/dd/yy")
            grdfmt(.Columns("ITEM_BASIC_PROMO"), P, 40, "BP", Color.LightBlue)
            grdfmt(.Columns("ITEM_WEIGHT"), P, 70, "Wgt", Color.Gold, "#.00")
            grdfmt(.Columns("ITEM_UNIT_LENGTH"), P, 70, "Len", Color.Gold, "#.00")
            grdfmt(.Columns("ITEM_UNIT_WIDTH"), P, 70, "Wid", Color.Gold, "#.00")
            grdfmt(.Columns("ITEM_UNIT_HEIGHT"), P, 70, "Hgt", Color.Gold, "#.00")
            grdfmt(.Columns("ITEM_STD_PACK_PUR"), P, 70, "Inr-Pur", Color.Orange, "#,##0")
            grdfmt(.Columns("ITEM_STD_PACK_SLS"), P, 70, "Factor", Color.Orange, "#,##0")
            grdfmt(.Columns("ITEM_PO_QTY_MIN"), P, 70, "PO Min", Color.Orange, "#,##0")
            grdfmt(.Columns("ITEM_PO_QTY_MULT"), P, 70, "PO Mult", Color.Orange, "#,##0")
            grdfmt(.Columns("ITEM_SO_QTY_MIN"), P, 70, "SO Min", Color.Orange, "#,##0")
            grdfmt(.Columns("ITEM_SO_QTY_MULT"), P, 70, "SO Mult", Color.Orange, "#,##0")
            grdfmt(.Columns("CARTON_PACK_QTY"), P, 70, "Ctn Pack", Color.Orange, "#,##0")
            grdfmt(.Columns("CASE_WEIGHT_GRS"), P, 70, "Wgt-Ctn", Color.Orange, "#.00")
            grdfmt(.Columns("ITEM_CASE_CUBAGE"), P, 70, "Cub-Ctn", Color.Orange, "#.00")
            grdfmt(.Columns("ITEM_CASE_LENGTH"), P, 70, "Len-Ctn", Color.Orange, "#.00")
            grdfmt(.Columns("ITEM_CASE_WIDTH"), P, 70, "Wid-Ctn", Color.Orange, "#.00")
            grdfmt(.Columns("ITEM_CASE_HEIGHT"), P, 70, "Hgt-Ctn", Color.Orange, "#.00")
            grdfmt(.Columns("ITEM_PALLET_QTY"), P, 70, "Qty-Pal", Color.LightPink, "#,##0")
            grdfmt(.Columns("ITEM_PALLET_WEIGHT"), P, 70, "Wgt-Pal", Color.LightPink, "#.00")
            grdfmt(.Columns("ITEM_PALLET_CUBAGE"), P, 70, "Cub-Pal", Color.LightPink, "#.00")
            grdfmt(.Columns("ITEM_PALLET_LENGTH"), P, 70, "Len-Pal", Color.LightPink, "#.00")
            grdfmt(.Columns("ITEM_PALLET_WIDTH"), P, 70, "Wid-Pal", Color.LightPink, "#.00")
            grdfmt(.Columns("ITEM_PALLET_HEIGHT"), P, 70, "Hgt-Pal", Color.LightPink, "#.00")
            grdfmt(.Columns("ITEM_DESC2"), P, 40, "Desc2")
            grdfmt(.Columns("COUNTRY_CODE"), P, 40, "Country", , , , True)
            grdfmt(.Columns("COMMODITY_CODE"), P, 60, "Commodity", Color.Yellow)
            grdfmt(.Columns("ITEM_CONTAINS_ALCOHOL"), P, 40, "Haz", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_CRITICAL_TO_SHIP"), P, 40, "Crit", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_LOT_CONTROL"), P, 40, "LotCtl", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_WEIGHT_CHECK"), P, 40, "WgtChk", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_APPR_1ST_REC"), P, 40, "ApprRec", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_CRITICAL_TO_SHIP"), P, 40, "Crit", Color.Yellow, "x")
            grdfmt(.Columns("ITEM_SHELF_LIFE_YRS"), P, 40, "Life", Color.Violet, "#,##0")
            grdfmt(.Columns("INIT_DATE"), P, 80, "Created", Color.Violet, "MM/dd/yy", , True)
            grdfmt(.Columns("INIT_OPER"), P, 40, "By", Color.Violet, , , True)
            grdfmt(.Columns("LAST_DATE"), P, 80, "Changed", Color.Violet, "MM/dd/yy", , True)
            grdfmt(.Columns("LAST_OPER"), P, 40, "By", Color.Violet, , , True)

            .Columns("ITEM_ALLOW_HALF_PACK").Style = UltraWinGrid.ColumnStyle.CheckBox
            .Columns("ITEM_ALLOW_HALF_PACK").Hidden = True
        End With

        ASCMAIN1.Add_Value_List(grdTATIPIC1, "IPSA_STATUS_CODE", Nothing, New String() {":", "P:Pending", "I:Ignore"})
        ASCMAIN1.Add_Value_List(grdTATIPIC1, "ITEM_BASIC_PROMO", Nothing, New String() {":", "B:Basic", "P:Promo"})
        ASCMAIN1.Add_Value_List(grdTATIPIC1, "COMMODITY_CODE", "Select COMMODITY_CODE, COMMODITY_DESC from ICTCOMM1")
        ASCMAIN1.Add_Value_List(grdTATIPIC1, "COMPLETED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")
        ASCMAIN1.Add_Value_List(grdTATIPIC1, "APPROVED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")

        For Each grdCol As UltraWinGrid.UltraGridColumn In grdTATIPIC1.DisplayLayout.Bands(0).Columns
            grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Select Case grdCol.Key
                Case "ITEM_DESC", "ITEM_RETAIL_PRICE", "ITEM_VALUE", "ITEM_COST_STD", "MARKETING_NOTES", "IPSA_STATUS_CODE"
                    grdCol.CellActivation = If(MENU_ITEM_OBJECT = "ICFITEMM", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                    grdCol.CellAppearance.BackColor = grdCol.Header.Appearance.BackColor2
                Case "PROD_CODE", "COLLECTION_CODE"
                    grdCol.CellActivation = If(MENU_ITEM_OBJECT = "ICFITEMM", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                    grdCol.CellAppearance.BackColor = grdCol.Header.Appearance.BackColor2
                Case "APPROVED", "COMPLETED"
                    grdCol.CellActivation = If(MENU_ITEM_OBJECT = "ICFITEMM", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                    grdCol.CellAppearance.BackColor = grdCol.Header.Appearance.BackColor2
            End Select
        Next

        For Each grdCol As UltraWinGrid.UltraGridColumn In grdTATIPIC1.DisplayLayout.Bands(0).Columns
            If grdCol.CellActivation = Activation.AllowEdit Then
                lstTatipic1AllowEdit.Add(grdCol.Key)
            End If
        Next

        grdTATIPIC1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdTATIPIC1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

        Select Case MENU_ITEM_OBJECT
            Case "ICFITEMM"
                grdTATIPIC1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Case Else
                grdTATIPIC1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End Select
        grdTATIPIC1.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
    End Sub

    Sub grdfmt(gcol As UltraWinGrid.UltraGridColumn,
               ByRef Position As Integer,
               Width As Integer,
               Caption As String,
               Optional colColor As System.Drawing.Color = Nothing,
               Optional Mask As String = "",
               Optional Fixed As Boolean = False,
               Optional Locked As Boolean = False,
               Optional ViewName As String = "")

        If (Caption & String.Empty).ToString.Trim.Length = 0 Then
            Caption = StrConv(gcol.Key.Replace("_", " "), VbStrConv.ProperCase)
        End If

        With gcol
            Position += 1 : .Header.SetVisiblePosition(Position, False)
            .Width = Width
            .Header.Caption = Caption
            If Mask <> "" Then
                If Mask = "x" Then
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .CellAppearance.TextHAlign = HAlign.Center
                Else
                    .Format = Mask
                End If
            End If
            If Fixed Then
                .Header.Fixed = Fixed
            End If
            If Locked Then
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .CellAppearance.BackColor = Color.Beige
            End If
            If colColor = Nothing Then
                colColor = System.Drawing.Color.LightGray
            End If

            With .Header.Appearance
                .BackColor = Color.White
                .BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = colColor
            End With

            If ViewName <> "" Then
                .Style = UltraWinGrid.ColumnStyle.EditButton
                .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            End If
        End With
    End Sub

#End Region

End Class