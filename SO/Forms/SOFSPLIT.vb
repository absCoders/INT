Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class SOFSPLIT

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Get_PARM("SOTPARM1")

            Create_TDA(.Tables.Add, "SOTORDR0", "*")

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_STATUS = 'O' AND ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", ASCMAIN1.sql, 0, True, "V")
            .Tables("SOTORDR1").Columns.Add("ADDED", GetType(Int16))
            .Tables("SOTORDR1").Columns("ADDED").DefaultValue = 0

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_STATUS = 'O' AND ORDR_GROUP_NO = :PARM1) AND ORDR_STATUS = 'O'"
            Create_TDA(.Tables.Add, "SOTORDR2", ASCMAIN1.sql, 0, True, "V")

            ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_STATUS = 'O' AND ORDR_GROUP_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTORDR5", ASCMAIN1.sql, 0, True, "V")

            ASCMAIN1.sql = "SELECT DISTINCT ICTWHSE2.ITEM_CODE, ICTWHSE2.WHSE_CODE
                            FROM ICTWHSE1, ICTWHSE2
                            WHERE ICTWHSE1.WHSE_CODE = ICTWHSE2.WHSE_CODE
                            AND LP_CODE IN (SELECT LP_CODE FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1)"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "V", 1)
            dst.Tables("ICTWHSEX").PrimaryKey = New DataColumn() {dst.Tables("ICTWHSEX").Columns("ITEM_CODE")}
        End With

        AUDIT.Add("SOTORDR1", "*")
        AUDIT.Add("SOTORDR2", "*")

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_STATUS")
        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        ASCMAIN1.Add_Value_List(grdSOTORDR2, "ORDR_STATUS")
        Create_Summary(grdSOTORDR2, "ORDR_NO", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Dim ORDR_GROUP_NO As String = Absx1.txtFor("ORDR_GROUP_NO").Text
                Dim drSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)

                If drSOTORDR0 Is Nothing Then
                    EMsg &= vbCr & "Invalid Order Group No"
                    Exit Select
                End If

                If drSOTORDR0.Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
                    EMsg &= vbCr & "Order Group represents an EDI order group. Please note that EDI sales orders are not eligible for modification."
                    Exit Select
                End If

                Fill_Records("SOTORDR1", {ORDR_GROUP_NO})
                If dst.Tables("SOTORDR1").Rows.Count = 0 Then
                    EMsg &= vbCr & "The provided Order Group does not have any sales orders with a status of Open."
                    Exit Select
                End If

                Dim CUST_CODE As String = drSOTORDR0.Item("CUST_CODE") & String.Empty

                ' These locks came from SOFORDR1
                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub

                For Each dr As DataRow In dst.Tables("SOTORDR1").Select("")
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", dr.Item("ORDR_NO")) Then Exit Sub
                Next

            Case "Cancel"
                If MessageBox.Show($"Do you want to {eItemKey}?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If MessageBox.Show($"Do you want to {eItemKey}?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        SplitContainer1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTORDR0").Rows.Clear()
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR5").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("ORDR_GROUP_NO").Clear()
        Absx1.txtFor("WHSE_CODE").Clear()
        Absx1.txtFor("CUST_CODE").Clear()
        grdSOTORDR2.Text = "Sales Order Details"

    End Sub

    Private Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        ASCMAIN1.Progress("Now Loading Sales Orders")

        EnforceConstraints(False)
        Fill_Record("SOTORDR0", {HFs("ORDR_GROUP_NO")})
        Fill_Record("SOTORDR1", {HFs("ORDR_GROUP_NO")})
        Fill_Record("SOTORDR2", {HFs("ORDR_GROUP_NO")})
        Fill_Record("SOTORDR5", {HFs("ORDR_GROUP_NO")})

        Dim drSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(HFs("ORDR_GROUP_NO"))
        Dim WHSE_CODE_GROUP As String = drSOTORDR0.Item("WHSE_CODE")
        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE_GROUP
        Absx1.txtFor("CUST_CODE").Text = drSOTORDR0.Item("CUST_CODE")

        Dim SO_PARM_DEF_PICK_WHSE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        Fill_Records("ICTWHSEX", {WHSE_CODE_GROUP})

        For Each drSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            If drSOTORDR1.Item("WHSE_CODE") <> WHSE_CODE_GROUP Then
                drSOTORDR1.Item("WHSE_CODE") = WHSE_CODE_GROUP
                drSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                drSOTORDR1.Item("LAST_DATE") = DateTime.Now
            End If

            Dim ORDR_NO As String = drSOTORDR1.Item("ORDR_NO")

            For Each drSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim ITEM_CODE As String = drSOTORDR2.Item("ITEM_CODE")

                Dim rowICTWHSEX As DataRow = dst.Tables("ICTWHSEX").Rows.Find(ITEM_CODE)
                If rowICTWHSEX IsNot Nothing Then
                    drSOTORDR2.Item("WHSE_CODE") = rowICTWHSEX.Item("WHSE_CODE")
                Else
                    drSOTORDR2.Item("WHSE_CODE") = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                End If
            Next
        Next

        For Each drSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = drSOTORDR1.Item("ORDR_NO") & String.Empty
            SplitOrdersOnItemWarehouse(ORDR_NO)
        Next

        For Each drSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = drSOTORDR1.Item("ORDR_NO") & String.Empty
            If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND ORDR_STATUS = 'O'").Length > 0 Then
                dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "O"
            Else
                dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "C"
            End If
        Next

        Sort_grdColumns(grdSOTORDR1, "CUST_STORE_LOCATION")

        EnforceConstraints(True)

        grdSOTORDR1.DisplayLayout.Override.AllowColMoving = AllowColMoving.WithinBand
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Header.VisiblePosition = 1
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("WHSE_CODE").Header.VisiblePosition = 2
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ORDR_STATUS").Header.VisiblePosition = 3
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ADDED").Style = ColumnStyle.CheckBox
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ADDED").Header.VisiblePosition = 0

        grdSOTORDR2.DisplayLayout.Override.AllowColMoving = AllowColMoving.WithinBand
        grdSOTORDR2.DisplayLayout.Bands(0).Columns("WHSE_CODE").Header.VisiblePosition = 1
        grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_STATUS").Header.VisiblePosition = 2

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dependent_Updates_SOTORDR1(-1, ORDR_NO)
            Next

            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTORDR5")

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dependent_Updates_SOTORDR1(1, ORDR_NO)
            Next

            Dim tbl_ORDR_GROUP_NOs As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), {"ORDR_GROUP_NO"})
            For Each rowORDR_GROUP_NO As DataRow In tbl_ORDR_GROUP_NOs.Select
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {rowORDR_GROUP_NO.Item("ORDR_GROUP_NO")}, New String() {"ORDR_GROUP_NO_IN"})
            Next


            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Restore Canceled Items", "Auto size columns")
        Load_Popup_Menu(grdSOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Auto size columns")
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
            Case "Auto size columns"
                grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

            Case "Restore Canceled Items"
                Dim ORDR_STATUS As String = grd.ActiveRow.Cells("ORDR_STATUS").Value
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value

                For Each drSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE = '{WHSE_CODE}' AND ORDR_STATUS = 'C'")
                    drSOTORDR2.Item("ORDR_STATUS") = "O"
                    drSOTORDR2.Item("ORDR_QTY_OPEN") = drSOTORDR2.Item("ORDR_QTY_ORIG")
                    drSOTORDR2.Item("ORDR_QTY_CANC") = 0
                    If Val(drSOTORDR2.Item("ORDR_QTY") & String.Empty) = 0 Then
                        drSOTORDR2.Item("ORDR_QTY") = drSOTORDR2.Item("ORDR_QTY_OPEN")
                    End If
                Next

                If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND ORDR_STATUS = 'O'").Length > 0 Then
                    dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "O"
                Else
                    dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "C"
                End If

        End Select
    End Sub

#End Region

    Private Sub SplitOrdersOnItemWarehouse(ByVal ORDR_NO As String)

        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
        rowSOTORDR1.Item("ADDED") = 0
        Dim DICT_ORDR_GROUP_NOs As New Dictionary(Of String, String)
        Dim tblDistinct As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), {"WHSE_CODE", "ORDR_GROUP_NO"})
        For Each row As DataRow In tblDistinct.Select()
            If Not DICT_ORDR_GROUP_NOs.Keys.Contains(row.Item("WHSE_CODE")) Then
                DICT_ORDR_GROUP_NOs.Add(row.Item("WHSE_CODE"), row.Item("ORDR_GROUP_NO"))
            End If
        Next

        Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

        If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'").Length = 0 Then
            Exit Sub
        End If

        'Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Select($"ORDR_NO = '{ORDR_NO}'")(0)
        Dim TBL_WHSES As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'"), "WHSE_CODE")
        For Each rowWhses As DataRow In TBL_WHSES.Rows
            Dim WHSE_CODE_DTL As String = rowWhses.Item("WHSE_CODE")
            If Not DICT_ORDR_GROUP_NOs.ContainsKey(WHSE_CODE_DTL) Then
                DICT_ORDR_GROUP_NOs.Add(WHSE_CODE_DTL, ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO"))
            End If
            Dim rowSOTORDR1_new As DataRow = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1_new.ItemArray = rowSOTORDR1.ItemArray
            Dim ORDR_NO_NEW As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
            rowSOTORDR1_new.Item("ORDR_NO") = ORDR_NO_NEW
            rowSOTORDR1_new.Item("WHSE_CODE") = WHSE_CODE_DTL
            rowSOTORDR1_new.Item("ORDR_GROUP_NO") = DICT_ORDR_GROUP_NOs(WHSE_CODE_DTL)
            rowSOTORDR1_new.Item("ADDED") = 1
            rowSOTORDR1_new.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTORDR1_new.Item("INIT_DATE") = DateTime.Now
            rowSOTORDR1_new.Item("LAST_OPER") = rowSOTORDR1_new.Item("INIT_OPER")
            rowSOTORDR1_new.Item("LAST_DATE") = rowSOTORDR1_new.Item("INIT_DATE")

            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_new)

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE = '{WHSE_CODE_DTL}'")
                Dim ROWSOTORDR2_new As DataRow = dst.Tables("SOTORDR2").NewRow
                ROWSOTORDR2_new.ItemArray = rowSOTORDR2.ItemArray
                ROWSOTORDR2_new.Item("ORDR_NO") = ORDR_NO_NEW
                dst.Tables("SOTORDR2").Rows.Add(ROWSOTORDR2_new)
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = rowSOTORDR2.Item("ORDR_QTY_ORIG")
                rowSOTORDR2.Item("ORDR_QTY") = 0
                rowSOTORDR2.Item("ORDR_STATUS") = "C"
                rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_DATE") = DateTime.Now
            Next

            For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim rowSOTORDR5_new As DataRow = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5_new.ItemArray = rowSOTORDR5.ItemArray
                rowSOTORDR5_new.Item("ORDR_NO") = ORDR_NO_NEW
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5_new)
            Next
            'Dim rowSOTCSTO3_new As DataRow = dst.Tables("SOTCSTO3").NewRow
            'rowSOTCSTO3_new.ItemArray = rowSOTCSTO3.ItemArray
            'rowSOTCSTO3_new.Item("ORDR_NO") = ORDR_NO_NEW
            'rowSOTCSTO3_new.Item("CSO_ADDR_LNO") = Val(dst.Tables("SOTCSTO3").Compute("MAX(CSO_ADDR_LNO)", "")) + 1
            'dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3_new)
        Next
    End Sub

    Sub Dependent_Updates_SOTORDR1(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            If S = -1 Then
            Else
                ' Update_Record_TDA("SOTORDR2")
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub

    Sub Update_ICTSTAT2(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN",
                           New Object() {ITEM_CODE, WHSE_CODE,
                                         0, 0, 0,
                                         QTY, 0, 0},
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN",
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in",
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Update_ICTSTAT2_Cancel(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN",
                           New Object() {ITEM_CODE, WHSE_CODE,
                                         0, 0, 0,
                                         QTY, 0, 0},
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN",
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in",
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR1.AfterRowActivate

        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & String.Empty

        Dim dvw2 As DataView = dst.Tables("SOTORDR2").DefaultView
        dvw2.RowFilter = $"ORDR_NO = '{ORDR_NO}'"

        grdSOTORDR2.Text = $"Details for Sales Order {ORDR_NO}"
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

    End Sub

    Private Sub grdSOTORDR2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDR2.InitializeRow
        Try
            Dim WHSE_CODE As String = grdSOTORDR1.ActiveRow.Cells("WHSE_CODE").Value
            If e.Row.Cells("WHSE_CODE").Value <> WHSE_CODE Then
                e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdSOTORDR1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDR1.InitializeRow
        If Val(e.Row.Cells("ADDED").Value & String.Empty) = 1 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
        End If
    End Sub

End Class