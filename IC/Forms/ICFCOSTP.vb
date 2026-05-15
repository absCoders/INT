Imports Infragistics.Win.UltraWinGrid

Public Class ICFCOSTP

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst

            Create_TDA(.Tables.Add, "ICTCOSTP", "*", 1)
            .Tables("ICTCOSTP").Columns.Add("ITEM_DESC", GetType(String))

            ASCMAIN1.sql = $"Select OPS_YYYYPP, LEGEND from GLTPARM2
                                where OPS_YYYYPP >= '{ASCMAIN1.CYP}'
                                and OPS_YYYYPP <= '{ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12)}'"

            Dim dvw1 As DataView = ASCDATA1.GetDataTable.DefaultView
            dvw1.Sort = "OPS_YYYYPP DESC"
            cbeYP1.DataSource = dvw1
            cbeYP1.ValueMember = "OPS_YYYYPP"
            cbeYP1.DisplayMember = "LEGEND"
            cbeYP1.Value = ASCMAIN1.CYP

            Create_TDA(.Tables.Add, "ICTTRFC1", "*")
            Fill_Records("ICTTRFC1", String.Empty, True, "SELECT * FROM ICTTRFC1")

            Create_TDA(.Tables.Add, "ICTFRTC1", "*")
            Fill_Records("ICTFRTC1", String.Empty, True, "SELECT * FROM ICTFRTC1")

            Create_TDA(.Tables.Add, "ICTITEM1", "*")
        End With

        ASCMAIN1.grdInitializeLayout(grdICTCOSTP)

        grdICTCOSTP.DataSource = dst.Tables("ICTCOSTP")
        Create_Summary(grdICTCOSTP, "ITEM_CODE", "Count")
        ASCMAIN1.Add_Value_List(grdICTCOSTP, "ITEM_COST_TRF_CLASS", "SELECT TRF_CLASS_CODE, TRF_CLASS_CODE || ' - ' || TRF_CLASS_DESC DESCRIP FROM ICTTRFC1")
        ASCMAIN1.Add_Value_List(grdICTCOSTP, "ITEM_COST_FRT_CLASS", "SELECT FRT_CLASS_CODE, FRT_CLASS_CODE || ' - ' || FRT_CLASS_DESC DESCRIP FROM ICTFRTC1")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If MessageBox.Show("Do you want to Update changes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
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

            Case "Show Attachments"
                Show_Attachments()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1

                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Show Attachments").Visible = ScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        grdICTCOSTP.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("ICTCOSTP").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Cost Information")

        EnforceConstraints(False)

        ASCMAIN1.sql = $"SELECT ICTCOSTP.*, ICTITEM1.ITEM_DESC, ICTITEM1.VEND_CODE
                            FROM ICTCOSTP, ICTITEM1
                            WHERE ICTCOSTP.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                            AND ICTCOSTP.OPS_YYYYPP = '{HFs("OPS_YYYYPP")}'"

        Fill_Records("ICTCOSTP", String.Empty, True, ASCMAIN1.sql)
        grdICTCOSTP.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

        If dst.Tables("ICTCOSTP").Rows.Count > 0 Then
            ASCMAIN1.sql = $"SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT ITEM_CODE FROM ICTCOSTP WHERE OPS_YYYYPP = '{HFs("OPS_YYYYPP")}')"
            Fill_Records("ICTITEM1", String.Empty, True, ASCMAIN1.sql)
        End If

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            INIT_LAST("ICTCOSTP", True, "", True)
            Update_Record_TDA("ICTCOSTP", $"OPS_YYYYPP = '{HFs("OPS_YYYYPP")}'")
            CommitTrans()
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCOSTP, "SSB", "Show Filter", "Show GroupBox", "Auto Size Columns")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Auto Size Columns"
                grd.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

        End Select
    End Sub

#End Region

#Region "grdICTCOSTP"

    Private Sub grdICTCOSTP_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTCOSTP.ClickCellButton

        If grdICTCOSTP.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Dim VIEW_NAME As String = ""

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

            Case "ITEM_COST_FRT_CLASS"
                VIEW_NAME = "FRT_CLASS_CODE"

            Case "ITEM_COST_TRF_CLASS"
                VIEW_NAME = "TRF_CLASS_CODE"

            Case Else
                Exit Sub

        End Select

        grdClickCellButton(grdICTCOSTP, sql_where, False, e.Cell.Column.Key, VIEW_NAME)

    End Sub

    Private Sub grdICTCOSTP_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTCOSTP.BeforeRowUpdate

        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & String.Empty
        Dim ITEM_COST_FRT_CLASS As String = e.Row.Cells("ITEM_COST_FRT_CLASS").Value & String.Empty
        Dim ITEM_COST_TRF_CLASS As String = e.Row.Cells("ITEM_COST_TRF_CLASS").Value & String.Empty
        Dim errorMsgs As New List(Of String)

        e.Row.Cells("OPS_YYYYPP").Value = HFs("OPS_YYYYPP")

        Dim drICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        If drICTITEM1 Is Nothing Then
            drICTITEM1 = Fill_Record("ICTITEM1", ITEM_CODE)
        End If

        If drICTITEM1 Is Nothing Then
            errorMsgs.Add("Invalid Item Code")
        Else
            e.Row.Cells("ITEM_DESC").Value = drICTITEM1.Item("ITEM_DESC") & String.Empty
        End If

        Dim drICTTRFC1 As DataRow = dst.Tables("ICTTRFC1").Rows.Find(ITEM_COST_TRF_CLASS)
        If drICTTRFC1 Is Nothing Then
            errorMsgs.Add("Invalid Tariff Class")
        End If

        Dim drICTFRTC1 As DataRow = dst.Tables("ICTFRTC1").Rows.Find(ITEM_COST_FRT_CLASS)
        If drICTFRTC1 Is Nothing Then
            errorMsgs.Add("Invalid Freight Class")
        End If
        If errorMsgs.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, errorMsgs.ToArray), "Update Errors", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

    End Sub

#End Region
    Sub Show_Attachments()
        ' Care that modes is true - in mode settings method will show attachments only if modes is true

        Dim Entity As Dropped_On_Entity = Dropped_On_Context()
        ' instead of an attachmetn for a journl no it's for a period
        Dim OPS_YYYYPP As String = Absx1.cbeFor("OPS_YYYYPP").Value

        Entity.TABLE_NAME = "ICTCOSTP"
        Entity.CODE_VALUE = OPS_YYYYPP
        Entity.COLUMN_NAME = "OPS_YYYYPP"
        Entity.DESC_VALUE = "Period Cost Upload"
        'Entity.OTHER_ENTITIES = New List(Of Dropped_On_Entity_Other)

        Dim F As New ASFATTA1
        F.ENTITY = Entity
        F.ShowDialog()
        F.Dispose()
    End Sub

End Class