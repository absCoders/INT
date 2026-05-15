Imports System.Drawing
Imports System.Math

Public Class ICFPHYS1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            ASCMAIN1.sql = "Select WHSE_CODE, WHSE_DESC, WHSE_LOCATOR, LP_CODE, WHSE_YYYYPP_LAST_PHY, WHSE_PHYS_STATUS from ICTWHSE1"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, True, "", 1, "WHSE_PHYS_STATUS")
            .Tables("ICTWHSE1").Columns.Add("SEL")
            .Tables("ICTWHSE1").Columns("SEL").DefaultValue = "0"
        End With

        grdICTWHSE1.DataSource = dst.Tables("ICTWHSE1")

        Create_Summary(grdICTWHSE1, "WHSE_CODE", "Count")

        With grdICTWHSE1.DisplayLayout.Bands("ICTWHSE1")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If
            Next
            .Columns("SEL").Header.Fixed = True
            .Columns("WHSE_CODE").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdICTWHSE1, "WHSE_PHYS_STATUS", Nothing, New String() {":", ":Not Initialized", "C:Initialized"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Initialize"
                If dst.Tables("ICTWHSE1").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Warehouses Selected"
                End If

                If EMsg = "" Then
                    If MsgBox("This Action will Initialize the Warehouses Selected for Physical Inventory Processing." & vbCrLf _
                              & vbCrLf & vbCrLf & "This includes Clearing all data in the Counts Files for these Warehouses," _
                              & vbCrLf & " and taking a snapshot of the present Book Inventory values by Item/Location" _
                              & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
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

            Case "Initialize"
                Update_Record()
                Fill_Records("ICTWHSE1")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Initialize").Settings.Enabled = not_iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTWHSE1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("ICTWHSE1")
        Sort_grdColumns(grdICTWHSE1, "WHSE_CODE")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        For Each rowICTWHSE1 As DataRow In dst.Tables("ICTWHSE1").Select("SEL = '1'")
            Dim WHSE_CODE As String = rowICTWHSE1.Item("WHSE_CODE")
            rowICTWHSE1.Item("WHSE_PHYS_STATUS") = "C"

            ASCMAIN1.sql = "Delete from ICTPHYC1 where WHSE_CODE = '" & WHSE_CODE & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from ICTPHYC2 where WHSE_CODE = '" & WHSE_CODE & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from WHTLOCB0 where WHSE_CODE = '" & WHSE_CODE & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into WHTLOCB0 Select * from WHTLOCB1 where WHSE_CODE = '" & WHSE_CODE & "'"
            ASCDATA1.ExecuteSQL()

        Next

        Update_Record_TDA("ICTWHSE1")

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSE1, "SS", "Show Filter", "Show GroupBox")
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

        End Select
    End Sub
#End Region 

    Private Sub grdICTWHSE1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTWHSE1.InitializeRow
        If e.Row.Cells("WHSE_PHYS_STATUS").Value & "" <> "" Then
            e.Row.Cells("WHSE_PHYS_STATUS").Appearance.ForeColor = Color.Blue
        End If
    End Sub
End Class