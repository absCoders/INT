Imports Infragistics.Win.UltraWinGrid

Public Class SOFORDE1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select * from SOTORDE1 where EDOC_STATUS = :PARM1 and EDOC_DATE_RECEIVED >= :PARM2"
            Create_TDA(.Tables.Add, "SOTORDE1", "**", 0, False, "VV")

            Create_TDA(.Tables.Add, "SOTORDE2", "*", 1, False)

            Create_TDA(.Tables.Add, "SOTORDEE", "*")


            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1, False)

        End With

        grdSOTORDE1.DataSource = dst.Tables("SOTORDE1")
        grdSOTORDE2.DataSource = dst.Tables("SOTORDE2")

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        Create_Summary(grdSOTORDE1, "EDOC_NO", "Count")
        Create_Summary(grdSOTORDE2, "EDOC_LNO", "Count")

        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")

        Show_Filter(grdSOTORDE1)


        optStatus.Value = "P"
        dteSince.DateTime = Now.AddDays(-60)


        ASCMAIN1.Add_Value_List(grdSOTORDE1, "EDOC_STATUS")

    End Sub


    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"


            Case "Done"

            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Refresh"
                Refresh_Screen()

            Case "Load"
                'EntryMode = "L"
                'Load_Record()
                'Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                ' date_Record()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Refresh").Settings.Enabled = not_iScreenMode
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                '.Items("Update").Settings.Enabled = iScreenMode

                ' UNTIL WE DECIDE THAT WE HAVE A USE CASE FOR FLIPPING TO TRUE AND SHOWING MORE DETAUL
                .Items("Load").Visible = False
                .Items("Done").Visible = False

            End With

            .Groups("Status").Visible = Not ScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        ' grdSOTORDE1.Visible = ScreenMode

        If ScreenMode And EntryMode = "E" Then
            For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTORDE1, grdSOTORDE2, grdSOTORDR1, grdSOTORDR2}
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grd.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            Next
        Else
            For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTORDE1, grdSOTORDE2, grdSOTORDR1, grdSOTORDR2}
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grd.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            Next
        End If

        If Not ScreenMode Then
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDE1", "SOTORDE1", "SOTORDR1", "SOTORDR1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTORDE1)
        Clear_All_Filters(grdSOTORDE2)
        Clear_All_Filters(grdSOTORDR1)
        Clear_All_Filters(grdSOTORDR2)


        Refresh_Screen()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Fill_Records("SOTORDE1")
        Fill_Records("SOTORDE2")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Updating...")


        'MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Verification")

        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDE1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show email", "Reject/Delete")
        Load_Popup_Menu(grdSOTORDE2, "B", "Show PDF", "Show XLS")
        Load_Popup_Menu(grdSOTORDR1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
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

        For Each tool As Infragistics.Win.UltraWinToolbars.ToolBase In tlb_pop.Tools

        Next

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTORDE1"
                    tlb_btn = DirectCast(tlb.Tools("Reject/Delete"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (optStatus.Value = "P") Or (optStatus.Value = "E")

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
            Case "Show email"
                Dim EDOC_NO As String = grd.ActiveRow.Cells("EDOC_NO").Value
                Dim FILENAME As String = $"{EDOC_NO}.EML"
                Dim FOLDERNAME As String = ROWs("SOTPARM1").Item("SO_PARM_EDOC_FOLDER") & "EML\"
                Show_Document(FOLDERNAME & FILENAME)

            Case "Show PDF", "Show XLS"

                Dim EXT As String = Mid(e.Tool.Key, e.Tool.Key.Length - 2, 3)
                Dim EXTF As String = EXT
                If EXTF = "XLS" Then EXTF = "XLSX"
                Dim EDOC_NO As String = grd.ActiveRow.Cells("EDOC_NO").Value
                Dim EDOC_LNO As Int32 = Val(grd.ActiveRow.Cells("EDOC_LNO").Value & "")

                Dim FILENAME As String = $"{EDOC_NO}-{Format(EDOC_LNO, "000")}.{EXTF}"
                Dim FOLDERNAME As String = ROWs("SOTPARM1").Item("SO_PARM_EDOC_FOLDER") & $"{EXT}\"
                Show_Document(FOLDERNAME & FILENAME)

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Reject/Delete"
                MsgBox("To be developed after discussion")

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region


    Sub Refresh_Screen()

        Fill_Records("SOTORDE1", New Object() {optStatus.Value, Format(dteSince.DateTime, "dd-MMM-yyyy")})
        Sort_grdColumns(grdSOTORDE1, "EDOC_NO".ToLower)

        'If Not ASCMAIN1.Logical_Lock("SOCORDE1", "*") Then
        '    Exit Sub
        'End If

        ' Set_ScreenMode_Base(False)
        'ASCMAIN1.MultiTask_Release()

    End Sub

    Private Sub grdSOTORDE1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDE1.AfterRowActivate

        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()

        'splSOTORDRX.Visible = False
        grdSOTORDR1.Visible = False
        grdSOTORDR2.Visible = False

        Dim EDOC_NO As String = grdSOTORDE1.ActiveRow.Cells("EDOC_NO").Text
        Fill_Records("SOTORDE2", EDOC_NO)
        Sort_grdColumns(grdSOTORDE2, "EDOC_LNO")
        grdSOTORDE2.Text = "Files attached to email for EDOC No " & EDOC_NO
    End Sub

    Private Sub grdSOTORDE2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDE2.AfterRowActivate
        Dim ORDR_GROUP_NO As String = grdSOTORDE2.ActiveRow.Cells("ORDR_GROUP_NO").Text
        Fill_Records("SOTORDR1", ORDR_GROUP_NO)
        Sort_grdColumns(grdSOTORDR1, "ORDR_NO")
        grdSOTORDR1.Text = "Sales Order Headers for Order Group No " & ORDR_GROUP_NO
        grdSOTORDR1.Visible = True

    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR1.AfterRowActivate
        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Text
        Fill_Records("SOTORDR2", ORDR_NO)
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
        grdSOTORDR2.Text = "Sales Order Details for Order No " & ORDR_NO
        grdSOTORDR2.Visible = True
    End Sub

    Private Sub optStatus_ValueChanged(sender As Object, e As EventArgs) Handles optStatus.ValueChanged

        lblSince.Visible = (optStatus.Value = "C") Or (optStatus.Value = "X")
        dteSince.Visible = (optStatus.Value = "C") Or (optStatus.Value = "X")

        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Screen()
    End Sub
End Class