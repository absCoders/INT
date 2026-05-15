Imports System.Drawing
Imports System.Math

Public Class SPFPOSF1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GMTPARM1")

        With dst
            'ASCMAIN1.sql = "Select DEPT_CODE, DEPT_DESC, X.MDS from GMTDEPT1, (SELECT SUBSTR(DGC_CODE,1,1) DEPT_CODE, COUNT (*) MDS FROM SPTSPXP1 WHERE GMTDEPT1.DEPT_CODE (+) = X.DEPT_CODE GROUP BY SUBSTR(DGC_CODE,1,1)) X"
            ASCMAIN1.sql = "Select X.DEPT_CODE, GMTDEPT1.DEPT_DESC, X.MDS " _
                & " from GMTDEPT1, (SELECT SUBSTR(DGC_CODE,1,1) DEPT_CODE, COUNT (*) MDS FROM SPTSPXP1 " _
                & " group by SUBSTR(DGC_CODE,1,1)) X " _
                & " where GMTDEPT1.DEPT_CODE (+) = X.DEPT_CODE"

            Create_TDA(.Tables.Add, "SPTPOSFX", "**", 0, False, "", 1)
            .Tables("SPTPOSFX").Columns.Add("SEL")
            .Tables("SPTPOSFX").Columns("SEL").DefaultValue = "0"

            'ASCMAIN1.sql = ""
            'Create_TDA(.Tables.Add, "SPTPOSF1", "**", 0, False, "", 0)

        End With

        grdSPTPOSFX.DataSource = dst.Tables("SPTPOSFX")
        '  grdSPTPOSF1.DataSource = dst.Tables("SPTPOSF1")

        Create_Summary(grdSPTPOSFX, "DEPT_CODE", "Count")

        'With grdSPTPOSFX.DisplayLayout.Bands("ICTWHSE1")
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        If gcol.Key = "SEL" Then
        '            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
        '        Else
        '            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '            gcol.CellAppearance.BackColor = Color.Beige
        '        End If
        '    Next
        '    .Columns("SEL").Header.Fixed = True
        '    .Columns("DEPT_CODE").Header.Fixed = True
        'End With

        ' ASCMAIN1.Add_Value_List(grdSPTPOSFX, "WHSE_PHYS_STATUS", Nothing, New String() {":", ":Not Initialized", "C:Initialized"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If dst.Tables("SPTPOSFX").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Departments Selected"
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
                Load_Record()
                Mode_Settings(True)

            Case "Create File"
                Update_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Create File").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
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
        For Each TABLE_NAME As String In New String() {"SPTPOSFX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("SPTPOSFX")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        ' load data

        Dim OP_GROUP_CODE As String = ""
        Dim DEPT_CODE As String = ""
        Dim SEASON_CODE As String = ""
        Dim OPS_YYYYWW As String = ""

        Create_File(OP_GROUP_CODE, DEPT_CODE, SEASON_CODE, OPS_YYYYWW)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        ' copy file to folder
        ' make links

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
        Load_Popup_Menu(grdSPTPOSFX, "SSB", "Show Filter", "Show GroupBox", "Show this to Bill")
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
            Case "Show this to Bill"
                MsgBox("Hi Bill")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub
#End Region

    Private Sub grdSPTPOSFX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTPOSFX.InitializeRow
        If e.Row.Cells("DEPT_CODE").Value & "" <> "1" Then
            e.Row.Cells("WHSE_PHYS_STATUS").Appearance.ForeColor = Color.Blue
        End If
    End Sub

    Sub Create_File(OP_GRP_CODE As String, _
                    DEPT_CODE As String, _
                    SEASON_CODE As String, _
                    OPS_YYYYWW As String)

        Dim FILENAME As String = OP_GRP_CODE & DEPT_CODE & SEASON_CODE & OPS_YYYYWW
        Dim FOLDERNAME As String = ASCMAIN1.Folders("Temp")

        '00001347
        '11182013
        '136159100699T2
        '203270300699T2
        '203274900699T2

        Using sw As New System.IO.StreamWriter(FOLDERNAME & FILENAME)
            sw.WriteLine("HEADER1")
            sw.WriteLine("HEADER2")
            ASCMAIN1.sql = "Select stuff"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "SKU_NUMBER")
                Dim LINE As String = ""
                sw.WriteLine(LINE)
            Next
        End Using
    End Sub
End Class