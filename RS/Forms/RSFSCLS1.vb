Public Class RSFSCLS1

    Dim RSTSCLS1 As String
    Dim RSTRETLX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Create_Work_Tables()
        With dst
            ASCMAIN1.sql = "Select * from " & RSTSCLS1
            Create_TDA(.Tables.Add, "RSTSCLS1", "**", 0, False)
            .Tables("RSTSCLS1").Columns.Add("NEW_CLS_CALC", GetType(System.String), "IIF(CUST_STORE_CLASS_CODE='NEW' and PRDS>12,'COMP','')")
            .Tables("RSTSCLS1").Columns.Add("NEW_CLS", GetType(System.String))

            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2," & RSTSCLS1 & " RSTSCLS1" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = RSTSCLS1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = RSTSCLS1.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "", 2, "CUST_STORE_CLASS_CODE")
        End With

        grdRSFSCLS1.DataSource = dst.Tables("RSTSCLS1")

        Create_Summary(grdRSFSCLS1, "CUST_CODE", "Count")
        Show_Filter(grdRSFSCLS1, True)

        With grdRSFSCLS1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSFSCLS1.DisplayLayout.Bands(0).Columns
            If gcol.Key = "NEW_CLS" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        ASCMAIN1.Add_Value_List(grdRSFSCLS1, "NEW_CLS", Nothing, New String() {":", "DOTCOM:DOTCOM", "NEW:NEW", "COMP:COMP"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                 
            Case "Update"
                 
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
 
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdRSFSCLS1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"RSTSCLS1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        Create_Work_Tables()
        Fill_Records("RSTSCLS1")
        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("RSTSCLS1").Select("PRDS>=12 AND CUST_STORE_CLASS_CODE = 'NEW'")
            row.Item("NEW_CLS") = row.Item("NEW_CLS_CALC")
        Next

        Sort_grdColumns(grdRSFSCLS1, "CUST_CODE, CUST_STORE_NO")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Fill_Records("ARTCUST2")
        For Each row As DataRow In dst.Tables("RSTSCLS1").Select _
            ("ISNULL(NEW_CLS,'?') <> ISNULL(CUST_STORE_CLASS_CODE,'?')", "")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim NEW_CLS As String = row.Item("NEW_CLS") & ""
            If NEW_CLS <> "" Then
                Dim CUST_STORE_CLASS_CODE As String = row.Item("CUST_STORE_CLASS_CODE") & ""
                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                rowARTCUST2.Item("CUST_STORE_CLASS_CODE") = NEW_CLS
                Write_Audit_Trail(rowARTCUST2, "E")
            End If
        Next
        Update_Record_TDA("ARTCUST2")

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

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSFSCLS1, "SSBBB", "Show Filter", "Show GroupBox", "Set to NEW", "Set to COMP", "Set to DOTCOM")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdRSTBUDFX"
  
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
            Case "Set to NEW", "Set to COMP", "Set to DOTCOM"
                Dim NEW_CLS As String = Split(e.Tool.Key, " ")(2)
                For Each grow As UltraWinGrid.UltraGridRow In grdRSFSCLS1.Selected.Rows
                    grow.Cells("NEW_CLS").Value = NEW_CLS
                    grow.Update()
                Next

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "CUST_CODE"
            '    If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If EntryMode = "" Then
            '        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            '            LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
            '            If cdr IsNot Nothing Then

            '            End If
            '        End If
            '    End If
        End Select
    End Sub
     
#End Region
 
    Sub Create_Work_Tables()

        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
            & ", Min (OPS_YYYYPP) YPMIN, Max (OPS_YYYYPP) YPMAX" & vbCrLf _
            & " from RSTRETL1 where QTY_SOLD > 0 group by CUST_CODE, CUST_STORE_NO"

        If RSTRETLX = "" Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "QTY_SOLD > 0", "QTY_SOLD > 0 and ROWNUM < 1")
            RSTRETLX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTRETLX & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTRETLX)
            ASCDATA1.ExecuteSQL("Insert into " & RSTRETLX & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_CLASS_CODE" & vbCrLf _
            & ", X.YPMIN, X.YPMAX, 0 PRDS" & vbCrLf _
            & " from ARTCUST2, " & RSTRETLX & " X" & vbCrLf _
            & " where ARTCUST2.CUST_CODE in (Select Distinct CUST_CODE from RSTRETL1)" & vbCrLf _
            & "   and X.CUST_CODE (+) = ARTCUST2.CUST_CODE and X.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO"

        If RSTSCLS1 = "" Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            RSTSCLS1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSCLS1 & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
            ASCDATA1.ExecuteSQL("Create Index I_" & RSTSCLS1 & "_1 on " & RSTSCLS1 & " (YPMIN,YPMAX)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSCLS1)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSCLS1 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select Distinct YPMIN, YPMAX from " & RSTSCLS1 & " where YPMIN is Not Null and YPMAX is Not Null"
            For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim YPMIN As String = ROW.Item("YPMIN") & ""
                Dim YPMAX As String = ROW.Item("YPMAX") & ""
                Dim PRDS As Integer = ASCMAIN1.Period_Diff(YPMIN, ASCMAIN1.CYP) + 1
                ASCMAIN1.sql = "Update " & RSTSCLS1 & " set PRDS = " & CStr(PRDS) & " where YPMIN = '" & YPMIN & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If
    End Sub

    Private Sub grdRSFSCLS1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdRSFSCLS1.InitializeRow
        If e.Row.Cells("CUST_STORE_CLASS_CODE").Value & "" <> e.Row.Cells("NEW_CLS").Value & "" Then
            e.Row.Cells("NEW_CLS").Appearance.ForeColor = Color.Red
        Else
            e.Row.Cells("NEW_CLS").Appearance.ForeColor = Color.Empty
        End If
    End Sub
End Class