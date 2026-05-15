Public Class SOFWORK1

    Dim sqlSOTWORK1 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlSOTWORK1 = "Select SOTWORK1.*" & vbCrLf _
                & ", DECODE(WO_REF_TYPE,'P',POTORDR1.VEND_CODE,'S',SOTORDR1.CUST_CODE,'R',SOTRSRV1.CUST_CODE,'?') REFERENCE1" & vbCrLf _
                & ", DECODE(WO_REF_TYPE,'P',POTORDR1.PO_REFERENCE,'S',SOTORDR1.ORDR_CUST_PO,'R',SOTRSRV1.ORDR_CUST_PO,'?') REFERENCE2" & vbCrLf _
                & ", DECODE(WO_REF_TYPE,'P',POTORDR1.PO_DATE_SHIP_BY,'S',SOTORDR1.ORDR_SHIP_DATE,'R',SOTRSRV1.ORDR_SHIP_DATE,'?') DATE1" & vbCrLf _
                & ", DECODE(WO_REF_TYPE,'P',POTORDR1.PO_DATE_ETA,'S',SOTORDR1.ORDR_CANCEL_DATE,'R',SOTRSRV1.ORDR_CANCEL_DATE,'?') DATE2" & vbCrLf _
                & " from SOTWORK1,SOTORDR1,SOTRSRV1,POTORDR1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO (+) = SOTWORK1.WO_REF_NO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO (+) = SOTWORK1.WO_REF_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO (+) = SOTWORK1.WO_REF_NO" & vbCrLf
            ASCMAIN1.sql = sqlSOTWORK1 & "   and SOTWORK1.WO_STATUS = 'O'" & vbCrLf
            Create_TDA(.Tables.Add, "SOTWORK1", "**", 0, True, "", 0)

            ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO in " _
                & " (Select WO_NO from SOTWORK1 where WO_REF_TYPE = :PARM1 and WO_REF_NO = :PARM2)"
            Create_TDA(.Tables.Add, "SOTWORK2", "**", 0, , "VV", 1)

        End With

        grdSOTWORK1.DataSource = dst.Tables("SOTWORK1")
        Create_Summary(grdSOTWORK1, "WO_NO", "Count")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTWORK1}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        Next

        With grdSOTWORK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "WO_NO" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"WO_REF_TYPE", "WO_REF_NO", "REFERENCE1", "REFERENCE2", "DATE1", "DATE2"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
            .Columns("WO_NO").Header.Fixed = True
            .Columns("WO_DESC").Header.Fixed = True
            .Columns("WO_DUE").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdSOTWORK1, "WO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdSOTWORK1, "WO_TYPE", "Select WO_TYPE, WO_TYPE_DESC from SOTWORKT")
        ASCMAIN1.Add_Value_List(grdSOTWORK1, "WO_REF_TYPE", Nothing, New String() {":", "P:Purchase Order", "R:Reservation", "S:Sales Order"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Load_SOTWORK1()

            Case "Print"
                Print_Record()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    '.Items("View").Visible = (EntryMode = "L" Or Not ScreenMode)
                    '.Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSOTWORK1.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"SOTWORK1"} ' , "POTSHIP3"}
            ' dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        txtUSER_ID.Text = ASCMAIN1.USER_ID

        Load_SOTWORK1()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Generate_Report("PORWREC2")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = "USER_ID in (Select Distinct WO_ASSIGNED_TO from SOTWORK1 union Select Distinct WO_ASSIGNED_TO from SOTWORK2)"
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTWORK1, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Sales Order Inquiry", "Sales Reservation Inquiry", "email Work Order")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTWORK1"
                tlb_btn = DirectCast(tlb_pop.Tools("PO Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("WO_REF_TYPE").Value = "P"
                tlb_btn = DirectCast(tlb_pop.Tools("Sales Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("WO_REF_TYPE").Value = "S"
                tlb_btn = DirectCast(tlb_pop.Tools("Sales Reservation Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("WO_REF_TYPE").Value = "R"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim WO_REF_TYPE As String = grd.ActiveRow.Cells("WO_REF_TYPE").Value
                Dim WO_REF_NO As String = grd.ActiveRow.Cells("WO_REF_NO").Value
                If WO_REF_TYPE = "P" Then Context_Launch("View", WO_REF_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "Sales Order Inquiry"
                Dim WO_REF_TYPE As String = grd.ActiveRow.Cells("WO_REF_TYPE").Value
                Dim WO_REF_NO As String = grd.ActiveRow.Cells("WO_REF_NO").Value
                If WO_REF_TYPE = "S" Then Context_Launch("View", WO_REF_NO, e.Tool.Key, "SOFORDRI", "F", "SOE")

            Case "Sales Reservation Inquiry"
                Dim WO_REF_TYPE As String = grd.ActiveRow.Cells("WO_REF_TYPE").Value
                Dim WO_REF_NO As String = grd.ActiveRow.Cells("WO_REF_NO").Value
                If WO_REF_TYPE = "R" Then Context_Launch("View", WO_REF_NO, e.Tool.Key, "SOFRSRVI", "F", "SOE")

            Case "email Work Order"

                Dim WO_REF_TYPE As String = grdSOTWORK1.ActiveRow.Cells("WO_REF_TYPE").Value
                Dim WO_REF_NO As String = grdSOTWORK1.ActiveRow.Cells("WO_REF_NO").Value
                Dim REFERENCE1 As String = grdSOTWORK1.ActiveRow.Cells("REFERENCE1").Value & ""
                Dim REFERENCE2 As String = grdSOTWORK1.ActiveRow.Cells("REFERENCE2").Value & ""
                Dim DATE1 As Date = grdSOTWORK1.ActiveRow.Cells("DATE1").Value & ""
                Dim DATE2 As Date = grdSOTWORK1.ActiveRow.Cells("DATE2").Value & ""
                Dim WO_NO As String = grdSOTWORK1.ActiveRow.Cells("WO_NO").Value
                Dim WO_DATE As Date = grdSOTWORK1.ActiveRow.Cells("WO_DATE").Value
                Dim WO_DUE As Date = grdSOTWORK1.ActiveRow.Cells("WO_DUE").Value

                Dim FILENAME As String = ""
                Dim ATTACHMENT As String = ""
                Dim SUBJECT As String = ""

             
                'FILENAME = "S:\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                'ATTACHMENT = ASCMAIN1.Folders("Temp") & STMT_NO & "." & "PDF"

                Dim REF As String = ""
                Dim BODY As String = grdSOTWORK1.ActiveRow.Cells("WO_DESC").Value & vbCrLf & vbCrLf

                Select Case WO_REF_TYPE
                    Case "P"
                        REF = " PO No " & WO_REF_NO & ", Supplier " & REFERENCE1 & " Ref " & REFERENCE2 & ", Ship " & Format(DATE1, "MM/dd/yy") & ", ETA " & Format(DATE2, "MM/dd/yy")
                        BODY &= "PO No " & WO_REF_NO & vbCrLf _
                            & "Supplier " & REFERENCE1 & vbCrLf _
                            & "Ref " & REFERENCE2 & vbCrLf _
                            & "Ship " & Format(DATE1, "MM/dd/yy") & vbCrLf _
                            & "ETA " & Format(DATE2, "MM/dd/yy")

                    Case "S"
                        REF = " Order No " & WO_REF_NO & ", Customer " & REFERENCE1 & " PO " & REFERENCE2 & ", Ship " & Format(DATE1, "MM/dd/yy") & ", Cancel " & Format(DATE2, "MM/dd/yy")
                        BODY &= "Order No " & WO_REF_NO & vbCrLf _
                            & "Customer " & REFERENCE1 & vbCrLf _
                            & "PO " & REFERENCE2 & vbCrLf _
                            & "Ship " & Format(DATE1, "MM/dd/yy") & vbCrLf _
                            & "Cancel " & Format(DATE2, "MM/dd/yy")

                    Case "R"
                        REF = " Reservation No " & WO_REF_NO & ", Customer " & REFERENCE1 & " PO " & REFERENCE2 & ", Ship " & Format(DATE1, "MM/dd/yy") & ", Cancel " & Format(DATE2, "MM/dd/yy")
                        BODY &= "Reservation No " & WO_REF_NO & vbCrLf _
                            & "Customer " & REFERENCE1 & vbCrLf _
                            & "PO " & REFERENCE2 & vbCrLf _
                            & "Ship " & Format(DATE1, "MM/dd/yy") & vbCrLf _
                            & "Cancel " & Format(DATE2, "MM/dd/yy")

                End Select

                SUBJECT = "Work Order " & WO_NO & REF
                Send_email(FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, REFERENCE1, "Name of " & REFERENCE1, BODY)

        End Select
    End Sub


    Sub Send_email(ByVal FILENAME As String, ByVal ATTACHMENT As String, ByVal SUBJECT As String, ENTITY_CODE As String, ENTITY_NAME As String, BODY As String)

        Dim rowTATMAIL1 As DataRow = LookUp("TATMAIL1", "WO")

        Dim rowASTUSER1_EMAIL_FROM As DataRow = LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM") & "", True)
        Dim rowASTUSER1_EMAIL_BCC As DataRow = LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_BCC") & "", True)

        Dim USER_SIGNATURE As String = _
          rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & vbCrLf _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & vbCrLf, "") _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & vbCrLf, "") _
        & "Tel: " & ASCMAIN1.FormatTel(rowASTUSER1_EMAIL_FROM.Item("USER_TELEPHONE") & "", rowASTUSER1_EMAIL_FROM.Item("USER_EXT") & "") & vbCrLf _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & "" <> "", "Fax: " & ASCMAIN1.FormatTel(rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & "") & vbCrLf, "") _
        & rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & vbCrLf

        If FILENAME <> ATTACHMENT Then
            Try
                If My.Computer.FileSystem.FileExists(ATTACHMENT) Then
                    My.Computer.FileSystem.DeleteFile(ATTACHMENT)
                End If
                My.Computer.FileSystem.CopyFile(FILENAME, ATTACHMENT)
            Catch ex As Exception
                MsgBox("Error Processing Statement File " & ATTACHMENT & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Copy Statement PDF")
                Exit Sub
            End Try
        End If

        Dim CUST_CONTACT As String = "" ' Absx1.txtFor("CUST_CONTACT").Text
        Dim CUST_PHONE As String = "" ' Absx1.CtlFor("CUST_PHONE").Text


        Dim frmTAFSEND1 As New TAFSEND1(Me)
        frmTAFSEND1.SEND_FROM = rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & ""
        frmTAFSEND1.SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
        frmTAFSEND1.SEND_FROM_SIGNATURE = USER_SIGNATURE
        frmTAFSEND1.SEND_TO = "" ' Absx1.txtFor("CUST_EMAIL").Text
        frmTAFSEND1.SEND_TO_NAME = CUST_CONTACT
        frmTAFSEND1.SEND_CC = ""
        'frmTAFSEND1.SEND_CC_NAME = ""
        frmTAFSEND1.SEND_BCC = rowASTUSER1_EMAIL_BCC.Item("USER_EMAIL") & ""
        frmTAFSEND1.SEND_BCC_NAME = rowASTUSER1_EMAIL_BCC.Item("USER_NAME") & ""
        frmTAFSEND1.SEND_SUBJECT = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " " & SUBJECT
        frmTAFSEND1.SEND_BODY = rowTATMAIL1.Item("EMAIL_BODY") & vbCrLf & vbCrLf & BODY
        frmTAFSEND1.SEND_ENTITY_TABLE = "SOTWORK1"
        frmTAFSEND1.SEND_ENTITY_KEY = ENTITY_CODE
        frmTAFSEND1.SEND_ENTITY_NAME = ENTITY_NAME
        frmTAFSEND1.SEND_METHOD = "E"
        frmTAFSEND1.SEND_ENTITY_CAPTION = "Customer"
        frmTAFSEND1.SEND_ATTACHMENT = ATTACHMENT
        frmTAFSEND1.EMAIL_KEY = "WO"

        frmTAFSEND1.ShowDialog()

        frmTAFSEND1.Dispose()
        frmTAFSEND1 = Nothing
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Sub Load_SOTWORK1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        Dim sqlw As String = ""
        If optStatus.Value = "O" Then
            sqlw &= " and SOTWORK1.WO_STATUS = 'O'"
        ElseIf optStatus.Value = "C" Then
            sqlw &= " and SOTWORK1.WO_STATUS = 'C'"
        End If

        If optUser.Value = "A" Then
            sqlw &= " and SOTWORK1.WO_ASSIGNED_TO = '" & txtUSER_ID.Value & "'"
        ElseIf optUser.Value = "L" Then
            sqlw &= " and SOTWORK1.WO_NO in " _
                & " (Select Distinct WO_NO from SOTWORK1 where WO_ASSIGNED_TO = '" & txtUSER_ID.Value & "'" _
                & " union " _
                & "  Select Distinct WO_NO from SOTWORK2 where WO_ASSIGNED_TO = '" & txtUSER_ID.Value & "')"
        End If

        If optWO_REF_TYPE.Value <> "A" Then
            sqlw &= " and SOTWORK1.WO_REF_TYPE = '" & optWO_REF_TYPE.Value & "'"
        End If

        ASCMAIN1.sql = sqlSOTWORK1 & sqlw & vbCrLf
        Fill_Records("SOTWORK1", "", True, ASCMAIN1.sql)

        grdSOTWORK1.Rows.ExpandAll(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSOTWORK1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTWORK1.DoubleClickRow
        If grdSOTWORK1.ActiveRow IsNot Nothing Then
            If grdSOTWORK1.ActiveRow.IsDataRow Then
                Dim WO_REF_TYPE As String = grdSOTWORK1.ActiveRow.Cells("WO_REF_TYPE").Value
                Dim WO_REF_NO As String = grdSOTWORK1.ActiveRow.Cells("WO_REF_NO").Value
                Dim REFERENCE1 As String = grdSOTWORK1.ActiveRow.Cells("REFERENCE1").Value & ""
                Dim REFERENCE2 As String = grdSOTWORK1.ActiveRow.Cells("REFERENCE2").Value & ""
                Dim DATE1 As Date = grdSOTWORK1.ActiveRow.Cells("DATE1").Value & ""
                Dim DATE2 As Date = grdSOTWORK1.ActiveRow.Cells("DATE2").Value & ""

                Dim WO_NO As String = grdSOTWORK1.ActiveRow.Cells("WO_NO").Value

                Dim WO_DATE As Date = grdSOTWORK1.ActiveRow.Cells("WO_DATE").Value
                Dim WO_DUE As Date = grdSOTWORK1.ActiveRow.Cells("WO_DUE").Value

                If optWO.Value = "A" Then
                    Fill_Records("SOTWORK2", New String() {WO_REF_TYPE, WO_REF_NO})

                    Using F As New TAC.SOFWORK1(Me, WO_REF_TYPE, WO_REF_NO, False, REFERENCE1, REFERENCE2, DATE1, DATE2, _
                                                "Work Orders relating to PO " & WO_REF_NO)
                        F.ShowDialog()
                    End Using
                Else

                    ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO = '" & WO_NO & "'"
                    Fill_Records("SOTWORK2", "", , ASCMAIN1.sql)

                    dst.Tables("SOTWORK1").DefaultView.RowFilter = "WO_NO = '" & WO_NO & "'"
                    Using F As New TAC.SOFWORK1(Me, "W", WO_NO, False, WO_REF_TYPE, WO_REF_NO, WO_DATE, WO_DUE, _
                            "Work Order " & WO_NO)
                        F.ShowDialog()
                    End Using
                    dst.Tables("SOTWORK1").DefaultView.RowFilter = ""
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTWORK1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTWORK1.InitializeLayout

    End Sub

    Private Sub optUser_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optUser.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTWORK1()
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTWORK1()
    End Sub

    Private Sub optWO_REF_TYPE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optWO_REF_TYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTWORK1()
    End Sub
End Class