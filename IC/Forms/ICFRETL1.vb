Public Class ICFRETL1

    Dim ITEM_CODE As String
    Dim rowICTITEM1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InquiryMode = (MENU_ITEM_OBJECT = "ICFRETLI")

        Get_PARM("ICTPARM1")
        With dst

            Dim UpdateColumns As String = "PROD_CODE,COST_CATGY_CODE,ITEM_SNU_CODE,ITEM_BASIC_PROMO,ITEM_RETAIL_PRICE,ITEM_NEW_RETAIL_PRICE,ITEM_NEW_RETAIL_PRICE_DATE,ITEM_VALUE,LAST_OPER,LAST_DATE"
            Create_TDA(.Tables.Add, "ICTITEM1", "*",,,,, UpdateColumns)

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ITEM_NEW_RETAIL_PRICE, ITEM_NEW_RETAIL_PRICE_DATE from ICTITEM1 where ITEM_NEW_RETAIL_PRICE_DATE is Not Null"
            Create_TDA(.Tables.Add, "ICTRETLN", "**", 0, False)

            Create_TDA(.Tables.Add, "ICTRETLA", "*", 1)
            Create_TDA(dst.Tables.Add, "TATALRT1", "*")
        End With

        grdICTRETLA.DataSource = dst.Tables("ICTRETLA")
        grdICTRETLN.DataSource = dst.Tables("ICTRETLN")
        Set_Read_Only_for_ctl(Absx1.txtFor("PROD_CODE"), True)
        Create_Summary(grdICTRETLN, "ITEM_CODE", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"

                Validate_Code("ITEM_CODE")

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        ITEM_CODE = Absx1.txtFor("ITEM_CODE").Text
                        If Not ASCMAIN1.Logical_Lock("ICTITEM1", ITEM_CODE) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("COST_CATGY_CODE").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Cost Category Code"
                Else
                    Select Case Absx1.optFor("ITEM_SNU_CODE").Value & ""
                        Case "S"
                            If Val(Absx1.numFor("ITEM_RETAIL_PRICE").Value & "") = 0 Then
                                EMsg &= vbCr & "Invalid Value for Retail Price for a Saleable Item"
                            End If

                            If Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE").Value & "" <> "" Then
                                If Val(Absx1.numFor("ITEM_NEW_RETAIL_PRICE").Value & "") = 0 Then
                                    EMsg &= vbCr & "Invalid Value for New Retail Price for a Saleable Item"
                                End If
                            Else
                                If Val(Absx1.numFor("ITEM_NEW_RETAIL_PRICE").Value & "") <> 0 Then
                                    EMsg &= vbCr & "New Retail Price requires an Effective Date"
                                End If
                            End If

                        Case "N", "U"
                            If Val(Absx1.numFor("ITEM_RETAIL_PRICE").Value & "") <> 0 Then
                                EMsg &= vbCr & "Invalid Value for Retail Price for a Non-Saleable Item"
                            End If

                            'If Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE").Value & "" <> "" Then
                            If Val(Absx1.numFor("ITEM_NEW_RETAIL_PRICE").Value & "") <> 0 Then
                                EMsg &= vbCr & "Invalid Value for New Retail Price for a Non-Saleable Item"
                            End If
                            'End If
                        Case Else
                            EMsg &= vbCr & "Unable to determine SNU"
                    End Select
                End If

                If Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE").Value & "" <> "" Then
                    Dim DTE As Date = CDate(Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE").Value)
                    If Format(DTE, "dd") <> "01" Or Format(DTE, "yyyyMM") <= ASCMAIN1.CYM Then
                        EMsg &= vbCr & "New Retail Price Date must be the 1st of a Future Month"
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

            Case "Edit"
                EntryMode = "E"
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
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("View").Visible = InquiryMode
                    .Items("Edit").Visible = Not InquiryMode
                    .Items("Update").Visible = Not InquiryMode
                    .Items("Cancel").Visible = Not InquiryMode
                    .Items("Done").Visible = InquiryMode
                End With
            End With
        End If

        If InquiryMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("PROD_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("COST_CATGY_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_SNU_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_BASIC_PROMO"), True)
            Set_Read_Only_for_ctl(Absx1.numFor("ITEM_RETAIL_PRICE"), True)
            Set_Read_Only_for_ctl(Absx1.numFor("ITEM_NEW_RETAIL_PRICE"), True)
            Set_Read_Only_for_ctl(Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE"), True)
            Set_Read_Only_for_ctl(Absx1.numFor("ITEM_VALUE"), True)
        End If
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpMemoDetails.Visible = tf
        grdICTRETLN.Visible = Not tf
         
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTITEM1", "ICTRETLA", "ICTRETLN"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        rowICTITEM1 = Fill_Record("ICTITEM1", ITEM_CODE)
        Fill_Records("ICTRETLA", ITEM_CODE)
        Sort_grdColumns(grdICTRETLA, "OPS_YYYYPP".ToLower)
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        email_ALERT()

        BeginTrans()
        Write_Audit_Trail(rowICTITEM1, "E")
        Update_Record_TDA("ICTITEM1")
        CommitTrans("Update Complete")
    End Sub

    Sub email_ALERT()

        Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
        Dim ITEM_RETAIL_PRICE_orig As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE", DataRowVersion.Original) & "")

        Dim ITEM_NEW_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE") & "")
        Dim ITEM_NEW_RETAIL_PRICE_orig As Decimal = Val(rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE", DataRowVersion.Original) & "")

        Dim blnITEM_NEW_RETAIL_PRICE_DATE As Boolean = False
        Dim ITEM_NEW_RETAIL_PRICE_DATE As Date
        If rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE") & "" <> "" Then
            ITEM_NEW_RETAIL_PRICE_DATE = rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE") & ""
            blnITEM_NEW_RETAIL_PRICE_DATE = True
        End If

        Dim blnITEM_NEW_RETAIL_PRICE_DATE_orig As Boolean = False
        Dim ITEM_NEW_RETAIL_PRICE_DATE_orig As Date
        If rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE", DataRowVersion.Original) & "" <> "" Then
            ITEM_NEW_RETAIL_PRICE_DATE_orig = rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE", DataRowVersion.Original) & ""
            blnITEM_NEW_RETAIL_PRICE_DATE_orig = True
        End If

        Dim blnITEM_NEW_RETAIL_PRICE_DATE_changed As Boolean = (blnITEM_NEW_RETAIL_PRICE_DATE Or blnITEM_NEW_RETAIL_PRICE_DATE_orig) ' True
        If blnITEM_NEW_RETAIL_PRICE_DATE And blnITEM_NEW_RETAIL_PRICE_DATE_orig Then
            If Format(ITEM_NEW_RETAIL_PRICE_DATE, "yyyyMMdd") = Format(ITEM_NEW_RETAIL_PRICE_DATE_orig, "yyyyMMdd") Then
                blnITEM_NEW_RETAIL_PRICE_DATE_changed = False
            End If
        End If

        If ITEM_RETAIL_PRICE_orig = ITEM_RETAIL_PRICE And ITEM_NEW_RETAIL_PRICE_orig = ITEM_NEW_RETAIL_PRICE And Not blnITEM_NEW_RETAIL_PRICE_DATE_changed Then
            ' no email
            Exit Sub
        End If

        Dim ALERT_SUBJECT As String = ""
        Dim ALERT_MESSAGE As String = ""

        Dim ALERT_EMAIL = ROWs("ICTPARM1").Item("IC_PARM_EMAIL_ALERT") & ""

        Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
        With rowTATALRT1
            Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
            .Item("ALERT_NO") = ALERT_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("FORM_NAME") = "ICTITEM1"
            .Item("FORM_KEY") = ALERT_NO
            .Item("ALERT_EMAIL") = ALERT_EMAIL
            .Item("ALERT_EML") = "1"

            .Item("ALERT_EML_DATE") = DATETIME_STAMP

            ALERT_SUBJECT = "Retail Price Change: " & ITEM_CODE
            .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)

            ALERT_MESSAGE &= vbCrLf _
                        & " Item: " & ITEM_CODE & " : " & rowICTITEM1.Item("ITEM_DESC") & vbCrLf & vbCrLf _
                        & "Item Retail Price" & vbCrLf _
                        & "Previous Value: " & Format(ITEM_RETAIL_PRICE_orig, "$#.00") & vbCrLf _
                        & " Current Value: " & Format(ITEM_RETAIL_PRICE, "$#.00") & vbCrLf & vbCrLf _
                        & "Item Future Retail Price " & vbCrLf _
                        & "Previous Value: " & Format(ITEM_NEW_RETAIL_PRICE_orig, "$#.00") & vbCrLf _
                        & " Current Value: " & Format(ITEM_NEW_RETAIL_PRICE, "$#.00") & vbCrLf & vbCrLf _
                        & "Item Future Price Date " & vbCrLf _
                        & IIf(Not blnITEM_NEW_RETAIL_PRICE_DATE_orig, "", "Previous Value: " & Format(ITEM_NEW_RETAIL_PRICE_DATE_orig, "MM/dd/yyyy") & vbCrLf) _
                        & IIf(Not blnITEM_NEW_RETAIL_PRICE_DATE, "", " Current Value: " & Format(ITEM_NEW_RETAIL_PRICE_DATE, "MM/dd/yyyy"))

            .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
        End With
        dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(ALERT_EMAIL, "Retail Price Change Alert")

        Dim SEND_NO As String = ""
        If ASCMAIN1.Running_in_VS Then
            SEND_NO = "TESTING"
            '   Stop
        Else
            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                    ALERT_SUBJECT, "PC_PRCEXC", True, False, ITEM_CODE, ITEM_CODE, "Item Code", ALERT_MESSAGE)
        End If

        rowTATALRT1.Item("SEND_NO") = SEND_NO
        Update_Record_TDA("TATALRT1")

        TAC.TACMAIN1.Record_Event("ICTITEM1", ITEM_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "RTLALR", "Retail Price Change Alert emailed", SEND_NO, "ICFRETL1")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Edit", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "ITEM_CODE"
                
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
          
        End Select
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTRETLN, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        Select Case grd.Name
            'Case "grdSOTRMAF2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub

#End Region
    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Fill_Records("ICTRETLN")
        Sort_grdColumns(grdICTRETLN, "ITEM_CODE")
        'grdICTRETLN.Text = "Entered in " & cbeYP.Text
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdICTRETLN_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTRETLN.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ITEM_CODE").Text = e.Row.Cells("ITEM_CODE").Value & ""
            Click_Command("Edit")
        End If
    End Sub
     
End Class