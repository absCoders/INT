Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class ICFIADJ1

    Dim rowICTIADJ1 As DataRow
    Dim location_support As Boolean = False
    Dim processing3PL As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIADJI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTIADJ1.*" _
            & " from ICTIADJ1 where ICTIADJ1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTIADJX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            & ", ICTIADJ1.ADJ_DATE, ICTIADJ1.WHSE_CODE, ICTIADJ1.REASON_CODE" _
            & ", ICTIADJ1.ADJ_NOTE, ICTIADJ1.INIT_OPER, ICTIADJ1.INIT_DATE" _
            & ", ICTIADJ1.ADJ_SOURCE, ICTIADJ1.OPS_YYYYPP, ICTIADJ1.RTRN_NO" _
            & " from ICTIADJ1,ICTIADJ3,GLTACCT1 where ICTIADJ1.OPS_YYYYPP = :PARM1" _
            & " and GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE" _
            & " and ICTIADJ3.ADJ_NO = ICTIADJ1.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")

            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTITEM1.ITEM_DESC" _
            & " from ICTIADJ2,ICTITEM1 where ICTITEM1.ITEM_CODE = ICTIADJ2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ2", "**", 1)
            .Tables("ICTIADJ2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(ADJ_QTY,0) * ISNULL(ITEM_COST_STD,0)")

            ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIADJ3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("ICTIADJ0")
            .Tables("ICTIADJ0").Columns.Add("KEY")
            .Tables("ICTIADJ0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTREAS1"
            Create_TDA(.Tables.Add, "ICTREAS1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            ASCMAIN1.sql = "SELECT '0' SEL, EDTTRXN1.*, ICTITEM1.ITEM_DESC"
            ASCMAIN1.sql &= " FROM EDTTRXN1, ICTITEM1"
            ASCMAIN1.sql &= " WHERE EDTTRXN1.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
            ASCMAIN1.sql &= " AND NVL(EDTTRXN1.PROCESS_IND, '0') = '0' AND EDTTRXN1.TRANS_TYPE = 'ADJ' AND REASON_CODE <> 'A05'"
            Create_TDA(.Tables.Add, "EDTTRXNX", "**", 0, False, String.Empty, 0)
            Create_TDA(.Tables.Add, "EDTTRXN1", "*")

            .Tables("EDTTRXNX").Columns.Add("REASON_DESC", GetType(String), "REASON_CODE")

        End With

        Set_Read_Only(grpTotals, True)

        Fill_Records("ICTREAS1")
        Fill_Records("ICTCLAS1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIADJ0.DataSource = dst.Tables("ICTIADJ0")
        grdICTIADJ2.DataSource = dst.Tables("ICTIADJ2")
        grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
        grdICTIADJX.DataSource = dst.Tables("ICTIADJX")
        grdICTIADJG.DataSource = dst.Tables("ICTIADJG")

        grdEDTTRXNX.DataSource = dst.Tables("EDTTRXNX")

        Create_Summary(grdEDTTRXNX, "SEL", "Sum")
        Create_Summary(grdEDTTRXNX, "TRANS_NUM", "Count")
        Create_Summary(grdEDTTRXNX, "TRAN_QTY", "Sum")

        Create_Summary(grdICTIADJX, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJX, "TOTAL_COSTS")

        Create_Summary(grdICTIADJG, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJG, "DIST_AMT")

        Create_Summary(grdICTIADJ2, "ADJ_LNO", "Count")
        Create_Summary(grdICTIADJ2, "ADJ_QTY")
        Create_Summary(grdICTIADJ2, "LINE_COSTS")

        Create_Summary(grdICTIADJ3, "ADJ_GNO", "Count")
        Create_Summary(grdICTIADJ3, "DIST_AMT")

        With grdICTIADJX.DisplayLayout.Bands("ICTIADJX")
            .Columns("ADJ_NO").Header.Fixed = True
        End With

        With grdICTIADJG.DisplayLayout.Bands("ICTIADJG")
            .Columns("ADJ_NO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdICTIADJX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        ASCMAIN1.Add_Value_List(grdICTIADJG, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        ASCMAIN1.Add_Value_List(grdEDTTRXNX, "REASON_DESC", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIADJ0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdICTIADJ3, "ICTIADJ3")

        Set_Read_Only(grpTotals, True)
        'If InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0 Then
        '    grpTotals.Visible = False
        '    With grdICTIADJ2.DisplayLayout.Bands(0)
        '        .Columns("ITEM_COST_STD").Hidden = True
        '        .Columns("LINE_COSTS").Hidden = True
        '        .Columns("COST_CATGY_CODE").Hidden = True
        '        .Columns("PROD_CODE").Hidden = True
        '    End With
        'End If

        grpHeader.Visible = False
        Set_SEGS(grdICTIADJG, "ICTIADJG")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("WHSE_CODE")

                Dim DT As Date = Absx1.dteFor("ADJ_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        Else
                            If rowICTWHSE1.Item("LP_CODE") & "" <> "" AndAlso Not processing3PL Then
                                If MsgBox("Warehouse Entered Is A 3PL." & vbCrLf & vbCrLf & "Do you want to Manually Adjust anyway?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                                Else
                                    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Adjustments Allowed"
                                End If
                            End If
                        End If
                    End If
                End If

                If EMsg.Length > 0 Then
                    processing3PL = False
                ElseIf processing3PL Then
                    ' need to multi task by the Edi Doc ref No
                    Dim rowEDTTRXN1 As DataRow = Nothing
                    For Each row As DataRow In dst.Tables("EDTTRXNX").Select("SEL = '1'")
                        Dim TRX_NO As String = row.Item("TRX_NO")
                        Dim TRX_LNO As Int16 = row.Item("TRX_LNO")

                        If Not ASCMAIN1.Logical_Lock("ICFIADJ1", TRX_NO & "-" & TRX_LNO) Then
                            Exit Sub
                        End If

                        ASCMAIN1.sql = "Select * from EDTTRXN1 WHERE TRX_NO = '" & TRX_NO & "' AND TRX_LNO = " & TRX_LNO
                        rowEDTTRXN1 = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                        If rowEDTTRXN1 Is Nothing OrElse Val(rowEDTTRXN1.Item("PROCESS_IND") & String.Empty) <> 0 Then
                            MessageBox.Show("Invalid or already processed adjustment.", "Process 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            ASCMAIN1.MultiTask_Release()
                            Exit Sub
                        End If
                    Next
                End If

            Case "View"
                If Absx1.txtFor("ADJ_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIADJ1 = LookUp("ICTIADJ1", Absx1.txtFor("ADJ_NO").Text)
                    If rowICTIADJ1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("ADJ_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                If Absx1.txtFor("REASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Reason"
                Else
                    Dim rowICTREAS1 As DataRow = LookUp("ICTREAS1", Absx1.txtFor("REASON_CODE").Text)
                    If rowICTREAS1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Reason"
                    End If
                End If

                If grdICTIADJ2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTIADJ2 As DataRow In dst.Tables("ICTIADJ2").Select("", "", DataViewRowState.CurrentRows)
                        If rowICTIADJ2.Item("COST_CATGY_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Cost Category for " & rowICTIADJ2.Item("ITEM_CODE") & ""
                        End If
                        If rowICTIADJ2.Item("PROD_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Product Code for " & rowICTIADJ2.Item("ITEM_CODE") & ""
                        End If
                    Next
                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTIADJ2")

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIADJ2", Absx1.txtFor("WHSE_CODE").Text, "ADJ_QTY", 1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK to Continue Anyway?",
                                  MsgBoxStyle.YesNo,
                                  "The following Items do not have Sufficent Qty for this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Reverse"
                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Import Clarins Adj"
                If InquiryMode OrElse ASCMAIN1.SOLUTION <> "INT" Then
                    Exit Sub
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reverse"
                Set_Up_Reversal()
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Import Clarins Adj"
                ImportClarinsAdj()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") And Not InquiryMode _
                        AndAlso rowICTIADJ1 IsNot Nothing _
                        AndAlso rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") Is DBNull.Value _
                        AndAlso rowICTIADJ1.Item("REVERSES_ADJ_NO") Is DBNull.Value

                    .Items("Import Clarins Adj").Settings.Enabled = not_iScreenMode
                    .Items("Import Clarins Adj").Visible = Not InquiryMode AndAlso ASCMAIN1.SOLUTION = "INT"
                End With

                .Groups("GL Distribution").Visible = False ' ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Show if Entered in").Visible = Not ScreenMode ' And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = False ' ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0
            Set_Read_Only(grpHeader, (EntryMode = "V") OrElse InquiryMode)
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIADJ2}
                    If processing3PL Then
                        With grd.DisplayLayout.Override
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.False
                        End With
                    Else
                        With grd.DisplayLayout.Override
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        End With
                    End If
                Next

                With grdICTIADJ2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.LightYellow
                End With

            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIADJ2, grdICTIADJ3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                With grdICTIADJ2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.Empty
                End With
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTIADJ0", "ICTIADJ1", "ICTIADJ2", "ICTIADJ3", "ICTSTAT2", "EDTTRXNX", "EDTTRXN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If chkGL.Checked Then
            chkGL.Checked = False
        Else
            Refresh_Documents()
        End If
        Setup_tab0_GL()

        'ICTPARM1.IC_PARM_WHSE_CODE
        Absx1.txtFor("WHSE_CODE").Text = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE") & String.Empty
        Absx1.dteFor("ADJ_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("ADJ_NO").Clear()
        grdICTIADJ2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop

        optGL.Tag = ""

        If processing3PL Then
            tab0.SelectedTab = tab0.Tabs("3PL")
        End If
        processing3PL = False

        If MENU_ITEM_OBJECT = "ICFIADJ1" Then
            grdEDTTRXNX.DisplayLayout.Bands(0).Columns("REASON_CODE").CellActivation = Activation.AllowEdit
            grdEDTTRXNX.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = Activation.AllowEdit
        End If
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTIADJ1 = dst.Tables("ICTIADJ1").NewRow
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("TRAN_NO_A")
            Else
                rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
            End If
            rowICTIADJ1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTIADJ1.Item("ADJ_DATE") = HFs("ADJ_DATE")
            rowICTIADJ1.Item("ADJ_SOURCE") = "E"
            rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("REGISTER_IND") = "0"
            rowICTIADJ1.Item("JOURNAL_IND") = "0"
            dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)
        Else
            Fill_Record("ICTIADJ1", Absx1.txtFor("ADJ_NO").Text)
            dst.AcceptChanges()

            dst.Tables("ICTIADJ0").Rows.Add(New String() {"Entered", Format(rowICTIADJ1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
            dst.Tables("ICTIADJ0").Rows.Add(New String() {"By", rowICTIADJ1.Item("INIT_OPER")})
            dst.Tables("ICTIADJ0").Rows.Add(New String() {"Source", rowICTIADJ1.Item("ADJ_SOURCE")})

            If rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") & "" <> "" Then
                Dim row As DataRow = LookUp("ICTIADJ1", rowICTIADJ1.Item("REVERSED_BY_ADJ_NO"))
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"using", rowICTIADJ1.Item("REVERSED_BY_ADJ_NO")})
            ElseIf rowICTIADJ1.Item("REVERSES_ADJ_NO") & "" <> "" Then
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"Reverses", rowICTIADJ1.Item("REVERSES_ADJ_NO")})
            End If
        End If


        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowICTIADJ1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTIADJ2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = True '  Not location_support
            .Columns("LOCATION_CODE").Hidden = Not location_support
        End With

        Fill_Records("ICTIADJ2", Absx1.txtFor("ADJ_NO").Text)
        Fill_Records("ICTIADJ3", Absx1.txtFor("ADJ_NO").Text)

        If EntryMode = "N" AndAlso processing3PL Then
            Dim sql As String = String.Empty
            dst.Tables("EDTTRXN1").Rows.Clear()
            For Each row As DataRow In dst.Tables("EDTTRXNX").Select("SEL = '1'")
                sql = "Select * from EDTTRXN1 Where TRX_NO = '" & row.Item("TRX_NO") & "' AND TRX_LNO = " & row.Item("TRX_LNO") & " AND TRANS_TYPE = 'ADJ'"
                Fill_Records("EDTTRXN1", "", False, sql)
            Next
            rowICTIADJ1.Item("REASON_CODE") = dst.Tables("EDTTRXNX").Compute("MAX(REASON_CODE)", "SEL = '1'") & String.Empty

            grdICTIADJ2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            For Each row As DataRow In dst.Tables("EDTTRXN1").Select("", "ITEM_CODE")
                grdICTIADJ2.DisplayLayout.Bands(0).AddNew()
                With grdICTIADJ2.ActiveRow
                    .Cells("ITEM_CODE").Value = row.Item("ITEM_CODE")
                    .Cells("ADJ_QTY").Value = row.Item("TRAN_QTY")
                    row.Item("PROCESS_IND") = "1"
                    .Update()
                End With
            Next
            grdICTIADJ2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdICTIADJ2, "ADJ_LNO")
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        If processing3PL Then
            Update_Record_TDA("EDTTRXN1")
        End If

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then
            Update_WHTLOCBX()
        End If
        CommitTrans("Update Complete")

    End Sub

    Sub Update_WHTLOCBX()
        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").Rows(0)
        TAC.ICCMAIN1.Update_WHTLOCBX("A", rowICTIADJ1.Item("ADJ_NO"))
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
        Load_Popup_Menu(grdICTIADJX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIADJ2, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdICTIADJG, "SS", "Show Filter", "Show GroupBox")

        If MENU_ITEM_OBJECT = "ICFIADJ1" Then
            Load_Popup_Menu(grdEDTTRXNX, "SSPBBPBBPBB", "Show Filter", "Show GroupBox", "Select All", "DeSelect All", "Process Selected", "Delete Selected", "Change Reason code to all selected", "Change Warehouse to all selected")
        Else
            Load_Popup_Menu(grdEDTTRXNX, "SSPBBPBB", "Show Filter", "Show GroupBox")
        End If
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

            Case "Change Reason code to all selected"
                Dim REASON_CODE As String = grd.ActiveRow.Cells("REASON_CODE").Value & String.Empty

                Dim numSelecteRows As Int16 = 0
                Dim rows As Infragistics.Win.UltraWinGrid.UltraGridRow() = grd.Rows.GetFilteredInNonGroupByRows()
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In rows
                    If row.Cells("SEL").Value & String.Empty = "1" Then
                        numSelecteRows += 1
                    End If
                Next

                If numSelecteRows = 0 Then
                    MessageBox.Show("At least one grid entry must be selected.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim zMsg As String = $"Do you want to change the Reason Code to '{REASON_CODE}' for the {numSelecteRows} selected rows?"

                If MessageBox.Show(zMsg, e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In rows
                    If row.Cells("SEL").Value & String.Empty = "1" Then
                        Dim TRX_NO As String = row.Cells("TRX_NO").Value & String.Empty
                        Dim TRX_LNO As Int16 = row.Cells("TRX_LNO").Value & String.Empty

                        Try
                            BeginTrans()
                            Dim dr As DataRow = dst.Tables("EDTTRXNX").Select($"TRX_NO = '{TRX_NO}' and TRX_LNO = {TRX_LNO}")(0)

                            Dim Sql As String = "UPDATE EDTTRXN1 SET REASON_CODE = :PARM1 WHERE TRX_NO = :PARM2 AND TRX_LNO = :PARM3"
                            ASCDATA1.ExecuteSQL(Sql, "VVN", {REASON_CODE, TRX_NO, TRX_LNO})
                            CommitTrans()
                            dr.Item("REASON_CODE") = REASON_CODE
                        Catch ex As Exception
                            Rollback()
                            MessageBox.Show(ex.Message, e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End Try

                    End If
                Next

                grd.Update()

            Case "Change Warehouse to all selected"
                Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value & String.Empty
                Dim numSelecteRows As Int16 = 0
                Dim rows As Infragistics.Win.UltraWinGrid.UltraGridRow() = grd.Rows.GetFilteredInNonGroupByRows()
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In rows
                    If row.Cells("SEL").Value & String.Empty = "1" Then
                        numSelecteRows += 1
                    End If
                Next

                If numSelecteRows = 0 Then
                    MessageBox.Show("At least one grid entry must be selected.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim zMsg As String = $"Do you want to change the Warehouse Code to '{WHSE_CODE}' for the {numSelecteRows} selected rows?"

                If MessageBox.Show(zMsg, e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In rows
                    If row.Cells("SEL").Value & String.Empty = "1" Then
                        Dim TRX_NO As String = row.Cells("TRX_NO").Value & String.Empty
                        Dim TRX_LNO As Int16 = row.Cells("TRX_LNO").Value & String.Empty

                        Try
                            BeginTrans()
                            Dim dr As DataRow = dst.Tables("EDTTRXNX").Select($"TRX_NO = '{TRX_NO}' and TRX_LNO = {TRX_LNO}")(0)

                            Dim Sql As String = "UPDATE EDTTRXN1 SET WHSE_CODE = :PARM1 WHERE TRX_NO = :PARM2 AND TRX_LNO = :PARM3"
                            ASCDATA1.ExecuteSQL(Sql, "VVN", {WHSE_CODE, TRX_NO, TRX_LNO})
                            CommitTrans()
                            dr.Item("WHSE_CODE") = WHSE_CODE
                        Catch ex As Exception
                            Rollback()
                            MessageBox.Show(ex.Message, e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End Try

                    End If
                Next

                grd.Update()

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Select All"
                For Each row As DataRow In dst.Tables("EDTTRXNX").Select()
                    row.Item("SEL") = "1"
                Next

            Case "DeSelect All"
                For Each row As DataRow In dst.Tables("EDTTRXNX").Select()
                    row.Item("SEL") = "0"
                Next

            Case "Process Selected"
                Load3PLAdjustments()

            Case "Delete Selected"
                grdEDTTRXNX.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                If dst.Tables("EDTTRXNX").Select("SEL = '1'").Length = 0 Then
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to delete the " & dst.Tables("EDTTRXNX").Select("SEL = '1'").Length & " selected Adjustments?",
                                     "Delete Selected", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ' need to multi task by the Edi Doc ref No
                Dim TRX_NO As String = String.Empty
                Dim TRX_LNO As Int16 = 0

                For Each row As DataRow In dst.Tables("EDTTRXNX").Select("SEL = '1'")
                    TRX_NO = row.Item("TRX_NO")
                    TRX_LNO = row.Item("TRX_LNO")
                    If Not ASCMAIN1.Logical_Lock("ICFIADJ1", TRX_NO & "-" & TRX_LNO) Then
                        Continue For
                    End If

                    ASCMAIN1.sql = "Update EDTTRXN1 set PROCESS_IND = '1' WHERE TRX_NO = '" & TRX_NO & "' AND TRX_LNO = " & TRX_LNO
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    row.Delete()
                Next
                ASCMAIN1.MultiTask_Release()

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "ADJ_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not InquiryMode Then
                    'Click_Command("New")
                End If
            Case "ADJ_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTIADJ2"

    Private Sub grdICTIADJ2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIADJ2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTIADJ2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("PROD_CODE").Value = cdr.Item("PROD_CODE")

                    'Dim rowICTSTAT2 = Fill_Record("ICTSTAT2", New String() {ITEM_CODE, WHSE_CODE}, True)
                    Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    Dim ITEM_COST_STD As Decimal = Val(cdr.Item("ITEM_COST_STD") & "")
                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD
                Else
                    grdICTIADJ2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "ADJ_QTY"

        End Select
    End Sub

    Private Sub grdICTIADJ2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterExitEditMode

        'Select Case grdICTIADJ2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdICTIADJ2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterRowActivate
        With grdICTIADJ2.DisplayLayout.Bands(0)
            If grdICTIADJ2.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTIADJ2.ActiveCell = grdICTIADJ2.ActiveRow.Cells("ITEM_CODE")
                grdICTIADJ2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdICTIADJ2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIADJ2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIADJ2.AfterRowUpdate
        DisplayTotals()
    End Sub


    Private Sub grdICTIADJ2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIADJ2.BeforeExitEditMode
        If grdICTIADJ2.ActiveCell Is Nothing Then Exit Sub
        With grdICTIADJ2.ActiveCell
            Select Case .Column.Key
                Case "ITEM_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Item Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        Else
                            e.Cancel = True
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdICTIADJ2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIADJ2.BeforeRowUpdate
        With grdICTIADJ2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                If e.Row.Cells("LOCATION_CODE").Text = "" Then
                    MsgBox("Entry of a Location Code is Mandatory",
                                                MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")",
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If

            End If

            If Val(e.Row.Cells("ADJ_QTY").Text) = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("ADJ_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("ADJ_NO").Text = "" Then
                    .ActiveRow.Cells("ADJ_NO").Value = Absx1.CtlFor("ADJ_NO").Text
                    .ActiveRow.Cells("ADJ_LNO").Value = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTIADJ2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIADJ2.ClickCellButton

        If grdICTIADJ2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTIADJ2, sql_where, False)

    End Sub

    Private Sub grdICTIADJ2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIADJ2.Error
        grdICTIADJ2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIADJ2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTIADJX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIADJX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ADJ_NO").Text = e.Row.Cells("ADJ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTIADJG_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIADJG.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ADJ_NO").Text = e.Row.Cells("ADJ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
                Dim dvw As DataView = dst.Tables("ICTIADJ3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
                Dim dvw As DataView = dst.Tables("ICTIADJ3").DefaultView
                Dim ADJ_LNO As Integer = 0
                If grdICTIADJ2.ActiveRow IsNot Nothing Then
                    ADJ_LNO = Val(grdICTIADJ2.ActiveRow.Cells("ADJ_LNO").Text)
                End If
                dvw.RowFilter = "ADJ_LNO = " & CStr(ADJ_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIADJ3").Clone
                Dim ADJ_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIADJ3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("ICTIADJ3").Compute _
                    ("SUM(DIST_AMT)",
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("ADJ_NO") = Absx1.txtFor("ADJ_NO").Text
                    row.Item("ADJ_LNO") = 0
                    ADJ_GNO += 1
                    row.Item("ADJ_GNO") = ADJ_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIADJ3.DataSource = tbl
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value

        Fill_Records("ICTIADJX", YP)
        Sort_grdColumns(grdICTIADJX, "ADJ_NO".ToLower)
        grdICTIADJX.Text = "Entered in " & cbeYP.Text

        If chkGL.Checked Then
            Fill_Records("ICTIADJG", YP)
            grdICTIADJG.Text = "Entered in " & cbeYP.Text
        End If

        Fill_Records("EDTTRXNX")
        grdEDTTRXNX.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String,
                       ByVal WHSE_CODE As String,
                       ByVal QTY_FIELD As String,
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from ICTSTAT2 where ITEM_CODE = '" & ITEM_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            Dim WHSE_QTY_ON_HAND As Integer = 0
            If rowICTSTAT2 IsNot Nothing Then
                WHSE_QTY_ON_HAND = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
            If WHSE_QTY_ON_HAND + S * QTY < 0 Then
                msg &= vbCr & Format("Style/Color " & ITEM_CODE & " has only " & CStr(WHSE_QTY_ON_HAND) & " On Hand")
            End If
        Next

        Return msg
    End Function

    Private Sub chkGL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGL.CheckedChanged
        Setup_tab0_GL()
    End Sub

    Sub Setup_tab0_GL()
        If Not chkGL.Checked Then
            tab0.Tabs(0).Selected = True
        Else
            Refresh_Documents()
        End If
        tab0.Tabs("GL").Visible = chkGL.Checked

        If chkGL.Checked Then
            tab0.Tabs("GL").Selected = True
        End If
    End Sub

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_ADJ_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            REVERSED_BY_ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            REVERSED_BY_ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        Dim rowICTIADJ1_orig As DataRow = dst.Tables("ICTIADJ1").NewRow
        rowICTIADJ1_orig.ItemArray = rowICTIADJ1.ItemArray

        rowICTIADJ1 = dst.Tables("ICTIADJ1").Rows(0)
        rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") = REVERSED_BY_ADJ_NO
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIADJ1")

        rowICTIADJ1.ItemArray = rowICTIADJ1_orig.ItemArray
        rowICTIADJ1.AcceptChanges()
        rowICTIADJ1.SetAdded()

        With rowICTIADJ1
            .Item("REVERSES_ADJ_NO") = .Item("ADJ_NO")
            .Item("ADJ_NO") = REVERSED_BY_ADJ_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("ADJ_DATE") = DATETIME_STAMP.Date
            .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
            .Item("JOURNAL_IND") = "0"
            .Item("JOURNAL_XNO") = DBNull.Value
        End With

        'Set new RTRN_NO and reverse all quantities for this return.
        For Each row As DataRow In dst.Tables("ICTIADJ2").Rows
            row.Item("ADJ_NO") = REVERSED_BY_ADJ_NO
            If row.Item("ADJ_QTY") IsNot DBNull.Value Then
                row.Item("ADJ_QTY") *= -1
            End If
            row.Item("OPS_YYYYPP") = ASCMAIN1.CYP

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Private Sub grdEDTTRXNX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDTTRXNX.DoubleClickRow
        Load3PLAdjustments()
    End Sub

    Private Sub Load3PLAdjustments()

        Try
            grdEDTTRXNX.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

            If dst.Tables("EDTTRXNX").Select("SEL = '1'").Length = 0 Then
                MessageBox.Show("There are no 3PL adjustments selected", "Process Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            ' See if all adjustments are for the same warehouse
            If ASCDATA1.SelectDistinct(dst.Tables("EDTTRXNX").Select("SEL = '1'"), New String() {"WHSE_CODE"}).Rows.Count > 1 Then
                MessageBox.Show("You selected multiple warehouses. You cannot continue?", "Process Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            ' Do not permit selection of Adjustment for Re-Work
            If dst.Tables("EDTTRXNX").Select("SEL = '1' AND (REASON_CODE = 'PRODUCTION' OR REASON_CODE = 'A24')").Length > 0 Then
                MessageBox.Show("You selected at least 1 adjustment with reason code PRODUCTION or A24 - not permitted", "Process Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If ASCDATA1.SelectDistinct(dst.Tables("EDTTRXNX").Select("SEL = '1'"), New String() {"REASON_CODE"}).Rows.Count > 1 Then
                If MessageBox.Show("You selected multiple reason codes. Do you want to continue?", "Process Selected", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
            End If

            'If MyBase.Absx1.txtFor("WHSE_CODE").TextLength = 0 Then
            '    If ASCDATA1.SelectDistinct(dst.Tables("EDTTRXNX").Select("SEL = '1'"), New String() {"WHSE_CODE"}).Rows.Count > 1 Then
            '        MessageBox.Show("You selected multiple Warehouse codes?", "Process Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        Exit Sub
            '    End If

            '    MyBase.Absx1.txtFor("WHSE_CODE").Text = dst.Tables("EDTTRXNX").Select("SEL = '1'")(0).Item("WHSE_CODE") & String.Empty
            'End If

            MyBase.Absx1.txtFor("WHSE_CODE").Text = dst.Tables("EDTTRXNX").Select("SEL = '1'")(0).Item("WHSE_CODE") & String.Empty
            processing3PL = True
            Click_Command("New")

        Catch ex As Exception
            MessageBox.Show("Load 3PL Adjustments Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ImportClarinsAdj()

        Try

            If InquiryMode OrElse ASCMAIN1.SOLUTION <> "INT" Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            Dim messages As New List(Of String)
            Dim success As Boolean = False
            ' Change for ADS 07/16/2025, force developer to supply the LP CODE
            Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC945O1", "ADJ", "ADJ", messages, success, "CLA")

            Dim userMessage As String = String.Empty

            For Each msg As String In messages
                userMessage &= vbCr & msg
            Next

            If userMessage.Length > 0 Then
                MessageBox.Show(userMessage, "Import", MessageBoxButtons.OK)
            End If

            Fill_Records("EDTTRXNX")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub grdEDTTRXNX_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdEDTTRXNX.ClickCellButton

        If grdEDTTRXNX.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "WHSE_CODE"
                sql_where = "LP_CODE IS NOT NULL"

            Case "REASON_CODE"

            Case Else
                Exit Sub

        End Select

        grdClickCellButton(grdEDTTRXNX, sql_where, True)

    End Sub

    Private Sub grdEDTTRXNX_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdEDTTRXNX.AfterRowUpdate

        Try
            BeginTrans()
            Dim TRX_NO As String = e.Row.Cells("TRX_NO").Value & String.Empty ' dr.Item("TRX_NO")
            Dim TRX_LNO As Int16 = e.Row.Cells("TRX_LNO").Value & String.Empty
            Dim WHSE_CODE As String = e.Row.Cells("WHSE_CODE").Value & String.Empty
            Dim REASON_CODE As String = e.Row.Cells("REASON_CODE").Value & String.Empty

            Dim Sql As String = "UPDATE EDTTRXN1 SET WHSE_CODE = :PARM1, REASON_CODE = :PARM2 WHERE TRX_NO = :PARM3 AND TRX_LNO = :PARM4"
            ASCDATA1.ExecuteSQL(Sql, "VVVN", {WHSE_CODE, REASON_CODE, TRX_NO, TRX_LNO})
            CommitTrans()
        Catch ex As Exception
            Rollback()
            MessageBox.Show(ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class