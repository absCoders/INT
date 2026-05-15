Imports Infragistics.Win.UltraWinGrid

Public Class SAFSTKR1

    Dim sqlSATSTKRX As String = ""
    Dim sqlSATSTKRS As String = ""
    Dim SQLWEEKS As String = ""

    Dim SATSTKRX As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE, QTY_COUNTED QTY_OPEN, QTY_COUNTED QTY_PICK, QTY_COUNTED QTY_SHIP, QTY_COUNTED QTY_ONHD from ICTSREQ1 where ROWNUM < 1"
            SATSTKRX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SATSTKRX} Add Primary Key (STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE)")

            sqlSATSTKRX = $"Select ICTSREQ1.*, ICTITEM1.PROD_CODE, ICTITEM1.COLLECTION_CODE
                , ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.HC_CODE, GLTPARM3.LEGEND
                from ICTSREQ1, ICTITEM1, ICTCOLL1, GLTPARM3
                where ICTITEM1.ITEM_CODE = ICTSREQ1.ITEM_CODE
                and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE
                and GLTPARM3.YYYYWW = ICTSREQ1.OPS_YYYYWW"

            ASCMAIN1.sql = Replace(sqlSATSTKRX, "from ICTSREQ1",
                                   $"from ICTSREQ1, {SATSTKRX} SATSTKRX") _
                                   & " and ICTSREQ1.STOCK_REQ_CTL_NO = SATSTKRX.STOCK_REQ_CTL_NO
                                       and ICTSREQ1.SELL_CODE = SATSTKRX.SELL_CODE 
                                       and ICTSREQ1.CUST_CODE = SATSTKRX.CUST_CODE 
                                       and ICTSREQ1.CUST_STORE_NO = SATSTKRX.CUST_STORE_NO
                                       and ICTSREQ1.ITEM_CODE = SATSTKRX.ITEM_CODE"

            ASCMAIN1.sql = Replace(ASCMAIN1.sql,
                "ICTSREQ1.*",
                "ICTSREQ1.*, SATSTKRX.QTY_OPEN, SATSTKRX.QTY_PICK, SATSTKRX.QTY_SHIP, SATSTKRX.QTY_ONHD")

            Create_TDA(.Tables.Add, "SATSTKRX", "**", 0, False, "")
            With dst.Tables("SATSTKRX")
                ' .Columns.Add("LEGEND", GetType(System.Int32))
                '    .Columns.Add("QTY_OPEN", GetType(System.Int32))
                '    .Columns.Add("QTY_PICK", GetType(System.Int32))
                '    .Columns.Add("QTY_SHIP", GetType(System.Int32))
                '    .Columns.Add("QTY_ONHD", GetType(System.Int32))
            End With

            Dim sqlICTITEMX As String = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.PROD_CODE, ICTITEM1.COLLECTION_CODE
                , ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.HC_CODE
                from ICTITEM1, ICTCOLL1 where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"

                ASCMAIN1.sql = sqlICTITEMX & " and ICTITEM1.ITEM_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "V", 1)

                ASCMAIN1.sql = sqlICTITEMX & " and ICTITEM1.ITEM_STK_REQ = '1'"
                Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False,, 1)

                ASCMAIN1.sql = "Select ICTSREQ1.CUST_CODE
                , Case when ICTSREQ1.CUST_CODE = 'MACYS' then ICTCOLL1.COLLECTION_GENDER ELSE 'A' END COLLECTION_GENDER
                , ARTCUST1.SREP_CODE
                , COUNT (*) RECORDS
                from ICTSREQ1,ICTITEM1,ICTCOLL1,ARTCUST1
                where ICTITEM1.ITEM_CODE = ICTSREQ1.ITEM_CODE
                and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE
                and ARTCUST1.CUST_CODE = ICTSREQ1.CUST_CODE
                group by ICTSREQ1.CUST_CODE
                , Case when ICTSREQ1.CUST_CODE = 'MACYS' then ICTCOLL1.COLLECTION_GENDER ELSE 'A' END
                , ARTCUST1.SREP_CODE"
                sqlSATSTKRS = ASCMAIN1.sql
                Create_TDA(.Tables.Add, "SATSTKRS", "**", 0, False,, 2)
                With dst.Tables("SATSTKRS")
                    .Columns.Add("SELECTED")
                    .Columns("SELECTED").DefaultValue = "0"
                End With
            End With


            Set_cmbYW("RYW0", ASCMAIN1.CYW, -5 * 52, 1 * 52, -13) '  -13 - 1 * 52)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -5 * 52, 1 * 52, 0) ' 0 - 1 * 52)


        grdSATSTKRX.DataSource = dst.Tables("SATSTKRX")
        grdSATSTKRS.DataSource = dst.Tables("SATSTKRS")
        grdICTITEMX.DataSource = dst.Tables("ICTITEMX")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSTKRX, grdSATSTKRS, grdICTITEMX}
            grd.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.LightGreen
            grd.DisplayLayout.Override.ActiveRowAppearance.ForeColor = System.Drawing.Color.Black

            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            If grd.Name = "grdICTITEMX" Then
                'grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                'grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If

            If grd.Name = "grdSATSTKRS" Then
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    c.CellActivation = Activation.NoEdit
                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If grd.Name = "grdSATSTKRS" Then
                        If c.Key = "SELECTED" Then
                            c.CellActivation = Activation.AllowEdit
                        End If
                        If c.Key = "RECORDS" Then
                            c.Format = "#,##0"
                        End If
                    End If
                Next
            End With
        Next

        spl.Panel1Collapsed = True


        ' ASCMAIN1.Add_Value_List(grdSATSTKRS, "COLLECTION_GENDER")
        ASCMAIN1.Add_Value_List(grdSATSTKRS, "COLLECTION_GENDER", Nothing, New String() {":", "M:Mens", "W:Womens", "A:All"})

        Create_Summary(grdSATSTKRX, "SELL_CODE", "Count")
        Create_Summary(grdICTITEMX, "ITEM_CODE", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If dst.Tables("SATSTKRS").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Customers Selected"
                End If

                If EMsg = "" Then
                    If optNA.Value = "N" Then
                        For Each row As DataRow In ASCDATA1.SelectDistinct("SATSTKRS", New String() {"CUST_CODE", "COLLECTION_GENDER"}).Select
                            Dim CUST_CODE As String = row.Item("CUST_CODE")
                            Dim COLLECTION_GENDER As String = row.Item("COLLECTION_GENDER")
                            If Not ASCMAIN1.Logical_Lock("SATSTKRS", CUST_CODE & ":" & COLLECTION_GENDER) Then
                                Exit Sub
                            End If
                        Next
                    End If
                End If

            Case "Update"

                'If Absx1.txtFor("OPS_YYYYPP").Text & "" < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                '    EMsg &= vbCr & "You cannot Update a Closed GL Period"
                'End If

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
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Update").Visible = ScreenMode And (optNA.Value = "N")
                    .Items("Cancel").Visible = ScreenMode And (optNA.Value = "N")
                    .Items("Done").Visible = ScreenMode And (optNA.Value = "A")
                End With
            End With
        End If

        tabSATSTKRX.Visible = Not (ScreenMode)
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSATSTKRX.Visible = (ScreenMode)

        If ScreenMode Then
            If optNA.Value = "N" Then
                grdSATSTKRX.Text = "Stock Request Entries (Click Update to Clear All New Entries)"

            Else
                grdSATSTKRX.Text = $"Stock Request Entries (Archived from {Absx1.cmbFor("RYW0").Text} to {Absx1.cmbFor("RYW0").Text})"
            End If

            With grdSATSTKRX.DisplayLayout.Bands(0)
                .Columns("LAST_OPER").Hidden = (optNA.Value = "N")
                .Columns("LAST_DATE").Hidden = (optNA.Value = "N")
            End With
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATSTKRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("SATSTKRX")
        Sort_grdColumns(grdSATSTKRX, "INIT_DATE".ToLower & ", SELL_CODE, CUST_CODE, CUST_STORE_NO")

        Fill_Records("ICTITEMX")
        Sort_grdColumns(grdICTITEMX, "BRAND_CODE, COLLECTION_CODE, PROD_CODE, ITEM_CODE")

        Get_RecordCount()

        Setup_tabSATSTKRX()
        ' cmdSaveItemList.Tag = ""
        cmdSaveItemList.Visible = False
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        dst.Tables("SATSTKRX").Rows.Clear()

        ASCDATA1.ExecuteSQL($"Truncate Table {SATSTKRX}")

        Dim CYW = ASCMAIN1.CYW
        'CYW = "202233" : Stop
        Dim YW3AGO As String = ASCMAIN1.Week_Calc(CYW, -3)
        Dim YW1AGO As String = ASCMAIN1.Week_Calc(CYW, -1)

        For Each row As DataRow In dst.Tables("SATSTKRS").Select("SELECTED = '1'")

            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim COLLECTION_GENDER As String = row.Item("COLLECTION_GENDER")

            Dim SQLW As String = ""
            If optNA.Value = "N" Then
                SQLW = " and REPORTED = 'N'"
            Else
                SQLW = $" and REPORTED = 'A' {SQLWEEKS}"
            End If

            SQLW &= $" and ICTSREQ1.CUST_CODE = '{CUST_CODE}'"
            If COLLECTION_GENDER <> "A" Then
                SQLW &= $" and ICTCOLL1.COLLECTION_GENDER = '{COLLECTION_GENDER}'"
            End If

            ASCMAIN1.sql = $"Insert into {SATSTKRX} (STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE) (Select STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE from ({sqlSATSTKRX & SQLW}))"
            ASCDATA1.ExecuteSQL()

            'Fill_Records("SATSTKRX",, False, ASCMAIN1.sql)
        Next

        'ASCMAIN1.sql = $"
        '    Begin
        '     Declare Cursor C1 is 
        '        Select ITEM_CODE, SUM (WHSE_QTY_OPEN) QTY_OPEN, SUM (WHSE_QTY_PICK) QTY_PICK
        '            from ICTSTAT2 where ITEM_CODE in (Select Distinct ITEM_CODE from {SATSTKRX}) group by ITEM_CODE;
        '     Begin
        '      For R1 in C1 Loop
        '       Update {SATSTKRX} Set QTY_OPEN = R1.QTY_OPEN, QTY_PICK = R1.QTY_PICK where ITEM_CODE = R1.ITEM_CODE;
        '      End Loop;
        '     End;
        '    End;"
        'ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = $"
            Begin
             Declare Cursor C1 is 
                Select CUST_CODE, CUST_STORE_NO, ITEM_CODE, SUM (ORDR_QTY_OPEN) QTY_OPEN, SUM (ORDR_QTY_PICK) QTY_PICK
                    from SOTORDR2 
                    where (CUST_CODE, CUST_STORE_NO, ITEM_CODE) in 
                        (Select Distinct CUST_CODE, CUST_STORE_NO, ITEM_CODE from {SATSTKRX}) 
                    and ORDR_STATUS >= 'O' and ORDR_STATUS >= 'P'
                    group by CUST_CODE, CUST_STORE_NO, ITEM_CODE;
             Begin
              For R1 in C1 Loop
               Update {SATSTKRX} Set QTY_OPEN = R1.QTY_OPEN, QTY_PICK = R1.QTY_PICK 
                where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO and ITEM_CODE = R1.ITEM_CODE;
              End Loop;
             End;
            End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"
            Begin
             Declare Cursor C1 is 
                Select CUST_CODE, CUST_STORE_NO, ITEM_CODE, SUM (ORDR_QTY_SHIP) QTY_SHIP
                    from SOTINVH2 
                    where (CUST_CODE, CUST_STORE_NO, ITEM_CODE) in 
                        (Select Distinct CUST_CODE, CUST_STORE_NO, ITEM_CODE from {SATSTKRX}) 
                    and OPS_YYYYWW >= '{YW3AGO}'
                    group by CUST_CODE, CUST_STORE_NO, ITEM_CODE;
             Begin
              For R1 in C1 Loop
               Update {SATSTKRX} Set QTY_SHIP = R1.QTY_SHIP 
                where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO and ITEM_CODE = R1.ITEM_CODE;
              End Loop;
             End;
            End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"
            Begin
             Declare Cursor C1 is 
                Select CUST_CODE, CUST_STORE_NO, ITEM_CODE, SUM (QTY_EOW) QTY_ONHD
                    from RSTRETL1 
                    where (CUST_CODE, CUST_STORE_NO, ITEM_CODE) in 
                        (Select Distinct CUST_CODE, CUST_STORE_NO, ITEM_CODE from {SATSTKRX}) 
                    and OPS_YYYYWW = '{YW1AGO}'
                    group by CUST_CODE, CUST_STORE_NO, ITEM_CODE;
             Begin
              For R1 in C1 Loop
               Update {SATSTKRX} Set QTY_ONHD = R1.QTY_ONHD 
                where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO and ITEM_CODE = R1.ITEM_CODE;
              End Loop;
             End;
            End;"
        ASCDATA1.ExecuteSQL()

        Fill_Records("SATSTKRX")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        ASCMAIN1.sql = $"Update ICTSREQ1 Set REPORTED = 'A', LAST_OPER = '{ASCMAIN1.USER_ID}', LAST_DATE = SYSDATE
            where (STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE) in
            (Select STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE from {SATSTKRX})"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
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
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Me.Find_ITEM_CODE()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Find_ITEM_CODE()
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

    Public Overrides Sub cmb_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "RYW0"
                Set_Weeks()
            Case "RYW1"
                Set_Weeks()
        End Select
    End Sub

#End Region

    Sub Set_Weeks()
        Get_RecordCount()

        If Absx1.cmbFor("RYW0").Value & "" <> "" And Absx1.cmbFor("RYW1").Value & "" <> "" Then
            Dim C As Integer = 1 + ASCMAIN1.Week_Diff(Absx1.cmbFor("RYW0").Value, Absx1.cmbFor("RYW1").Value)
            lblWeeks.Text = CStr(C) & " Wks"
        Else
            lblWeeks.Text = ""
        End If
    End Sub

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATSTKRX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTITEMX, "SSB", "Show Filter", "Show GroupBox", "Load from XLS")
        Load_Popup_Menu(grdSATSTKRS, "SSB", "Show Filter", "Show GroupBox", "Select All", "De-Select All")
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

            Case "Load from XLS"
                Load_from_XLS()


            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next
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


    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            Case "ITEM_CODE"
                If chkActiveOnly.Checked Then sql_where &= " and ITEM_STATUS = 'A'"
                If optSN.Value <> "*" Then sql_where &= $" and ITEM_SNU_CODE = '{ optSN.Value}'"
                If optBP.Value <> "*" Then sql_where &= $" and ITEM_BASIC_PROMO = '{ optBP.Value}'"
        End Select
    End Sub

    Sub Find_ITEM_CODE()
        Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
        If ITEM_CODE <> "" Then
            Dim rowICTITEMX As DataRow = Add_Item(ITEM_CODE)
        End If

        Absx1.txtFor("ITEM_CODE").Text = ""
    End Sub


    Sub Load_from_XLS()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim ITEM_CODEs As New Dictionary(Of Integer, Integer)
            Dim c As Integer = 1
            Do While oSheet.Cells(c, 0).Value & "" <> ""
                Dim ITEM_CODE As String = oSheet.Cells(c, 0).Value & ""
                Add_Item(ITEM_CODE)

                c += 1
            Loop

            Sort_grdColumns(grdICTITEMX, "BRAND_CODE, COLLECTION_CODE, ITEM_CODE")

            MsgBox("XLS has been Loaded", MsgBoxStyle.OkOnly, "Success")
        End If

    End Sub

    Function Add_Item(ITEM_CODE As String) As DataRow
        Dim rowICTITEM1 As DataRow = Fill_Record("ICTITEM1", ITEM_CODE)

        Dim rowICTITEMX As DataRow = Nothing
        If rowICTITEM1 IsNot Nothing Then
            rowICTITEMX = dst.Tables("ICTITEMX").Rows.Find(ITEM_CODE)
            If rowICTITEMX Is Nothing Then
                rowICTITEMX = dst.Tables("ICTITEMX").NewRow
                rowICTITEMX.ItemArray = rowICTITEM1.ItemArray
                dst.Tables("ICTITEMX").Rows.Add(rowICTITEMX)
            End If
        End If

        If rowICTITEMX IsNot Nothing Then cmdSaveItemList.Visible = True

        Return rowICTITEMX

    End Function

    Private Sub tabSATSTKRX_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSATSTKRX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabSATSTKRX()
    End Sub

    Sub Setup_tabSATSTKRX()
        UltraExplorerBar1.Groups("Screen Control").Visible = Not (tabSATSTKRX.SelectedTab.Key = "Items")
        UltraExplorerBar1.Groups("Items").Visible = (tabSATSTKRX.SelectedTab.Key = "Items")
    End Sub

    Private Sub cmdSaveItemList_Click(sender As Object, e As EventArgs) Handles cmdSaveItemList.Click

        BeginTrans()

        ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_STK_REQ  = '0'"
        ASCDATA1.ExecuteSQL()

        For Each row As DataRow In dst.Tables("ICTITEMX").Select()
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_STK_REQ  = '1' where ITEM_CODE = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", ITEM_CODE)
        Next

        CommitTrans("List Saved")

        cmdSaveItemList.Visible = False
    End Sub

    Private Sub grdICTITEMX_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdICTITEMX.AfterRowsDeleted
        cmdSaveItemList.Visible = True
    End Sub
    Sub Get_RecordCount()

        SQLWEEKS = $" and OPS_YYYYWW between '{Absx1.cmbFor("RYW0").Value}' and '{Absx1.cmbFor("RYW1").Value}'"

        ASCMAIN1.sql = sqlSATSTKRS
        If optNA.Value = "N" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by", " and REPORTED = 'N' group by")
        If optNA.Value = "A" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by", $" and REPORTED = 'A' {SQLWEEKS} group by")
        Fill_Records("SATSTKRS", , , ASCMAIN1.sql)
        Sort_grdColumns(grdSATSTKRS, "CUST_CODE, COLLECTION_GENDER")
    End Sub

    Private Sub optNA_ValueChanged(sender As Object, e As EventArgs) Handles optNA.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        grpWEEK_RANGE.Visible = (optNA.Value = "A")
        Get_RecordCount()
    End Sub
End Class