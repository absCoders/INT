Imports Infragistics.Win.UltraWinGrid

Public Class SAFCSTK1

    Dim SELL_CODE As String
    Dim CSREQ_NO As String
    Dim rowSATCSTK1 As DataRow


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select SATCSTK1.* from SATCSTK1"
            Create_TDA(.Tables.Add, "SATCSTKX", "**", 1, False)

            ASCMAIN1.sql = "Select SATCSTK1.* from SATCSTK1 where SATCSTK1.CSREQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "SATCSTK1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select SATCSTK2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_SO_QTY_MULT
                , ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE
                , SOTALLO1.DATE_START, SOTALLO1.DATE_END, SOTALLO3.QTY_ALLO
                from SATCSTK2, ICTITEM1, ICTCOLL1, SOTALLO1, SOTALLO3
                where ICTITEM1.ITEM_CODE = SATCSTK2.ITEM_CODE 
                and SATCSTK2.CSREQ_NO = :PARM1
                and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE 
                and SOTALLO1.ALLO_CTL_NO (+) = SATCSTK2.ALLO_CTL_NO 
                and SOTALLO3.ALLO_CTL_NO (+) = SATCSTK2.ALLO_CTL_NO 
                and SOTALLO3.CUST_CODE (+) = :PARM2
                and SOTALLO3.CUST_STORE_NO (+) = :PARM3"
            Create_TDA(.Tables.Add, "SATCSTK2", "**", 0, True, "VVV", 2)
            With dst.Tables("SATCSTK2")
                ' .Columns.Add("LEGEND", GetType(System.Int32))
            End With


            ASCMAIN1.sql = "SELECT X.*, SOTALLO3.QTY_ALLO FROM SOTALLO3,(
                Select SOTALLO1.ALLO_CTL_NO, SOTALLO1.DATE_START, SOTALLO1.DATE_END, ICTCOLL1.BRAND_CODE,
                 ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_CODE, ICTCOLL1.COLLECTION_GENDER
                 from ICTITEM1, SOTALLO1, ICTCOLL1, ICTBRAN1
                 WHERE ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE
                 AND NVL(ICTITEM1.ITEM_STATUS, 'A') = 'A'
                 AND NVL(ICTCOLL1.COLLECTION_STATUS, 'A') = 'A'
                 AND NVL(ICTBRAN1.BRAND_STATUS, 'A') = 'A'
                 AND (ICTITEM1.ITEM_SNU_CODE = 'N' or ICTITEM1.PROD_CODE = 'DS')
                 AND ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE
                 AND ICTCOLL1.BRAND_CODE = ICTBRAN1.BRAND_CODE
                 AND ((SOTALLO1.DATE_START <= :PARM1 AND SOTALLO1.DATE_END >= :PARM2)
                 or (SOTALLO1.DATE_START <= :PARM3 AND SOTALLO1.DATE_END >= :PARM4))
                ) X WHERE SOTALLO3.ALLO_CTL_NO = X.ALLO_CTL_NO AND SOTALLO3.CUST_CODE = :PARM5 AND SOTALLO3.CUST_STORE_NO = :PARM6"
            Create_TDA(.Tables.Add, "SATCSTKA", "**", 0, False, "VVVVVV", 0)

            ASCMAIN1.sql = "Select 'A' ADDR_TYPE, CUST_STORE_NO KEY_CODE, 'STORE' KEY_TYPE
            , CUST_STORE_NAME, CUST_STORE_ADDR1, CUST_STORE_ADDR2, CUST_STORE_ADDR3
            , CUST_STORE_CITY, CUST_STORE_STATE, CUST_STORE_ZIP_CODE
            , CUST_STORE_PHONE, CUST_STORE_EXT, CUST_STORE_EMAIL
               from ARTCUST2 where CUST_CODE = :PARM1
            and (SELL_CODE_AC = :PARM2 or SELL_CODE = :PARM2) 
            and NVL(CUST_STORE_STATUS, 'A') = 'A'
            union
            Select 'S' ADDR_TYPE, SELL_CODE KEY_CODE, SELL_TYPE KEY_TYPE
            , SELL_NAME CUST_STORE_NAME, SELL_ADDR1 CUST_STORE_ADDR1, SELL_ADDR2 CUST_STORE_ADDR2, SELL_ADDR3 CUST_STORE_ADDR3
            , SELL_CITY CUST_STORE_CITY, SELL_STATE CUST_STORE_STATE, SELL_ZIP_CODE CUST_STORE_ZIP_CODE
            , SELL_PHONE CUST_STORE_PHONE, SELL_EXT CUST_STORE_EXT, SELL_EMAIL CUST_STORE_EMAIL
               from SOTSELL1 where (SELL_CODE = :PARM2 OR SELL_CODE_MGR = :PARM2)
            and NVL(SELL_STATUS, 'A') = 'A'
            union
            Select 'R' ADDR_TYPE, RSSP_CODE KEY_CODE, RSSP_TYPE KEY_TYPE
            , RSSP_SHIP_TO_NAME CUST_STORE_NAME, RSSP_SHIP_TO_ADDR1 CUST_STORE_ADDR1, RSSP_SHIP_TO_ADDR2 CUST_STORE_ADDR2, RSSP_SHIP_TO_ADDR3 CUST_STORE_ADDR3
            , RSSP_SHIP_TO_CITY CUST_STORE_CITY, RSSP_SHIP_TO_STATE CUST_STORE_STATE, RSSP_SHIP_TO_ZIP_CODE CUST_STORE_ZIP_CODE
            , RSSP_PHONE CUST_STORE_PHONE, RSSP_EXT CUST_STORE_EXT, RSSP_EMAIL CUST_STORE_EMAIL
               from SPTRSSP1 where SELL_CODE = :PARM2
            and NVL(RSSP_STATUS, 'A') = 'A'"

            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "VV", 2)

            'If rowSOTSELL1.Item("SELL_TYPE") & String.Empty = "AC" Then
            '    sqlARTCUST2 = "SELECT * FROM ARTCUST2 WHERE CUST_CODE = '" & AESalesOrderCustomer & "' AND (SELL_CODE = '" & rowSOTSELL1.Item("SELL_CODE_MGR") & "' or SELL_CODE = '" & AccountExecutive & "')"
            'Else
            '    sqlARTCUST2 = "SELECT * FROM ARTCUST2 WHERE CUST_CODE = '" & AESalesOrderCustomer & "' AND (SELL_CODE = '" & AccountExecutive & "' or SELL_CODE IN ( SELECT SELL_CODE FROM SOTSELL1 WHERE SELL_CODE_MGR = '" & AccountExecutive & "'))"
            'End If


            ASCMAIN1.sql = "Select SATCSTK3.* from SATCSTK3 where SATCSTK3.CSREQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "SATCSTK3", "**", 0, True, "V", 4)

            'Create_Relation("ARTCUSTX", "SATCSTK3", "CUST_CODE,CUST_STORE_NO")

            'With dst.Tables("SATCSTK3")
            '    .Columns.Add("CUST_STORE_NAME", GetType(System.String), "PARENT.CUST_STORE_NAME")
            'End With

        End With

        grdSATCSTKX.DataSource = dst.Tables("SATCSTKX")
        grdSATCSTK2.DataSource = dst.Tables("SATCSTK2")
        grdSATCSTK3.DataSource = dst.Tables("SATCSTK3")
        grdARTCUSTX.DataSource = dst.Tables("ARTCUSTX")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATCSTKX, grdSATCSTK2, grdSATCSTK3, grdARTCUSTX}
            grd.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.LightGreen
            grd.DisplayLayout.Override.ActiveRowAppearance.ForeColor = System.Drawing.Color.Black

            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            'If grd.Name = "grdICTITEMX" Then
            '    'grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '    'grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            '    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            'End If

            'If grd.Name = "grdSATSTKRS" Then
            '    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            'End If

            'With grd.DisplayLayout.Bands(0)
            '    For Each c As UltraWinGrid.UltraGridColumn In .Columns
            '        c.CellActivation = Activation.NoEdit
            '        c.Header.Appearance.BackColor = System.Drawing.Color.White
            '        c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            '        c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            '        If grd.Name = "grdSATSTKRS" Then
            '            If c.Key = "SELECTED" Then
            '                c.CellActivation = Activation.AllowEdit
            '            End If
            '            If c.Key = "RECORDS" Then
            '                c.Format = "#,##0"
            '            End If
            '        End If
            '    Next
            'End With
        Next

        ASCMAIN1.Add_Value_List(grdSATCSTK2, "COLLECTION_GENDER")
        ' ASCMAIN1.Add_Value_List(grdSATSTKRS, "COLLECTION_GENDER", Nothing, New String() {":", "M:Mens", "W:Womens", "A:All"})

        Create_Summary(grdSATCSTK2, "ITEM_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"

                If Absx1.txtFor("SELL_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify an AE before creating a New Car-Stock Request"
                Else
                    SELL_CODE = Absx1.txtFor("SELL_CODE").Text
                    If Not Validate_Code("SELL_CODE") Then
                        EMsg &= vbCr & $"Invalid AE Specified ({SELL_CODE})"
                    End If
                End If

            Case "View"

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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
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
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not ScreenMode
                    .Items("Edit").Visible = Not ScreenMode Or (EntryMode = "V")

                    .Items("Update").Visible = (EntryMode = "E" Or EntryMode = "N")
                    .Items("Cancel").Visible = (EntryMode = "E" Or EntryMode = "N")
                    .Items("Done").Visible = (EntryMode = "V")
                End With
            End With
        End If

        tabSATCSTKX.Visible = ScreenMode
        grdSATCSTKX.Visible = Not (ScreenMode)

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSATCSTK2.Visible = (ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATCSTKX", "SATCSTK1", "SATCSTK2", "SATCSTK3", "ARTCUSTX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'Fill_Records("SATCSTKX")
        'Sort_grdColumns(grdSATCSTKX)


    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            CSREQ_NO = ASCMAIN1.Next_Control_No("SATCSTK1.CSREQ_NO")

            rowSATCSTK1 = dst.Tables("SATCSTK1").NewRow
            With rowSATCSTK1
                .Item("CSREQ_NO") = CSREQ_NO
                .Item("SELL_CODE") = SELL_CODE

                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

            End With
            dst.Tables("SATCSTK1").Rows.Add(rowSATCSTK1)

            Dim DT1 As String = Format(CDate("10/01/2022"), "dd-MMM-yyyy")
            Dim DT2 As String = Format(CDate("10/08/2022"), "dd-MMM-yyyy")
            Fill_Records("SATCSTKA", New String() {DT1, DT1, DT2, DT2, "IPLBAE", "000" & SELL_CODE})

            dst.Tables("SATCSTK2").Rows.Clear()
            dst.Tables("SATCSTK3").Rows.Clear()
            Dim CSREQ_LNO As Integer = 0
            For Each row As DataRow In dst.Tables("SATCSTKA").Select("", "ITEM_CODE")

                Dim rowSATCSTK2 As DataRow = dst.Tables("SATCSTK2").NewRow
                With rowSATCSTK2
                    .Item("CSREQ_NO") = CSREQ_NO
                    CSREQ_LNO += 1
                    .Item("CSREQ_LNO") = CSREQ_LNO
                    For Each C As String In New String() _
                        {"ITEM_CODE", "ITEM_DESC", "ITEM_SO_QTY_MULT" _
                        , "COLLECTION_GENDER", "BRAND_CODE" _
                        , "DATE_START", "DATE_END", "QTY_ALLO", "ALLO_CTL_NO"}
                        .Item(C) = row.Item(C)
                    Next
                End With
                dst.Tables("SATCSTK2").Rows.Add(rowSATCSTK2)
            Next
        Else

            Fill_Records("SATCSTK2", CSREQ_NO)
        End If

        Fill_Records("ARTCUSTX", New String() {"IPLBAE", SELL_CODE})
        'Sort_grdColumns(grdARTCUSTX, "CUST_STORE_NO")



        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'ASCMAIN1.sql = $"Update ICTSREQ1 Set REPORTED = 'A', LAST_OPER = '{ASCMAIN1.USER_ID}', LAST_DATE = SYSDATE
        '    where (STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE) in
        '    (Select STOCK_REQ_CTL_NO, SELL_CODE, CUST_CODE, CUST_STORE_NO, ITEM_CODE from {SATSTKRX})"
        'ASCDATA1.ExecuteSQL()

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
                        'Me.Find_ITEM_CODE()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                ' Find_ITEM_CODE()
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
            'Case "RYW0"
            '    Set_Weeks()
            'Case "RYW1"
            '    Set_Weeks()
        End Select
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSTK2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSATCSTKX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTCUSTX, "SS", "Show Filter", "Show GroupBox")
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
               ' Load_from_XLS()


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
                'If chkActiveOnly.Checked Then sql_where &= " and ITEM_STATUS = 'A'"
                'If optSN.Value <> "*" Then sql_where &= $" and ITEM_SNU_CODE = '{ optSN.Value}'"
                'If optBP.Value <> "*" Then sql_where &= $" and ITEM_BASIC_PROMO = '{ optBP.Value}'"
        End Select
    End Sub

End Class