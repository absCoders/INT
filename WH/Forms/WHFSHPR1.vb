Public Class WHFSHPR1

    Private searchField As String = String.Empty
    Private searchValue As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "SELECT CUST_CODE, ORDR_CUST_PO, INV_SALES, VALUE FROM (" & vbCrLf _
                & "SELECT SOTINVH1.CUST_CODE, ARTCUST2.CUST_NO_3PL, SOTINVH1.ORDR_CUST_PO, SUM (SOTINVH1.INV_SALES) INV_SALES " & vbCrLf _
                & "FROM SOTINVH1,ARTCUST2" & vbCrLf _
                & " WHERE SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM1 " & vbCrLf _
                & "  and (TRUNC(SOTINVH1.INIT_DATE) <> TRUNC(SYSDATE) OR :PARM2 = '0')" & vbCrLf _
                & "AND ARTCUST2.CUST_CODE = SOTINVH1.CUST_CODE AND ARTCUST2.CUST_STORE_NO = SOTINVH1.CUST_STORE_NO" & vbCrLf _
                & "AND SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.WHSE_CODE LIKE 'CLA%'" & vbCrLf _
                & "GROUP BY SOTINVH1.CUST_CODE, ARTCUST2.CUST_NO_3PL, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ") ABSS, (" & vbCrLf _
                & "SELECT OHCUS1, OHCSPO, SUM (OHSSHIPVALUE ) VALUE" & vbCrLf _
                & "FROM CONV.CFG_DAILYSHIP  WHERE TO_CHAR(SHIPDATE,'YYYYMM') = :PARM3" & vbCrLf _
                & "GROUP BY OHCUS1, OHCSPO" & vbCrLf _
                & ") CFGS" & vbCrLf _
                & "WHERE ABSS.CUST_NO_3PL = CFGS.OHCUS1(+) AND ABSS.ORDR_CUST_PO = CFGS.OHCSPO(+)"

            Create_TDA(.Tables.Add, "WHTSHPRA", "**", 0, False, "VVV", 0)
            .Tables("WHTSHPRA").Columns.Add("DIFF", GetType(System.Decimal), "ISNULL(INV_SALES,0) - ISNULL(VALUE,0)")

            ASCMAIN1.sql = "SELECT OHCUS1, OHCSPO, SUM (OHSSHIPVALUE ) VALUE" & vbCrLf _
              & "FROM CONV.CFG_DAILYSHIP  WHERE TO_CHAR(SHIPDATE,'YYYYMM') = :PARM1" & vbCrLf _
              & "GROUP BY OHCUS1, OHCSPO"

            Create_TDA(.Tables.Add, "WHTSHPRC", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "CFG_SHIPHDR", "Select * from CONV.CFG_SHIPHDR")
            Create_TDA(.Tables.Add, "CFG_SHIPDTL", "Select * from CONV.CFG_SHIPDTL")
            Create_TDA(.Tables.Add, "CFG_CARTON", "Select * from CONV.CFG_CARTON")
            Create_TDA(.Tables.Add, "CFG_TRACK", "Select * from CONV.CFG_TRACK")

            .Tables("CFG_SHIPHDR").Columns("OHINVN").MaxLength = 20

        End With

        grdWHTSHPRA.DataSource = dst.Tables("WHTSHPRA")
        grdWHTSHPRC.DataSource = dst.Tables("WHTSHPRC")

        Create_Summary(grdWHTSHPRA, New String() {"INV_SALES", "VALUE", "DIFF"})
        Create_Summary(grdWHTSHPRA, "CUST_CODE", "Count")

        Create_Summary(grdWHTSHPRC, New String() {"VALUE"})
        Create_Summary(grdWHTSHPRC, "OHCUS1", "Count")

        Create_Relation("CFG_SHIPHDR", "CFG_SHIPDTL", "OHINVN", "SAINVN")
        Create_Relation("CFG_SHIPHDR", "CFG_CARTON", "OHINVN", "CHINVN")
        Create_Relation("CFG_SHIPHDR", "CFG_TRACK", "OHINVN", "INVN")
        grdShipHdr.DataSource = dst.Tables("CFG_SHIPHDR")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"WHTSHPRA", "WHTSHPRC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        lblRecMatch.Visible = False
        lblRecOff.Visible = False

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Dim YP As String = Absx1.txtFor("OPS_YYYYPP").Text
        grdWHTSHPRA.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim EXCL As String = IIf(chkExcludeToday.Checked, "1", "0")
        Fill_Records("WHTSHPRA", New String() {YP, EXCL, YP})
        Fill_Records("WHTSHPRC", New String() {YP})

        EnforceConstraints(True)

        Dim INV_SALES As Decimal = Val(dst.Tables("WHTSHPRA").Compute("SUM(INV_SALES)", ""))
        Dim VALUE As Decimal = Val(dst.Tables("WHTSHPRA").Compute("SUM(VALUE)", ""))
        If INV_SALES = VALUE Then
            lblRecMatch.Visible = True
        Else
            lblRecOff.Text = "Sales Off by " & Format(INV_SALES - VALUE, "$#,##0.00")
            lblRecOff.Visible = True
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTSHPRA, "SS", "Show Filter", "Show Groupbox")
        Load_Popup_Menu(grdWHTSHPRC, "SS", "Show Filter", "Show Groupbox")
        Load_Popup_Menu(grdShipHdr, "SS", "Show Filter", "Show Groupbox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdX"


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case ""


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Account Inquiry"


        End Select
    End Sub

#End Region

    Private Sub grdWHTSHPRA_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTSHPRA.InitializeLayout

    End Sub

    Private Sub grdWHTSHPRA_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWHTSHPRA.InitializeRow
        Dim INV_SALES As Decimal = Convert.ToDecimal(e.Row.Cells("INV_SALES").Value & "")
        Dim VALUE As Decimal = Convert.ToDecimal(Val(e.Row.Cells("VALUE").Value & ""))
        If INV_SALES <> VALUE Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        Try
            Static wkTable As String = String.Empty

            Dim value As String = txtValue.Text.Trim
            If value.Length = 0 Then
                MessageBox.Show("provide the search value.")
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor

            value = value.PadLeft(10, "0")
            txtValue.Text = value

            searchField = String.Empty
            searchValue = String.Empty

            Select Case optSearch.Value
                Case "P"
                    ASCMAIN1.sql = "SELECT SHIP_BOL_NO FROM SOTPICK1 WHERE PICK_NO = :PARM1"
                    searchField = "ABSPICKNBR"
                    searchValue = value
                Case "O"
                    ASCMAIN1.sql = "SELECT SHIP_BOL_NO FROM SOTPICK1 WHERE ORDR_NO = :PARM1"
                    searchField = "ABSORDNBR"
                    searchValue = value
                Case "S"
                    ASCMAIN1.sql = "SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE SHIP_BOL_NO = :PARM1"
            End Select

            If wkTable.Length = 0 Then
                wkTable = ASCMAIN1.Temp_Table("SELECT SHIP_BOL_NO FROM CONV.CFG_SHIPHDR WHERE ROWNUM < 1")
            End If

            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & wkTable)

            ASCDATA1.ExecuteSQL("INSERT INTO " & wkTable & " " & ASCMAIN1.sql, "V", value)

            EnforceConstraints(False)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & ")"
            Fill_Records("CFG_SHIPHDR", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT 'Missing ' || ROWNUM OHINVN, PICK_NO ABSPICKNBR, ORDR_NO ABSORDNBR, SHIP_BOL_NO FROM SOTPICK1 WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & ")" _
                & " AND PICK_NO NOT IN (SELECT ABSPICKNBR FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & "))"
            Fill_Records("CFG_SHIPHDR", String.Empty, False, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_SHIPDTL WHERE SAINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & "))"
            Fill_Records("CFG_SHIPDTL", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_CARTON WHERE CHINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & "))"
            Fill_Records("CFG_CARTON", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_TRACK WHERE INVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & wkTable & "))"
            Fill_Records("CFG_TRACK", String.Empty, True, ASCMAIN1.sql)

            EnforceConstraints(True)

            grdShipHdr.DisplayLayout.Bands(1).Hidden = False
            grdShipHdr.DisplayLayout.Bands(2).Hidden = False
            grdShipHdr.DisplayLayout.Bands(3).Hidden = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub grdShipHdr_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdShipHdr.InitializeRow

        If searchField.Length > 0 AndAlso searchValue.Length > 0 Then
            If e.Row.Band.Key = grdShipHdr.DisplayLayout.Bands(0).Key Then
                If e.Row.Cells(searchField).Value = searchValue Then
                    e.Row.Appearance.BackColor = Drawing.Color.LightGreen
                Else
                    e.Row.Appearance.BackColor = Nothing
                End If
            End If
        End If

    End Sub
End Class