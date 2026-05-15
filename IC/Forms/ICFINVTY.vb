Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinExplorerBar
Imports System.IO
Public Class ICFINVTY
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Dim ICTINVTY As String
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dteSTART_DATE.Value = DateTime.Today.AddDays(-7) ' 7 days ago
        dteEND_DATE.Value = DateTime.Today ' Today
        With dst
            ASCMAIN1.sql = "SELECT ICTIADJ2.OPS_YYYYPP, ICTIADJ2.ITEM_CODE, ICTIADJ1.WHSE_CODE" _
            & ", ICTIADJ1.ADJ_NO TRAN_NO, ICTIADJ1.ADJ_SOURCE TRAN_SOURCE" _
            & ", ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER" _
            & ", ICTIADJ1.ADJ_DATE TRAN_DATE, 'X' TRAN_TYPE" _
            & ", ICTIADJ2.ADJ_QTY TRAN_QTY, ICTIADJ1.ADJ_NOTE TRAN_NOTE" _
            & " FROM ICTIADJ1,ICTIADJ2 WHERE ROWNUM < 1"
            ICTINVTY = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select ICTINVTY.* from " & ICTINVTY & " ICTINVTY"
            Create_TDA(.Tables.Add, "ICTINVTY", "**", 0, False)
            With .Tables("ICTINVTY").Columns
                For Each C As String In New String() {"S", "R", "P", "A", "T", "X"}
                    .Add("TRAN_QTY_" & C, GetType(System.Int64), "IIF(TRAN_TYPE='" & C & "',TRAN_QTY,0)")
                Next
            End With
        End With

        grdICTTRANX.DataSource = dst.Tables("ICTINVTY")
        For Each C As String In New String() {"S", "R", "P", "A", "T", "X"}
            Dim COLUMN_NAME As String = "TRAN_QTY_" & C
            With grdICTTRANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .Width = 70
                Dim DX As String = ""
                Select Case C
                    Case "S"
                        DX = "Shp"
                    Case "R"
                        DX = "Rtn"
                    Case "P"
                        DX = "Rec"
                    Case "A"
                        DX = "Adj"
                    Case "T"
                        DX = "Xfr"
                    Case "X"
                        DX = "Con"
                End Select
                .Header.Caption = DX
            End With
            Create_Summary(grdICTTRANX, COLUMN_NAME)
        Next

        Create_Summary(grdICTTRANX, "TRAN_DATE", "Count")
        Create_Summary(grdICTTRANX, "TRAN_QTY")

        With grdICTTRANX.DisplayLayout.Bands("ICTINVTY")
            .Columns("TRAN_DATE").Header.Fixed = True
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If txtWHSE_CODE.Text & "" = "" Then
                    EMsg &= vbCr & "Warehouse Code is required."
                End If

                If dteSTART_DATE.Value > dteEND_DATE.Value Then
                    EMsg &= vbCr & "Start Date cannot be later than End Date."
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Set_Read_Only(UltraGroupBox2, ScreenMode)

        grdICTTRANX.Visible = ScreenMode
        If Not ScreenMode Then
            Clear_Record()
        End If

    End Sub
    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTINVTY"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub
    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Inventory Movements..")
        Me.Cursor = Cursors.WaitCursor

        Try
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTINVTY)

            Dim WHSE_CODE As String = txtWHSE_CODE.Text
            Dim fromDate As String = Format(dteSTART_DATE.Value, "yyyy-MM-dd")
            Dim toDate As String = Format(dteEND_DATE.Value, "yyyy-MM-dd")

            grdICTTRANX.Text = $"Transactions for Warehouse {WHSE_CODE} between {fromDate} and {toDate}"

            'Inventory Adjustments
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.ADJ_NO TRAN_NO, T1.ADJ_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.ADJ_DATE TRAN_DATE, 'A' TRAN_TYPE" _
            & ", T2.ADJ_QTY TRAN_QTY, ICTREAS1.REASON_DESC TRAN_NOTE" _
            & " FROM ICTIADJ1 T1, ICTIADJ2 T2, ICTREAS1" _
            & " WHERE T1.ADJ_NO = T2.ADJ_NO" _
            & " AND T1.WHSE_CODE = '" & WHSE_CODE & "'" _
            & " AND T1.ADJ_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')" _
            & " AND ICTREAS1.REASON_CODE = T1.REASON_CODE"
            ASCDATA1.ExecuteSQL()

            ' Product Consumption
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T1.OPS_YYYYPP, T4.ITEM_CODE, T2.VEND_WHSE_CODE WHSE_CODE" _
            & ", T1.RECEIPT_NO TRAN_NO, 'X' TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.RECEIPT_DATE TRAN_DATE, 'X' TRAN_TYPE" _
            & ", -1 * T4.QTY_CON TRAN_QTY, 'PO ' || T2.PO_ORDER_NO || ', BM# ' || T2.BM_ISSUE_NO TRAN_NOTE" _
            & " FROM ICTIREC1 T1, ICTIREC2 T2, ICTIREC4 T4" _
            & " WHERE T1.RECEIPT_NO = T2.RECEIPT_NO" _
            & " AND T2.VEND_WHSE_CODE = '" & WHSE_CODE & "'" _
            & " AND T1.RECEIPT_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')" _
            & " AND T1.RECEIPT_NO = T4.RECEIPT_NO" _
            & " AND T2.RECEIPT_LNO = T4.RECEIPT_LNO"
            ASCDATA1.ExecuteSQL()

            ' Warehouse Transfers
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
            & ", -1 * T2.XFR_QTY TRAN_QTY, 'XFR to ' || T1.WHSE_CODE_TO TRAN_NOTE" _
            & " FROM ICTIXFR1 T1, ICTIXFR2 T2" _
            & " WHERE T1.XFR_NO = T2.XFR_NO" _
            & " AND T1.WHSE_CODE = '" & WHSE_CODE & "'" _
            & " AND T1.XFR_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')"
            ASCDATA1.ExecuteSQL()

            ' Inbound Transfers
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE_TO" _
            & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
            & ", T2.XFR_QTY TRAN_QTY, 'XFR from ' || T1.WHSE_CODE TRAN_NOTE" _
            & " FROM ICTIXFR1 T1, ICTIXFR2 T2" _
            & " WHERE T1.XFR_NO = T2.XFR_NO" _
            & " AND T1.WHSE_CODE_TO = '" & WHSE_CODE & "'" _
            & " AND T1.XFR_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')"
            ASCDATA1.ExecuteSQL()

            ' PO Receipts
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T1.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.RECEIPT_NO TRAN_NO, 'P' TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.RECEIPT_DATE TRAN_DATE, 'P' TRAN_TYPE" _
            & ", T2.QTY_REC TRAN_QTY, 'PO ' || T1.PO_ORDER_NO TRAN_NOTE" _
            & " FROM ICTIREC1 T1, ICTIREC2 T2" _
            & " WHERE T1.RECEIPT_NO = T2.RECEIPT_NO" _
            & " AND T1.WHSE_CODE = '" & WHSE_CODE & "'" _
            & " AND T1.RECEIPT_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')"
            ASCDATA1.ExecuteSQL()

            ' Sales Shipments and Returns
            ASCMAIN1.sql = "INSERT INTO " & ICTINVTY _
            & " SELECT T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", DECODE(T2.INV_TYPE, 'I', 'Shp', 'C', DECODE(S1.RA_NO, NULL, 'Rtn', 'RA ' || S1.RA_NO), '?') TRAN_NO, 'S' TRAN_SOURCE" _
            & ", NULL INIT_DATE, NULL INIT_OPER" _
            & ", T1.INV_DATE TRAN_DATE, DECODE(T2.INV_TYPE, 'I', 'S', 'C', 'R', NULL) TRAN_TYPE" _
            & ", SUM(-1 * T2.ORDR_QTY_SHIP) TRAN_QTY, 'Line Items:' || COUNT(*) TRAN_NOTE" _
            & " FROM SOTINVH1 T1, SOTINVH2 T2, SOTRTRN1 S1" _
            & " WHERE T1.INV_NO = T2.INV_NO" _
            & " AND T1.WHSE_CODE = '" & WHSE_CODE & "'" _
            & " AND T1.INV_DATE BETWEEN TO_DATE('" & fromDate & "', 'YYYY-MM-DD') AND TO_DATE('" & toDate & "', 'YYYY-MM-DD')" _
            & " AND T1.INV_NO = S1.INV_NO(+)" _
            & " GROUP BY T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE, DECODE(T2.INV_TYPE, 'I', 'Shp', 'C', DECODE(S1.RA_NO, NULL, 'Rtn', 'RA ' || S1.RA_NO), '?'), T1.INV_DATE, DECODE(T2.INV_TYPE, 'I', 'S', 'C', 'R', NULL)"
            ASCDATA1.ExecuteSQL()

            dst.Tables("ICTINVTY").Rows.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please Call ABS - Problem mapping Transactions")
        End Try

        Dim sql As String = "Select * from " & ICTINVTY

        Fill_Records("ICTINVTY", "", True, sql)
        Sort_grdColumns(grdICTTRANX, "ITEM_CODE, TRAN_DATE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTTRANX, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.Keys.Contains(Mid(e.SourceControl.Name, 4)) Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing OrElse grd.Name <> "grdICTTRANX" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsFilterRow OrElse grd.ActiveRow.IsAddRow OrElse grd.ActiveRow.IsGroupByRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
            End Select

        End If
    End Sub
    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value & ""
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region
End Class