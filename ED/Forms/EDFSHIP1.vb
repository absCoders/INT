Imports Infragistics.Win.UltraWinGrid

' ABS change request  SR-6315 has been approved.

Public Class EDFSHIP1

    Private wkTable As String = String.Empty
    Private newSOTSHIP1 As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            wkTable = ASCMAIN1.Temp_Table("SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE SHIP_STATUS = 'P'")
            newSOTSHIP1 = ASCMAIN1.Temp_Table("SELECT * FROM SOTSHIP1 WHERE ROWNUM < 1")

            ASCMAIN1.sql = $"SELECT SOTSHIP1.SHIP_BOL_NO, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE,
                                SOTSHIP1.SHIP_VIA_CODE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.SHIP_856_IND, SOTSHIP1.WHSE_CODE, EDT945T1.EDI_BOL_NO, EDT945T1.EDI_MASTER_BOL_NO,
                                SOTSHIP1.SHIP_856_BATCH_NO
                                FROM SOTSHIP1, SOTORDR0,
                                (SELECT EDI_SHIPMENT_ID, MAX(EDI_BOL_NO) EDI_BOL_NO, MAX(EDI_MASTER_BOL_NO) EDI_MASTER_BOL_NO FROM EDT945T1 GROUP BY EDI_SHIPMENT_ID) EDT945T1
                                WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
                                AND SOTSHIP1.SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM {wkTable})
                                AND SOTSHIP1.SHIP_BOL_NO = EDT945T1.EDI_SHIPMENT_ID (+)"
            Create_TDA(.Tables.Add, "SOTSHIP1", ASCMAIN1.sql, 0, False, "", 0)

            ASCMAIN1.sql = $"SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTPICK1.SHIP_BOL_NO, TO_NUMBER(DECODE(NVL(SHIPHDR.ABSPICKNBR, ''), '', '0', '1')) RECD,
                                SOTORDR1.CUST_STORE_NO, NVL(SOTORDR1.CUST_STORE_LOCATION, ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION, 
                                SOTORDR2.EXT_PRICE, EDT945T1.EDI_BOL_NO, EDT945T1.EDI_MASTER_BOL_NO, SHIPHDR.IMPORT_DATE, SHIPHDR.IMPORT_FILENAME, SOTPICK1.ORDR_NO_3PL
                                FROM SOTPICK1, CONV.CFG_SHIPHDR SHIPHDR, SOTORDR1, EDT945T1, ARTCUST2,
                                (SELECT ORDR_NO, SUM(ORDR_QTY * ORDR_UNIT_PRICE) EXT_PRICE FROM SOTORDR2 GROUP BY ORDR_NO) SOTORDR2
                                WHERE SOTPICK1.SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM {wkTable})
                                AND SOTPICK1.PICK_NO = SHIPHDR.ABSPICKNBR (+)
                                AND SOTPICK1.PICK_STATUS = 'P'
                                AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO
                                AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO (+)
                                AND SOTPICK1.PICK_NO = EDT945T1.EDI_PICK_NO (+)
                                AND SOTORDR1.CUST_CODE = ARTCUST2.CUST_CODE (+)
                                AND SOTORDR1.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "SOTPICK1", ASCMAIN1.sql, 0, False, "", 0)

        End With

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        With dst.Tables("SOTSHIP1")
            .Columns.Add("NUM_TICKETS", GetType(Int32), "COUNT(CHILD.PICK_NO)")
            .Columns.Add("NUM_SHIPMENTS", GetType(Int32), "SUM(CHILD.RECD)")
            .Columns.Add("MISSING", GetType(Int32), "IIF(NUM_SHIPMENTS = 0, NULL, NUM_TICKETS - NUM_SHIPMENTS)")
            .Columns.Add("EXT_PRICE", GetType(Decimal), "SUM(CHILD.EXT_PRICE)")
        End With

        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        Create_Summary(grdSOTSHIP1, "CUST_CODE", "Count")
        Create_Summary(grdSOTSHIP1, "EXT_PRICE", "Sum")

        With grdSOTSHIP1.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("ORDR_CUST_PO").Header.Fixed = True
            .Columns("ORDR_GROUP_NO").Header.Fixed = True
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load", "Refresh"

            Case "Done"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load", "Refresh"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                End With

                '  Always show the filter
                ' .Groups("Filter").Visible = tf
            End With
        End If

        If ScreenMode Then

        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTSHIP1", "SOTPICK1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")

        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        ASCMAIN1.sql = $"TRUNCATE TABLE {wkTable}"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = $"INSERT INTO {wkTable} SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE SHIP_STATUS = 'P'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Fill_Records("SOTSHIP1")
        Fill_Records("SOTPICK1")

        optFilter_ValueChanged(Nothing, Nothing)

        EnforceConstraints(True)

        ASCMAIN1.Progress("", "")

        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIP1, "SSSPBBPBB", "Show Filter", "Show GroupBox", "Show Pins", "Partial ASN", "Email 3PL", "Auto Fit Columns", "Fix Duplicate BOL No")
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
        Dim tlb_btn1 As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
            Exit Sub
        End If

        Select Case e.SourceControl.Name
            Case grdSOTSHIP1.Name
                tlb_btn = DirectCast(tlb_pop.Tools("Partial ASN"), UltraWinToolbars.ButtonTool)
                tlb_btn1 = DirectCast(tlb_pop.Tools("Email 3PL"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key Then
                    Dim NUM_SHIPMENTS As Int32 = Val(grd.ActiveRow.Cells("NUM_SHIPMENTS").Value & String.Empty)
                    Dim MISSING As Int32 = Val(grd.ActiveRow.Cells("MISSING").Value & String.Empty)
                    tlb_btn.SharedProps.Enabled = NUM_SHIPMENTS > 0 AndAlso MISSING > 0
                    tlb_btn1.SharedProps.Enabled = NUM_SHIPMENTS > 0 AndAlso MISSING > 0
                Else
                    tlb_btn.SharedProps.Enabled = False
                    tlb_btn1.SharedProps.Enabled = False
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Fix Duplicate BOL No"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key Then
                    Dim NUM_SHIPMENTS As Int32 = Val(grd.ActiveRow.Cells("NUM_SHIPMENTS").Value & String.Empty)
                    Dim NUM_TICKETS As Int32 = Val(grd.ActiveRow.Cells("NUM_TICKETS").Value & String.Empty)
                    tlb_btn.SharedProps.Enabled = NUM_SHIPMENTS = NUM_TICKETS
                Else
                    tlb_btn.SharedProps.Enabled = False
                End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Const MT_LEVEL As Int16 = 987

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Auto Fit Columns"
                Me.Cursor = Cursors.WaitCursor
                grd.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
                Me.Cursor = Cursors.Default

            Case "Fix Duplicate BOL No"

                If Not ASCMAIN1.Logical_Lock("CLARINS", "IMPORT", False, True, True, MT_LEVEL) Then
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("R", "SORSHIPC", False, True, True, MT_LEVEL) Then
                    Exit Sub
                End If

                Dim EDI_BOL_NO As String = grd.ActiveRow.Cells("EDI_BOL_NO").Text
                If EDI_BOL_NO.Length = 0 Then
                    MessageBox.Show($"EDI BOL No may not be blank.", "Fix Duplicate BOL No", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Dim NUM_TICKETS As Int32 = Val(grd.ActiveRow.Cells("NUM_TICKETS").Text & "")
                Dim NUM_SHIPMENTS As Int32 = Val(grd.ActiveRow.Cells("NUM_SHIPMENTS").Text & "")

                If NUM_TICKETS <> NUM_SHIPMENTS Then
                    MessageBox.Show($"All Shipment records must be received before this process can be executed.", "Fix Duplicate BOL No", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Dim rowSOTSHIPB As DataRow = LookUp("SOTSHIPB", EDI_BOL_NO)
                If rowSOTSHIPB Is Nothing Then
                    MessageBox.Show($"Cannot locate EDI BOL No {EDI_BOL_NO} in SOTSHIPB.", "Fix Duplicate BOL No", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Dim tblSOTSHIP1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = :PARM1 ", "", "V", {EDI_BOL_NO})
                If tblSOTSHIP1.Rows.Count = 0 Then
                    MessageBox.Show($"Cannot locate EDI BOL No {EDI_BOL_NO} in SOTSHIP1.BILL_OF_LADING_NO", "Fix Duplicate BOL No", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Try
                    Try
                        BeginTrans()

                        Dim newBOL_NO As String = String.Empty
                        For Each ch As Char In "ABCDEFGHIJKLMNPRSTUVWXYZ"
                            newBOL_NO = ch & EDI_BOL_NO.Substring(1)
                            ASCMAIN1.Progress("Evaluate Previous BOLs", newBOL_NO)
                            rowSOTSHIPB = LookUp("SOTSHIPB", newBOL_NO)
                            If rowSOTSHIPB Is Nothing Then
                                Exit For
                            End If
                            newBOL_NO = String.Empty
                        Next

                        If newBOL_NO.Length > 0 Then
                            ASCMAIN1.sql = "UPDATE SOTSHIPB SET BOL_NO = :PARM1 WHERE BOL_NO = :PARM2"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {newBOL_NO, EDI_BOL_NO})

                            ASCMAIN1.sql = "UPDATE SOTSHIP1 SET BILL_OF_LADING_NO = :PARM1 WHERE BILL_OF_LADING_NO = :PARM2"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {newBOL_NO, EDI_BOL_NO})

                            CommitTrans("History BOL No updated.")
                        Else
                            Rollback("Could not find an Alternate BOL No.")
                        End If

                        ' Click_Command("Refresh")

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Fix Duplicate BOL No")
                    Exit Sub
                Finally
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    ASCMAIN1.Progress("", "")
                End Try

            Case "Email 3PL"

                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Text
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Text
                Dim NUM_TICKETS As Int32 = Val(grd.ActiveRow.Cells("NUM_TICKETS").Text & "")
                Dim NUM_SHIPMENTS As Int32 = Val(grd.ActiveRow.Cells("NUM_SHIPMENTS").Text & "")
                Dim MISSING As Int32 = Val(grd.ActiveRow.Cells("MISSING").Text & "")
                Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Text
                Dim EDI_BOL_NO As String = grd.ActiveRow.Cells("EDI_BOL_NO").Text

                If NUM_SHIPMENTS = 0 Then
                    MessageBox.Show($"Shipment {SHIP_BOL_NO} does not have any shipment files.", "Email 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Try
                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        MessageBox.Show($"Cannot locate Shipment {SHIP_BOL_NO}", "Email 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If rowSOTSHIP1.Item("SHIP_STATUS") & String.Empty <> "P" Then
                        MessageBox.Show($"Shipment {SHIP_BOL_NO} does not have a status of Pick", "Email 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    ASCMAIN1.sql = "SELECT PICK_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = :PARM1 and PICK_STATUS = 'P'
                                        MINUS
                                    SELECT ABSPICKNBR FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO = :PARM1"
                    Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {SHIP_BOL_NO})
                    If tbl.Rows.Count = 0 Then
                        MessageBox.Show($"Shipment {SHIP_BOL_NO} appears to have all of its shipment records.", "Email 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim lstPTs As New List(Of String)
                    For Each rowSOTPICK1 As DataRow In tbl.Select("")
                        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                        lstPTs.Add(PICK_NO)
                    Next

                    lstPTs.Sort()

                    Dim note As String = String.Empty

                    note = "We are missing the following shipment confirmation records" & Environment.NewLine
                    'note &= $"Customer: {CUST_CODE}" & Environment.NewLine
                    'note &= $"PO No: {ORDR_CUST_PO}" & Environment.NewLine
                    'note &= $"3PL BOL No: {EDI_BOL_NO}" & Environment.NewLine
                    'note &= Environment.NewLine
                    'note &= Environment.NewLine

                    note &= "<html>" '& Environment.NewLine
                    note &= "<head>" '& Environment.NewLine

                    note &= "<STYLE TYPE=""text/css"">" '& Environment.NewLine
                    note &= "<!--"
                    note &= "TD{font-family: Arial; font-size: 9pt;}" '& Environment.NewLine
                    note &= "--->"
                    note &= "</STYLE>" '& Environment.NewLine

                    note &= "</head>" '& Environment.NewLine
                    note &= "<body>" '& Environment.NewLine

                    note &= "<TABLE CLASS=""TD"" BORDER=""0"" CELLPADDING =""5"" CELLSPACING=""5"" WIDTH=""800"">" '& Environment.NewLine

                    note &= " <TR>"
                    note &= "   <TH style=""text-align:left"" ""font-weight:normal"" width=""200"">Missing Shipment Data</TH>"
                    note &= " </TR>" '& Environment.NewLine

                    note &= " <TR>"
                    note &= $"   <TD>Customer: {CUST_CODE}</TD>"
                    note &= " </TR>" '& Environment.NewLine

                    note &= " <TR>"
                    note &= $"   <TD>PO No: {ORDR_CUST_PO}</TD>"
                    note &= " </TR>" ' & Environment.NewLine

                    note &= " <TR>"
                    note &= $"   <TD>3PL BOL No: {EDI_BOL_NO}</TD>"
                    note &= " </TR>" '& Environment.NewLine

                    note &= " <TR>"
                    note &= "   <TD></TD>"
                    note &= "   <TD></TD>"
                    note &= "   <TD></TD>"
                    note &= " </TR>" '& Environment.NewLine

                    note &= " <TR>"
                    note &= "   <TH style=""text-align:left"" width=""150"">Our Pick No</TH>"
                    note &= "   <TH style=""text-align:left"" width=""150"">Our Order No</TH>"
                    note &= "   <TH style=""text-align:left"" width=""150"">Your Order No</TH>"
                    note &= "   <TH style=""text-align:left"" width=""350""></TH>"
                    note &= " </TR>" '& Environment.NewLine

                    ASCMAIN1.sql = "SELECT * FROM SOTPICK1 WHERE PICK_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
                    Dim tblPT As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", "C", {String.Join(",", lstPTs.ToArray)})
                    For Each row As DataRow In tblPT.Select("", "PICK_NO")
                        note &= " <TR>"
                        note &= $"   <TD>{row.Item("PICK_NO")}</TD>"
                        note &= $"   <TD>{row.Item("ORDR_NO")}</TD>"
                        note &= $"   <TD>{row.Item("ORDR_NO_3PL")}</TD>"
                        note &= "   <TD></TD>"
                        note &= " </TR>" '& Environment.NewLine
                    Next

                    note &= "</TABLE>" '& Environment.NewLine
                    note &= "</body>" '& Environment.NewLine
                    note &= "</html>" '& Environment.NewLine

                    Dim objASCNOTEE As New TAC.ASCNOTEE(ASCMAIN1.Folders, "EDFSHIP1_" & WHSE_CODE, dst)
                    objASCNOTEE.Note = note
                    objASCNOTEE.CreateComponents()
                    objASCNOTEE.EmailDocument()

                    MessageBox.Show("Email sent.", "Email 3PL")

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Email 3PL")
                End Try

            Case "Partial ASN"

                If Not ASCMAIN1.Logical_Lock("CLARINS", "IMPORT", False, True, True, MT_LEVEL) Then
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("R", "SORSHIPC", False, True, True, MT_LEVEL) Then
                    Exit Sub
                End If

                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Text
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Text
                Dim NUM_TICKETS As Int32 = Val(grd.ActiveRow.Cells("NUM_TICKETS").Text & "")
                Dim NUM_SHIPMENTS As Int32 = Val(grd.ActiveRow.Cells("NUM_SHIPMENTS").Text & "")
                Dim MISSING As Int32 = Val(grd.ActiveRow.Cells("MISSING").Text & "")

                If NUM_SHIPMENTS = 0 Then
                    MessageBox.Show($"Shipment {SHIP_BOL_NO} does Not have any shipment files To process.", "Partial ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Dim zMsg As String = $"Do you want To create a Partial ASN For Shipment {SHIP_BOL_NO}, Customer {CUST_CODE}, P.O. {ORDR_CUST_PO}?"
                If MessageBox.Show(zMsg, "Partial ASN", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                    Exit Sub
                End If

                Try
                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        MessageBox.Show($"Cannot locate Shipment {SHIP_BOL_NO}", "Partial ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                        Exit Sub
                    End If

                    If rowSOTSHIP1.Item("SHIP_STATUS") & String.Empty <> "P" Then
                        MessageBox.Show($"Shipment {SHIP_BOL_NO} does Not have a status of Pick", "Partial ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                        Exit Sub
                    End If

                    ASCMAIN1.sql = "SELECT PICK_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = :PARM1 and PICK_STATUS = 'P'
                                        MINUS
                                    SELECT ABSPICKNBR FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO = :PARM1"
                    Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {SHIP_BOL_NO})
                    If tbl.Rows.Count = 0 Then
                        MessageBox.Show($"Shipment {SHIP_BOL_NO} appears to have all of its shipment records.", "Partial ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                        Exit Sub
                    Else
                        Dim lstPTs As New List(Of String)
                        For Each rowSOTPICK1 As DataRow In tbl.Select("")
                            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                            lstPTs.Add(PICK_NO)
                        Next

                        lstPTs.Sort()

                        If MessageBox.Show($"The following Pick Tickets will be placed on another Shipment record.{Environment.NewLine}{Environment.NewLine}{String.Join(Environment.NewLine, lstPTs.ToArray)}{Environment.NewLine}{Environment.NewLine}Do you want to continue?", "Partial ASN", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                            Exit Sub
                        End If

                    End If

                    Try
                        Dim newHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")

                        BeginTrans()

                        ASCDATA1.ExecuteSQL($"DELETE FROM {newSOTSHIP1}")
                        ASCDATA1.ExecuteSQL($"INSERT INTO {newSOTSHIP1} SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO = :PARM1", "V", {SHIP_BOL_NO})
                        ASCDATA1.ExecuteSQL($"UPDATE {newSOTSHIP1} SET SHIP_BOL_NO = :PARM1", "V", {newHIP_BOL_NO})
                        ASCDATA1.ExecuteSQL($"INSERT INTO SOTSHIP1 SELECT * FROM {newSOTSHIP1}", "V", {SHIP_BOL_NO})

                        For Each rowSOTPICK1 As DataRow In tbl.Select("")
                            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                            ASCDATA1.ExecuteSQL($"UPDATE SOTPICK1 SET SHIP_BOL_NO = :PARM1 WHERE PICK_NO = :PARM2", "VV", {newHIP_BOL_NO, PICK_NO})
                        Next

                        CommitTrans("Partial ASN created.")
                        Click_Command("Refresh")
                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Partial ASN")
                Finally
                    ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                End Try

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTSHIP1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSHIP1.InitializeRow

        e.Row.Appearance.BackColor = Nothing

        Select Case e.Row.Band.Key
            Case grdSOTSHIP1.DisplayLayout.Bands(0).Key
                Dim NUM_SHIPMENTS As Int32 = Val(e.Row.Cells("NUM_SHIPMENTS").Value & String.Empty)
                Dim MISSING As Int32 = Val(e.Row.Cells("MISSING").Value & String.Empty)

                If NUM_SHIPMENTS > 0 AndAlso MISSING > 0 Then
                    e.Row.Appearance.BackColor = Drawing.Color.LightBlue
                End If

            Case grdSOTSHIP1.DisplayLayout.Bands(1).Key
                Dim MISSING As Int32 = Val(e.Row.ParentRow.Cells("MISSING").Value & String.Empty)
                If MISSING > 0 Then
                    If Val(e.Row.Cells("RECD").Value & String.Empty) = 0 Then
                        e.Row.Appearance.BackColor = Drawing.Color.LightBlue
                    End If
                End If

        End Select

    End Sub

    Private Sub optFilter_ValueChanged(sender As Object, e As EventArgs) Handles optFilter.ValueChanged

        If Not dst.Tables.Contains("SOTSHIP1") Then
            Exit Sub
        End If

        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")

        Dim dvwSOTSHIP1 As DataView
        dvwSOTSHIP1 = dst.Tables("SOTSHIP1").DefaultView
        Select Case optFilter.Value
            Case "A"
                dvwSOTSHIP1.RowFilter = String.Empty
            Case "M"
                dvwSOTSHIP1.RowFilter = "ISNULL(MISSING, 0) > 0"
        End Select

        grdSOTSHIP1.DataSource = dvwSOTSHIP1
        Sort_grdColumns(grdSOTSHIP1, "missing")

    End Sub

#End Region

End Class