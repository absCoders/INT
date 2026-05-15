Imports ABSolution

Public Class EDF81056

    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "   SELECT DISTINCT '0' SEL,  '0' EDI_STATUS, 'B' SEL_TYPE, SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_856_IND, "
            ASCMAIN1.sql &= "   SOTSHIP1.SHIP_810_IND, SOTSHIP1.SHIP_VIA_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.FRT_TERMS, SOTSHIP1.WHSE_CODE, SOTSHIP1.EDI_810_CREATED, SOTSHIP1.EDI_856_CREATED"
            ASCMAIN1.sql &= "   FROM SOTSHIP1, SOTORDR0"
            ASCMAIN1.sql &= "   WHERE ((NVL(SHIP_810_IND, '0') = '1' AND SHIP_810_BATCH_NO IS NULL)"
            ASCMAIN1.sql &= "   OR (NVL(SHIP_856_IND, '0') = '1' AND SHIP_856_BATCH_NO IS NULL))"
            ASCMAIN1.sql &= "   AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO "
            ASCMAIN1.sql &= "   AND SOTSHIP1.SHIP_STATUS = 'F'"
            ASCMAIN1.sql &= "   AND  SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL"

            ' Clarins recycles the Bil Of Lading No.
            ASCMAIN1.sql &= "   AND  SOTSHIP1.SHIP_DATE_SHIPPED > SYSDATE - 90"

            Create_TDA(.Tables.Add, "SOTSHIPX", ASCMAIN1.sql, 0, False, String.Empty, 0)

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT810O1", "*")
            Create_TDA(.Tables.Add, "EDT810O2", "*")
            Create_TDA(.Tables.Add, "EDT810O3", "*")
            Create_TDA(.Tables.Add, "EDT810O5", "*")

            Create_TDA(.Tables.Add, "EDT856O1", "*")
            Create_TDA(.Tables.Add, "EDT856O2", "*")
            Create_TDA(.Tables.Add, "EDT856O3", "*")
            Create_TDA(.Tables.Add, "EDT856O4", "*")
            Create_TDA(.Tables.Add, "EDT856O5", "*")

        End With

        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIPX")
        Create_Summary(grdSOTSHIP1, "BILL_OF_LADING_NO", "Count")
        Create_Summary(grdSOTSHIP1, "SEL", "Sum")

        With grdSOTSHIP1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With

        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTSHIP1.DisplayLayout.Bands(0).Columns
            If col.Key = "SEL" Then
                col.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1", "EDTTRPM1", String.Empty, Nothing)
        tblEDTSLSP1 = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1")
        tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Generate"

                ' Added 3/7/2018
                If ASCMAIN1.CLIENT = "INT" Then
                    If Not ASCMAIN1.Logical_Lock("CLARINS", "IMPORT", False, True) Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("R", "SORSHIPC", False, False, True) Then
                        Exit Sub
                    End If
                End If

                If dst.Tables("SOTSHIPX").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select at least one shipment."
                End If

                If EMsg.Length = 0 Then
                    If MessageBox.Show("Do you want to process the selected EDI Transactions?", "EDI Post", _
                                       MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    If dst.Tables("SOTSHIPX").Select("SEL = '1' AND EDI_STATUS = '1'").Length > 0 Then
                        MessageBox.Show("Please notify ABS that you are processing EDI transactions for customers set on Test Mode.", "EDI Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Generate"
                Me.Load_Record()
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Visible = False
            End With
        End If

        If ScreenMode Then

        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)

        For Each table As String In New String() {"SOTSHIP1", "EDTSYSIH", "EDT810O1", "EDT810O2", "EDT810O3", "EDT810O5", _
                                                  "EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5", "SOTSHIPX"}
            dst.Tables(table).Rows.Clear()
        Next

        ' Update BILL_OF_LADING_NO on SOTSHIP1 where there is an EDT945T1 entry
        ASCMAIN1.sql = "Begin Declare Cursor C1 is Select * from EDT945T1 Where EDI_SHIPMENT_ID IN"
        ASCMAIN1.sql &= " ("
        ASCMAIN1.sql &= " Select ship_bol_no from sotship1 where SHIP_STATUS = 'F' and SHIP_810_IND = '1' and edi_810_created is null"
        ASCMAIN1.sql &= " );"
        ASCMAIN1.sql &= " Begin for R1 in C1 Loop"
        ASCMAIN1.sql &= "   update sotship1 set bill_of_lading_no = r1.edi_bol_no where ship_bol_no = r1.edi_shipment_id;"
        ASCMAIN1.sql &= " End Loop; End; End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "   SELECT DISTINCT '1' SEL,  '0' EDI_STATUS, 'B' SEL_TYPE, SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_856_IND, "
        ASCMAIN1.sql &= "   SOTSHIP1.SHIP_810_IND, SOTSHIP1.SHIP_VIA_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.FRT_TERMS, SOTSHIP1.WHSE_CODE, SOTSHIP1.EDI_810_CREATED, SOTSHIP1.EDI_856_CREATED"
        ASCMAIN1.sql &= "   FROM SOTSHIP1, SOTORDR0"
        ASCMAIN1.sql &= "   WHERE ((NVL(SHIP_810_IND, '0') = '1' AND SHIP_810_BATCH_NO IS NULL)"
        ASCMAIN1.sql &= "   OR (NVL(SHIP_856_IND, '0') = '1' AND SHIP_856_BATCH_NO IS NULL))"
        ASCMAIN1.sql &= "   AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO "
        ASCMAIN1.sql &= "   AND SOTSHIP1.SHIP_STATUS = 'F'"
        ASCMAIN1.sql &= "   AND SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL"
        ASCMAIN1.sql &= "   UNION"
        ASCMAIN1.sql &= "    SELECT DISTINCT '1' SEL,  '0' EDI_STATUS, 'S' SEL_TYPE, SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_856_IND, "
        ASCMAIN1.sql &= "    SOTSHIP1.SHIP_810_IND, SOTSHIP1.SHIP_VIA_CODE, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.FRT_TERMS, SOTSHIP1.WHSE_CODE, SOTSHIP1.EDI_810_CREATED, SOTSHIP1.EDI_856_CREATED"
        ASCMAIN1.sql &= "   FROM SOTSHIP1, SOTORDR0  "
        ASCMAIN1.sql &= "   WHERE ((NVL(SHIP_810_IND, '0') = '1' AND SHIP_810_BATCH_NO IS NULL)"
        ASCMAIN1.sql &= "   OR (NVL(SHIP_856_IND, '0') = '1' AND SHIP_856_BATCH_NO IS NULL))"
        ASCMAIN1.sql &= "   AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO "
        ASCMAIN1.sql &= "   AND SOTSHIP1.SHIP_STATUS = 'F'"
        ASCMAIN1.sql &= "   AND  SOTSHIP1.BILL_OF_LADING_NO IS NULL"

        Fill_Records("SOTSHIPX", String.Empty, True, ASCMAIN1.sql)

        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select()
            Dim CUST_CODE As String = rowSOTSHIPX.Item("CUST_CODE") & String.Empty
            Dim rowEDTSLSP1 As DataRow = tblEDTSLSP1.Rows.Find(CUST_CODE)
            If rowEDTSLSP1 Is Nothing Then
                Continue For
            End If

            Dim Sql As String = String.Empty
            Sql = "(EDI_TP_QUAL = '" & rowEDTSLSP1.Item("EDI_QUAL_810") & "' and EDI_TP_ID = '" & rowEDTSLSP1.Item("EDI_ID_810") & "' and EDI_DOC_NO = '810')"
            Sql &= " or "
            Sql &= "(EDI_TP_QUAL = '" & rowEDTSLSP1.Item("EDI_QUAL_856") & "' and EDI_TP_ID = '" & rowEDTSLSP1.Item("EDI_ID_856") & "' and EDI_DOC_NO = '856')"

            For Each rowEDTTRPM1 As DataRow In tblEDTTRPM1.Select(Sql)
                Select Case rowEDTTRPM1.Item("EDI_DOC_NO") & String.Empty
                    Case "810"
                        If rowEDTTRPM1.Item("EDI_STATUS") & String.Empty = "T" _
                            AndAlso (rowSOTSHIPX.Item("SHIP_810_IND") & String.Empty = "1" OrElse rowSOTSHIPX.Item("CUST_FACTOR_TRANS_IND") & String.Empty = "1") Then
                            rowSOTSHIPX.Item("EDI_STATUS") = "1"
                        End If

                    Case "856"
                        If rowEDTTRPM1.Item("EDI_STATUS") & String.Empty = "T" _
                            AndAlso rowSOTSHIPX.Item("SHIP_856_IND") & String.Empty = "1" Then
                            rowSOTSHIPX.Item("EDI_STATUS") = "1"
                        End If
                End Select
            Next
        Next

        grdSOTSHIP1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdSOTSHIP1, "BILL_OF_LADING_NO")

        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        Try
            ASCMAIN1.Progress("Now generating")

            MyBase.EnforceConstraints(False)

            Dim clsEDC810O1 As New TAC.EDC810O1(dst.Tables("EDTSYSIH"), dst.Tables("EDT810O1"), dst.Tables("EDT810O2"), dst.Tables("EDT810O3"), dst.Tables("EDT810O5"))
            Dim clsEDC856O1 As New TAC.EDC856O1(dst.Tables("EDTSYSIH"), dst.Tables("EDT856O1"), dst.Tables("EDT856O2"), dst.Tables("EDT856O3"), dst.Tables("EDT856O4"), dst.Tables("EDT856O5"))

            Dim SHIP_810_BATCH_NO As String = String.Empty
            Dim SHIP_856_BATCH_NO As String = String.Empty
            Dim errorMessge As String = String.Empty
            Dim tblSOTSHIP1 As DataTable = Nothing

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL = '1'")
                ' Get a temp table of one or more SOTSHIP1 records to process
                Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("BILL_OF_LADING_NO")
                Select Case rowSOTSHIPX.Item("SEL_TYPE")
                    Case "B"
                        ' Convert Bill of lading to collection of Shipments
                        ' Some warehouses recycle the Bill of Lading Nos
                        tblSOTSHIP1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = :PARM1" _
                                                            & " AND SHIP_DATE_SHIPPED > SYSDATE - 90" _
                                                            & " AND (" _
                                                            & "     (NVL(SHIP_810_IND, '0') = '1' AND SHIP_810_BATCH_NO IS NULL)" _
                                                            & "  OR (NVL(SHIP_856_IND, '0') = '1' AND SHIP_856_BATCH_NO IS NULL)" _
                                                            & ")", "SOTSHIP1", "V", New Object() {SHIP_BOL_NO})
                        If tblSOTSHIP1.Rows.Count = 0 Then
                            ' Consolidated Invoices or Master Bill of Ladings may have processed these in a previous instance of the loop
                            Continue For
                        End If
                    Case "S"
                        tblSOTSHIP1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO = :PARM1", "SOTSHIP1", "V", New Object() {SHIP_BOL_NO})
                    Case Else
                        Throw New Exception("Could not create process entry for Ship Bol: " & SHIP_BOL_NO & ", SEL_TYPE is invalid (" & rowSOTSHIPX.Item("SEL_TYPE") & ")")
                End Select

                ' Process one or more SOTSHIP1 records
                ASCMAIN1.Progress("Processing Shipment " & SHIP_BOL_NO, "")

                If rowSOTSHIPX.Item("SHIP_810_IND") & String.Empty = "1" Then
                    ASCMAIN1.Progress("-", "EDT810O1")
                    For Each row As DataRow In tblSOTSHIP1.Select("ISNULL(SHIP_810_BATCH_NO, '') = ''")
                        ' May have been sent and we want to reprocess the 856s
                        If row.Item("EDI_810_CREATED") & String.Empty <> String.Empty Then
                            Continue For
                        End If

                        ' Never send an 810 if the 856 has not been transferred
                        If row.Item("EDI_856_CREATED") & String.Empty = String.Empty AndAlso row.Item("SHIP_856_IND") & String.Empty = "1" Then
                            Continue For
                        End If

                        SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
                        clsEDC810O1.Create810(SHIP_BOL_NO, SHIP_810_BATCH_NO)
                        ' if the invoices are $0.00 then SHIP_810_BATCH_NO will be empty - need to set to something
                        If SHIP_810_BATCH_NO = String.Empty Then
                            Throw New Exception("Could not create 810 entry for Ship Bol: " & SHIP_BOL_NO)
                        End If

                        Fill_Records("SOTSHIP1", SHIP_BOL_NO, False, String.Empty)
                        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                        rowSOTSHIP1.Item("EDI_810_CREATED") = DateTime.Now
                        rowSOTSHIP1.Item("SHIP_810_BATCH_NO") = SHIP_810_BATCH_NO
                        rowSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowSOTSHIP1.Item("LAST_DATE") = DateTime.Now
                    Next
                End If

                ' No longer send 856s from this screen. Must be sent from ADS information.
                ' as per Maria - 1/10/2014 && note: AndAlso 1 = 2 in If statement
                If rowSOTSHIPX.Item("SHIP_856_IND") & String.Empty = "1" AndAlso 1 = 2 Then
                    ASCMAIN1.Progress("-", "EDT856O1")
                    If tblSOTSHIP1.Select("EDI_856_CREATED is NULL").Length > 0 Then
                        ' Create one entry for the entire Bill of lading
                        SHIP_BOL_NO = tblSOTSHIP1.Rows(0).Item("SHIP_BOL_NO")
                        SHIP_856_BATCH_NO = clsEDC856O1.CreateEDI856(SHIP_BOL_NO, errorMessge)

                        If SHIP_856_BATCH_NO = String.Empty Then
                            Throw New Exception("Unable to process factor 856 for Ship Bol No: " & SHIP_BOL_NO & IIf(errorMessge.Length > 0, ", " & errorMessge, ""))
                        End If

                        If SHIP_856_BATCH_NO.Length > 0 Then
                            ' Stamp all the shipments for this Bill of Lading
                            For Each row As DataRow In tblSOTSHIP1.Select("")
                                SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
                                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                                If rowSOTSHIP1 Is Nothing Then
                                    Fill_Records("SOTSHIP1", SHIP_BOL_NO, False, String.Empty)
                                    rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                                End If
                                rowSOTSHIP1.Item("EDI_856_CREATED") = DateTime.Now
                                rowSOTSHIP1.Item("SHIP_856_BATCH_NO") = SHIP_856_BATCH_NO
                                rowSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                rowSOTSHIP1.Item("LAST_DATE") = DateTime.Now
                            Next
                        End If
                    End If
                End If

                Update_Record()
            Next
            MyBase.EnforceConstraints(True)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            MessageBox.Show("Process Complete", "Send Edi", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()

            For Each table As String In New String() {"SOTSHIP1", "EDTSYSIH", "EDT810O1", "EDT810O2", "EDT810O3", "EDT810O5", _
                                                 "EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5"}
                Update_Record_TDA(table)
            Next

            MyBase.CommitTrans()

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIP1, "BBPBB", "Select All", "De-select All", "Select All Customer", "De-select All Customer")

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

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTPICK1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

            Case "Select All", "De-select All"
                For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("")
                    rowSOTSHIPX.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next

            Case "Select All Customer", "De-select All Customer"
                For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("CUST_CODE = '" & grdSOTSHIP1.ActiveRow.Cells("CUST_CODE").Text & "'")
                    rowSOTSHIPX.Item("SEL") = IIf(e.Tool.Key.StartsWith("Select All"), "1", "0")
                Next


        End Select
    End Sub

#End Region

    Private Sub grdSOTSHIP1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIP1.InitializeRow
        If e.Row.Cells("SHIP_810_IND").Value & String.Empty = "1" Then
            If IsDate(e.Row.Cells("EDI_810_CREATED").Value & String.Empty) Then
                e.Row.Cells("SHIP_810_IND").Appearance.BackColor = Drawing.Color.Red
            Else
                e.Row.Cells("SHIP_810_IND").Appearance.BackColor = Drawing.Color.Green
            End If
        End If

        If e.Row.Cells("SHIP_856_IND").Value & String.Empty = "1" Then
            If IsDate(e.Row.Cells("EDI_856_CREATED").Value & String.Empty) Then
                e.Row.Cells("SHIP_856_IND").Appearance.BackColor = Drawing.Color.Red
            Else
                e.Row.Cells("SHIP_856_IND").Appearance.BackColor = Drawing.Color.Green
            End If
        End If

        'If e.Row.Cells("CUST_FACTOR_TRANS_IND").Value & String.Empty = "1" Then
        '    If IsDate(e.Row.Cells("FACTOR_TRANS_LAST_DATE").Value & String.Empty) Then
        '        e.Row.Cells("CUST_FACTOR_TRANS_IND").Appearance.BackColor = Drawing.Color.Red
        '    Else
        '        e.Row.Cells("CUST_FACTOR_TRANS_IND").Appearance.BackColor = Drawing.Color.Green
        '    End If
        'End If
    End Sub
End Class