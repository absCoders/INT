Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class APFINVHR

    Public updated As Boolean = False
    Public frmASFBASE0 As ASFBASE0
    Public VEND_CODE_ACC As String = ""
    Public VOUCHER_NO As String = ""
    Public INV_NUM As String = ""
    Public INV_DATE As Date

    Public CTL_NOs As New List(Of String)

    'Public Sub New(
    '              ByVal frmASFBASE0_in As ASFBASE0,
    '              ByVal ACCRUAL_CODE As String,
    '              ByVal VEND_CODE_ACC As String,
    '              ByVal CTL_NOTE As String,
    '              ByVal COST_ACT As Decimal)
    Public Sub New(
                  ByVal frmASFBASE0_in As ASFBASE0, VEND_CODE_ACC_in As String, VOUCHER_NO_in As String, INV_NUM_in As String, INV_DATE_in As Date)
        frmASFBASE0 = frmASFBASE0_in
        VEND_CODE_ACC = VEND_CODE_ACC_in
        VOUCHER_NO = VOUCHER_NO_in
        INV_NUM = INV_NUM_in
        INV_DATE = INV_DATE_in

        InitializeComponent()
    End Sub

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
                & ", ICTIREC2.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COST_CATGY_CODE, ICTIREC2.PO_COST, ICTIREC2.QTY_REC" & vbCrLf _
                & ", ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER" & vbCrLf _
                & " from ICTIREC1, ICTIREC2, ICTITEM1" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & "   and ICTIREC1.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "V", 2)
            .Tables("ICTIRECX").Columns.Add("SEL", GetType(System.String))
            .Tables("ICTIRECX").Columns("SEL").DefaultValue = "0"
            .Tables("ICTIRECX").Columns.Add("PO_COST_EXT", GetType(System.Decimal), "ISNULL(PO_COST,0) * ISNULL(QTY_REC,0)")
            .Tables("ICTIRECX").Columns.Add("ACT_AMT", GetType(System.Decimal))


            ASCMAIN1.sql = "Select ICTIREC2.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COST_CATGY_CODE" & vbCrLf _
                & ", SUM (ICTIREC2.QTY_REC) QTY_REC, SUM (ICTIREC2.PO_COST * ICTIREC2.QTY_REC) PO_COST_EXT" & vbCrLf _
                & " from ICTIREC1,ICTIREC2,ICTITEM1" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & "   and ICTIREC1.PO_ORDER_NO = :PARM1" & vbCrLf _
                & " group by ICTIREC2.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COST_CATGY_CODE"
            Create_TDA(.Tables.Add, "ICTIRECI", "**", 0, False, "V", 1)
            .Tables("ICTIRECI").Columns.Add("PO_COST", GetType(System.Decimal), "IIF(ISNULL(QTY_REC,0) = 0, 0, ISNULL(PO_COST_EXT,0) / ISNULL(QTY_REC,0))")
            .Tables("ICTIRECI").Columns.Add("ACT_AMT", GetType(System.Decimal))
        End With

        grdICTIRECX.DataSource = dst.Tables("ICTIRECX")
        grdICTIRECI.DataSource = dst.Tables("ICTIRECI")

        Create_Summary(grdICTIRECX, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECX, New String() {"QTY_REC", "PO_COST_EXT", "ACT_AMT"})
        Create_Summary(grdICTIRECI, "ITEM_CODE", "Count")
        Create_Summary(grdICTIRECI, New String() {"QTY_REC", "PO_COST_EXT", "ACT_AMT"})

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIRECX, grdICTIRECI}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        Next

        With grdICTIRECX.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "PO_ORDER_NO" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                ElseIf New String() {"SEL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    'gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        End With

        With grdICTIRECI.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "PO_COST_EXT" Or gcol.Key = "PO_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf New String() {"QTY_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf New String() {"ACT_AMT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Color.LightPink
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                    'gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        End With

        ' Sort_grdColumns(grdICTIRECX, "PO_ORDER_NO, PO_ORDER_LNO, RECEIPT_NO, RECEIPT_LNO")

        'ASCMAIN1.Add_Value_List(grdICTIRECX, "WO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        'ASCMAIN1.Add_Value_List(grdICTIRECI, "WO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})

        Absx1.txtFor("VEND_CODE_ACC").Text = VEND_CODE_ACC
        optProRate.Value = "I"

    End Sub


#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTIRECX, "SBB", "Show Filter", "Select All", "De-Select All")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        Try
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        Catch ex As Exception
            e.Cancel = True
            Exit Sub
        End Try

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        'If tlb_pop.Tools.Exists("Clone Line (w/Adj)") Then
        '    tlb_btn = DirectCast(tlb_pop.Tools("Clone Line (w/Adj)"), UltraWinToolbars.ButtonTool)
        '    tlb_btn.SharedProps.Visible = Not InquiryMode
        'End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdAPTSUBM1"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Re-Submit Submitted Email"), UltraWinToolbars.ButtonTool)
                    'If optShow.Value = "P" Then
                    '    tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
                    'Else
                    '    tlb_btn.SharedProps.Visible = False
                    'End If
                    'tlb_btn = DirectCast(tlb_pop.Tools("Delete Submitted Email"), UltraWinToolbars.ButtonTool)
                    'If optShow.Value = "U" Then
                    '    tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
                    'Else
                    '    tlb_btn.SharedProps.Visible = False
                    'End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PO_ORDER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Prepare_ICTIRECX(Absx1.txtFor("PO_ORDER_NO").Text)
                End If

            Case "COST_ACT"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Calculate_ProRata()
                End If

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PO_ORDER_NO"
                Prepare_ICTIRECX(Absx1.txtFor("PO_ORDER_NO").Text)
        End Select
    End Sub

    Overrides Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "COST_ACT"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Calculate_ProRata()
                End If

        End Select
    End Sub

#End Region

    Sub Prepare_ICTIRECX(PO_ORDER_NO As String)

    End Sub

    Private Sub grdICTIRECX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTIRECX.AfterRowActivate

    End Sub


    Private Sub grdICTIRECX_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIRECX.BeforeRowUpdate

    End Sub

    Private Sub grdICTIRECX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIRECX.AfterCellUpdate

    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdAddPO_Click(sender As Object, e As EventArgs) Handles cmdAddPO.Click

        Dim VEND_CODE_PO As String = Absx1.txtFor("VEND_CODE").Text
        If VEND_CODE_PO = "" Then
            MsgBox("Please specify a Vendor Code", MsgBoxStyle.OkOnly, "Cannot Find POs Received in the past year without a Vendor")
            Exit Sub

        End If
        Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)
        ASCMAIN1.sql = "Select POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_NO, POTORDR1.PO_STATUS" & vbCrLf _
            & " , COUNT (DISTINCT ICTIREC1.RECEIPT_NO) RECEIPTS" & vbCrLf _
            & " , MAX (ICTIREC1.RECEIPT_DATE) LAST_REC" & vbCrLf _
            & " from ICTIREC1, POTORDR1" & vbCrLf _
            & $" where ICTIREC1.VEND_CODE = '{Absx1.txtFor("VEND_CODE").Text}' and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO And ICTIREC1.OPS_YYYYPP >= '{YP}'" & vbCrLf _
            & " group by POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_NO, POTORDR1.PO_STATUS"
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        Using frm As New ASFMSGBF
            frm.Show_grd(TBL, frmASFBASE0, "Select the PO for which to load Receipts to create Accruals", "ROW")
            If frm.user_option = -1 Then
                ' USER CLICKED CANCEL
            Else
                Dim G As UltraWinGrid.UltraGridRow = frm.grow
                Dim PO_ORDER_NO = G.Cells("PO_ORDER_NO").Value & ""
                Fill_Records("ICTIRECX", PO_ORDER_NO)
                'Fill_Records("ICTIRECI", PO_ORDER_NO)
            End If
        End Using
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special(
     ByVal ctl As Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME


            Case "PO_ORDER_NO"
                Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)
                ASCMAIN1.sql = "Select POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_NO, POTORDR1.PO_STATUS" & vbCrLf _
                    & " , COUNT (DISTINCT ICTIREC1.RECEIPT_NO) RECEIPTS" & vbCrLf _
                    & " , MAX (ICTIREC1.RECEIPT_DATE) LAST_REC" & vbCrLf _
                    & " from ICTIREC1, POTORDR1" & vbCrLf _
                    & $" where ICTIREC1.VEND_CODE = '{Absx1.txtFor("VEND_CODE").Text}' and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO And ICTIREC1.OPS_YYYYPP >= '{YP}'" & vbCrLf _
                    & " group by POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_NO, POTORDR1.PO_STATUS"

                sql_where = $"PO_ORDER_NO in (Select PO_ORDER_NO from ({ASCMAIN1.sql}))"


            Case "VEND_CODE"
                sql_where = "VEND_CODE in (Select DISTINCT VEND_CODE from POTORDR1 where PO_DATE_ORDERED > SYSDATE - 100)"

        End Select
    End Sub

    Private Sub optProRate_ValueChanged(sender As Object, e As EventArgs) Handles optProRate.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Prorate_Option()
        Calculate_ProRata()
    End Sub

    Sub Set_Prorate_Option()

        SplitContainer3.Panel2Collapsed = (optProRate.Value <> "I")

        With grdICTIRECX.DisplayLayout.Bands(0)
            .Columns("ACT_AMT").Hidden = (optProRate.Value = "I")
        End With

    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click

        EMsg = ""

        Dim ACCRUAL_CODE As String = Absx1.txtFor("ACCRUAL_CODE").Text
        If ACCRUAL_CODE = "" Then
            EMsg &= vbCr & "No Value Specified for Accrual Code"
        Else
            Dim rowAPTACRM1 As DataRow = LookUp("APTACRM1", ACCRUAL_CODE)
            If rowAPTACRM1 Is Nothing Then
                EMsg &= vbCr & "Invalid Value Specified for Accrual Code"
            End If
        End If

        Dim ACT_AMT As Decimal = 0
        If optProRate.Value = "I" Then
            ACT_AMT = Val(dst.Tables("ICTIRECI").Compute("SUM(ACT_AMT)", "") & "")
        Else
            ACT_AMT = Val(dst.Tables("ICTIRECX").Compute("SUM(ACT_AMT)", "") & "")
        End If

        Dim COST_ACT As Decimal = Val(numCOST_ACT.Value & "")

        If COST_ACT = 0 Then
            EMsg &= vbCr & $"Actual Amount cannot be $0.00"
        End If

        If COST_ACT <> ACT_AMT Then
            EMsg &= vbCr & $"Actual Amount {Format(COST_ACT, "#,##0.00")} does not agree with the Sum of Details {Format(ACT_AMT, "#,##0.00")}"
        End If


        If dst.Tables("ICTIRECX").Select("SEL = '1' and ISNULL(PO_COST_EXT,0) = 0").Length > 0 Then
            EMsg &= vbCr & $"Cannot Prorate a Receipt Line with $0 PO Ext Cost"
        End If
        If dst.Tables("ICTIRECI").Select("ISNULL(ACT_AMT,0) <> 0 and ISNULL(PO_COST_EXT,0) = 0").Length > 0 Then
            EMsg &= vbCr & $"Cannot Prorate an Item with $0 PO Ext Cost"
        End If


        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Create Accrual")
            Exit Sub
        End If

        If optProRate.Value = "I" Then

            For Each rowICTIRECI As DataRow In dst.Tables("ICTIRECI").Select("ISNULL(ACT_AMT,0) <> 0")
                Dim ITEM_CODE As String = rowICTIRECI.Item("ITEM_CODE")
                Dim ACT_AMT_item As Decimal = Val(rowICTIRECI.Item("ACT_AMT") & "")
                Dim PO_COST_EXT_item As Decimal = Val(rowICTIRECI.Item("PO_COST_EXT") & "")
                For Each rowICTIRECX As DataRow In dst.Tables("ICTIRECX").Select($"ISNULL(SEL,'0') = '1' and ITEM_CODE = '{ITEM_CODE}'")
                    Dim PO_COST_EXT_line As Decimal = Val(rowICTIRECX.Item("PO_COST_EXT") & "")
                    rowICTIRECX.Item("ACT_AMT") = ACT_AMT_item * PO_COST_EXT_line / PO_COST_EXT_item
                    Create_Accrual(rowICTIRECX)
                Next
            Next
        Else
            For Each rowICTIRECX As DataRow In dst.Tables("ICTIRECX").Select("SEL = '1'")
                Create_Accrual(rowICTIRECX)
            Next
        End If

        updated = True
        Me.Close()

    End Sub

    Function Create_Accrual(ROW As DataRow)

        Dim rowAPTACRC1 As DataRow = frmASFBASE0.dst.Tables("APTACRC1").NewRow

        Dim CTL_NO As String = ASCMAIN1.Next_Control_No("APTACRC1.CTL_NO")
        CTL_NOs.Add(CTL_NO)

        With rowAPTACRC1
            .Item("CTL_NO") = CTL_NO
            .Item("VEND_CODE_ACC") = Absx1.txtFor("VEND_CODE_ACC").Text
            .Item("ACCRUAL_CODE") = Absx1.txtFor("ACCRUAL_CODE").Text
            .Item("CHARGEBACK_IND") = "0"
            .Item("CTL_DATE") = INV_DATE
            .Item("CTL_NOTE") = Absx1.txtFor("CTL_NOTE").Text
            .Item("VOUCHER_NO_ORIG") = VOUCHER_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("CTL_STATUS") = "0"
            .Item("INV_PRINT_IND") = "0"

            .Item("PO_ORDER_NO") = ROW.Item("PO_ORDER_NO")
            .Item("PO_ORDER_LNO") = ROW.Item("PO_ORDER_LNO")
            .Item("RECEIPT_NO") = ROW.Item("RECEIPT_NO")
            .Item("RECEIPT_LNO") = ROW.Item("RECEIPT_LNO")
            .Item("ITEM_CODE") = ROW.Item("ITEM_CODE")

            'Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", .Item("RECEIPT_NO"))
            'Dim rowICTIREC2 As DataRow = LookUp("ICTIREC2", New String() { .Item("RECEIPT_NO"), .Item("RECEIPT_LNO")})

            .Item("COST_CATGY_CODE") = ROW.Item("COST_CATGY_CODE")
            .Item("SOURCE_DOC_NO") = Absx1.txtFor("SOURCE_DOC_NO").Text

            ' *** SOME OF THESE FIELDS SHOULD BE SET AS SHOWN BELOW ONLY IF THIS IS A PREPAYMENT
            .Item("COST_ACT") = ROW.Item("ACT_AMT")
            .Item("COST_ORIG") = ROW.Item("ACT_AMT")
            .Item("CTL_TYPE") = "A" ' Manual
            .Item("PPD_IND") = "0" ' Pre-Paid
        End With
        frmASFBASE0.dst.Tables("APTACRC1").Rows.Add(rowAPTACRC1)

        'Add_APTINVH7(CTL_NO, True)
        'Clear_Other_Accrual_Controls()

        Return rowAPTACRC1

    End Function

    Private Sub grdICTIRECX_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICTIRECX.AfterRowUpdate
        Calculate_ProRata
    End Sub

    Sub Calculate_ProRata()


        If optProRate.Value = "I" Then
            Dim ITEM_CODEs As New Dictionary(Of String, Decimal)
            For Each row As DataRow In dst.Tables("ICTIRECI").Select("")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim ACT_AMT As Decimal = Val(row.Item("ACT_AMT") & "")
                If ACT_AMT <> 0 Then ITEM_CODEs.Add(ITEM_CODE, ACT_AMT)
            Next

            dst.Tables("ICTIRECI").Rows.Clear()
            For Each row As DataRow In dst.Tables("ICTIRECX").Select("SEL = '1'")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim QTY_REC As Int32 = Val(row.Item("QTY_REC") & "")
                Dim PO_COST As Decimal = Val(row.Item("PO_COST") & "")
                Dim PO_COST_EXT As Decimal = QTY_REC * PO_COST
                Dim ITEM_DESC As String = row.Item("ITEM_DESC")
                Dim rowICTIRECI As DataRow = dst.Tables("ICTIRECI").Rows.Find(ITEM_CODE)
                If rowICTIRECI Is Nothing Then
                    rowICTIRECI = dst.Tables("ICTIRECI").NewRow
                    With rowICTIRECI
                        .Item("ITEM_CODE") = ITEM_CODE
                        .Item("ITEM_DESC") = ITEM_DESC
                        .Item("QTY_REC") = QTY_REC
                        .Item("PO_COST_EXT") = PO_COST_EXT

                        If ITEM_CODEs.ContainsKey(ITEM_CODE) Then
                            .Item("ACT_AMT") = ITEM_CODEs(ITEM_CODE)
                        End If

                    End With
                    dst.Tables("ICTIRECI").Rows.Add(rowICTIRECI)
                Else
                    rowICTIRECI.Item("QTY_REC") = Val(rowICTIRECI.Item("QTY_REC") & "") + QTY_REC
                    rowICTIRECI.Item("PO_COST_EXT") = Val(rowICTIRECI.Item("PO_COST_EXT") & "") + PO_COST_EXT
                End If

            Next

        Else
            Dim PO_COST_EXT As Decimal = Val(dst.Tables("ICTIRECX").Compute("SUM(PO_COST_EXT)", "SEL = '1'") & "")
            Dim COST_ACT As Decimal = Val(numCOST_ACT.Value & "")

            For Each row As DataRow In dst.Tables("ICTIRECX").Select("")
                Dim PO_COST_EXT_LINE As Decimal = Val(row.Item("PO_COST_EXT") & "")
                Dim SEL As String = row.Item("SEL") & ""
                If SEL = "1" Then
                    Dim ACT_AMT As Decimal = IIf(PO_COST_EXT = 0, 0, System.Math.Round(PO_COST_EXT_LINE * COST_ACT / PO_COST_EXT, 2))
                    row.Item("ACT_AMT") = ACT_AMT
                Else
                    row.Item("ACT_AMT") = DBNull.Value
                End If

            Next

        End If

    End Sub
End Class