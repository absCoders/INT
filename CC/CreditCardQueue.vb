Public Class CreditCardQueue

    Private sCustomerCode As String = String.Empty
    Private initialized As Boolean = False

    Private rowSOTPARM1 As DataRow = Nothing
    Private rowARTCUST1 As DataRow = Nothing
    Private sAllowEdit As Boolean = False

    Private AbsCon As Object = Nothing ' New ABSConnector
    Private objCCProcessor As Object ' New TAC.ARCCCARD(TAC.ARCCCARD.ProcessingTypes.FDMS)
    Private allowACHEntries As Boolean = False
    Private sAllowAutoAuthForm As Boolean = False

    Private InitializingControl As Boolean = True
    
    Private Enum Encryption
        Decrypt
        Encrypt
    End Enum

    Public Sub New()

        InitializingControl = True

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sCustomerCode = String.Empty
        initialized = False
        AllowEdit = False
        AllowAutoAuthForm = False

        InitializingControl = False
    End Sub

#Region "Control Properties"

    ''' <summary>
    ''' Customer Code to get CC data for
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomerCode() As String
        Get
            Return sCustomerCode
        End Get
        Set(ByVal value As String)
            sCustomerCode = value
            grpHeader.Text = Space(20) & "Accounts for " & sCustomerCode
        End Set
    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGrid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdCC
        End Get

    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGridACH() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdACH
        End Get

    End Property

    ''' <summary>
    ''' Sets / Gets if the control allows editting
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllowEdit() As Boolean
        Get
            Return sAllowEdit
        End Get
        Set(ByVal value As Boolean)
            sAllowEdit = value
            chkCCEdit.Visible = sAllowEdit
            SplitContainer25.Panel2Collapsed = Not sAllowEdit
        End Set
    End Property

    ''' <summary>
    ''' Get whether in Edit Mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property InEditMode() As Boolean
        Get
            Return chkCCEdit.Checked
        End Get
    End Property

    ''' <summary>
    ''' Get / Set Allow Auto Authorization Form
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllowAutoAuthForm() As Boolean
        Get
            Return sAllowAutoAuthForm
        End Get
        Set(ByVal value As Boolean)
            sAllowAutoAuthForm = value
        End Set
    End Property

    Public Sub SetReadOnly()
        If InitializingControl Then Exit Sub

        chkCCEdit.Checked = False
        ToggleCCEditable()
    End Sub

    ''' <summary>
    ''' Fires if the Update is Successful. Passes in the Customer Master Row with modified data.
    ''' </summary>
    ''' <param name="rowARTCUST1"></param>
    ''' <remarks></remarks>
    Public Event UpdateClickEvent(ByRef rowARTCUST1 As DataRow)

    ''' <summary>
    ''' Fires when the user clicks the Cancel Button
    ''' </summary>
    ''' <remarks></remarks>
    Public Event CancelClickEvent()

#End Region

#Region "Control Public Methods"

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <param name="displayCustomerCode">Customer Code used to fill query</param>
    ''' <remarks></remarks>
    Public Sub DisplayData(ByVal displayCustomerCode As String, Optional ByVal unCheckEdit As Boolean = True)
        If InitializingControl Then Exit Sub
        Try

            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            sCustomerCode = displayCustomerCode
            If unCheckEdit Then
                chkCCEdit.Checked = False
            End If
            ToggleCCEditable()

            With AbsCon.dst
                If Not AbsCon.dst.Tables.Contains("ARTCUSTC") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUSTC", "*", 1)
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCCPA1") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCCPA1", "*")
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCUST1") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUST1", "*")
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCUSPA") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUSPA", "*")
                    .Tables("ARTCUSPA").Columns.Add("ACH_ACCT_NO_LAST4", GetType(System.String), "IIF(LEN(ACH_ACCT_NO) <= 4, ACH_ACCT_NO, SUBSTRING(ACH_ACCT_NO, LEN(ACH_ACCT_NO) - 3, 4))")
                End If

            End With

            AbsCon.Fill_Records("ARTCUSTC", displayCustomerCode)
            AbsCon.Fill_Records("ARTCUST1", displayCustomerCode)
            AbsCon.Fill_Records("ARTCUSPA", String.Empty, True, "SELECT * FROM ARTCUSPA WHERE CUST_CODE = '" & displayCustomerCode & "' AND ACH_ACCT_STATUS = 'A'")
            EncryptDecryptArtpymtw("ARTCUSPA", Encryption.Decrypt)
            AbsCon.dst.Tables("ARTCCPA1").Rows.Clear()

            ' Fill in Control Values
            If AbsCon.dst.Tables("ARTCUST1").Rows.count > 0 Then
                rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").Rows(0)
                grpHeader.Text = Space(20) & "Accounts for " & rowARTCUST1.Item("CUST_CODE") & " " & rowARTCUST1.Item("CUST_NAME")
            Else
                rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").newrow
                rowARTCUST1.Item("CUST_CODE") = displayCustomerCode
                grpHeader.Text = Space(20) & "Accounts for " & rowARTCUST1.Item("CUST_CODE") & " " & rowARTCUST1.Item("CUST_NAME")
            End If

            grdCC.DataSource = AbsCon.dst.Tables("ARTCUSTC")
            grdACH.DataSource = AbsCon.dst.Tables("ARTCUSPA")

            ' Initialization is done here; otherwise, the control crashes if the code in in 'New' when it looks for an Oralce Connection
            If Not initialized Then
                txtCUST_AUTO_CCPA_NOTE.MaxLength = AbsCon.dst.Tables("ARTCUST1").Columns("CUST_AUTO_CCPA_NOTE").MaxLength
                cbeCUST_AUTO_CC_OPER.DataSource = AbsCon.GetDataTable("SELECT USER_ID, USER_NAME FROM ASTUSER1")
                Dim CC_PROC_TYPE As String = ""
                Dim CC_PROC_FOLDER As String = ""
                rowSOTPARM1 = AbsCon.GetDataRow("Select * From SOTPARM1 Where SO_PARM_KEY = 'Z'")
                Dim rowARTCCPRC As DataRow = AbsCon.Lookup("ARTCCPRC", rowSOTPARM1.Item("SO_PARM_CC_PROC_CODE") & String.Empty)
                If rowARTCCPRC IsNot Nothing Then
                    CC_PROC_TYPE = rowARTCCPRC.Item("CC_PROC_TYPE") & ""
                    CC_PROC_FOLDER = rowARTCCPRC.Item("CC_PROC_FOLDER") & ""
                End If

                Select Case CC_PROC_TYPE

                    Case "F"
                        objCCProcessor = New TAC.ARCCCARD(CC_PROC_FOLDER, TAC.ARCCCARD.ProcessingTypes.FDMS)
                    Case "P"
                        objCCProcessor = New TAC.ARCCCARD(CC_PROC_FOLDER, TAC.ARCCCARD.ProcessingTypes.Paymentech)
                    Case "A"
                        objCCProcessor = New TAC.ARCCCARD(CC_PROC_FOLDER, TAC.ARCCCARD.ProcessingTypes.AuthorizeNet, GetType(ABSolution.ASCSCRTY))
                    Case Else
                        objCCProcessor = Nothing
                End Select

                AbsCon.Add_Value_List(grdCC, "CUST_CREDIT_CARD_TYPE")

                With grdCC.DisplayLayout.Bands(0)
                    .Columns("CUST_CREDIT_CARD_NO").Header.Fixed = True
                    .Columns("CUST_CREDIT_CARD_LAST4").Header.Fixed = True
                    .Columns("CUST_CREDIT_CARD_TYPE").Header.Fixed = True
                End With

                Dim vl As Infragistics.Win.ValueList

                If (Not grdACH.DisplayLayout.ValueLists.Exists("ACH_ACCT_TYPE_ID")) Then
                    vl = grdACH.DisplayLayout.ValueLists.Add("ACH_ACCT_TYPE_ID")
                    vl.ValueListItems.Add("0", "Unknown")
                    vl.ValueListItems.Add("1", "Business")
                    vl.ValueListItems.Add("2", "Personal")

                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_TYPE_ID").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_TYPE_ID").ValueList = grdACH.DisplayLayout.ValueLists("ACH_ACCT_TYPE_ID")
                End If


                If (Not grdACH.DisplayLayout.ValueLists.Exists("ACH_ACCT_STATUS")) Then
                    vl = grdACH.DisplayLayout.ValueLists.Add("ACH_ACCT_STATUS")
                    vl.ValueListItems.Add("", "Unknown")
                    vl.ValueListItems.Add("A", "Active")
                    vl.ValueListItems.Add("X", "Deleted")

                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_STATUS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_STATUS").ValueList = grdACH.DisplayLayout.ValueLists("ACH_ACCT_STATUS")

                End If

                With grdACH.DisplayLayout.Bands(0)
                    .Columns("ACH_ACCT_NAME").Header.Fixed = True
                End With

                initialized = True
            End If

            SetupCustomerMasterData()
            chkCUST_AUTO_CC_AUTH.Enabled = sAllowAutoAuthForm

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
        End Try

    End Sub

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DisplayData()
        If InitializingControl Then Exit Sub
        DisplayData(sCustomerCode)
    End Sub

    Public Sub ClearData()
        If InitializingControl Then Exit Sub
        Try

            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            If chkCCEdit.Checked Then
                cmdCCCancel_Click(Nothing, Nothing)
            End If

            If AbsCon IsNot Nothing AndAlso AbsCon.dst.Tables.Count > 0 Then
                AbsCon.dst.Tables("ARTCUSTC").Rows.Clear()
                AbsCon.dst.Tables("ARTCUSPA").Rows.Clear()
                AbsCon.dst.Tables("ARTCCPA1").Rows.Clear()
                AbsCon.dst.Tables("ARTCUST1").Rows.Clear()
                optCUST_AUTO_CCPA.Value = "0"
            End If

            txtCUST_AUTO_CCPA_NOTE.Clear()
            dteCUST_AUTO_CC_AUTH_DATE.Value = String.Empty
            cbeCUST_AUTO_CC_OPER.Text = String.Empty
            chkCUST_AUTO_CC_AUTH.Checked = False
            chkCUST_AUTOQ_WEB.Checked = False
            chkCCEdit.Checked = False
            grpHeader.Text = "."

            ToggleCCEditable()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub optCUST_AUTO_CCPA_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCUST_AUTO_CCPA.ValueChanged

        If InitializingControl Then Exit Sub

        Select Case optCUST_AUTO_CCPA.Value

            Case "0", "1", "2"
                grdCC.Visible = True
                grdACH.Visible = False

            Case "3"
                grdACH.Visible = True
                grdCC.Visible = False

        End Select
    End Sub

#End Region

#Region "Grid Control"

    Private Sub grdCC_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdCC.AfterCellUpdate
        If InitializingControl Then Exit Sub
        Select Case e.Cell.Column.Key

            Case "CUST_CREDIT_CARD_PREFERRED"
                If Not grdCC.ActiveRow.IsDataRow Then
                    grdCC.UpdateData()
                End If

            Case "CUST_CREDIT_CARD_NO"
                Dim CUST_CREDIT_CARD_NO As String = e.Cell.Text
                If Len(CUST_CREDIT_CARD_NO) >= 15 Then
                    e.Cell.Row.Cells("CUST_CREDIT_CARD_LAST4").Value = Mid(CUST_CREDIT_CARD_NO, Len(CUST_CREDIT_CARD_NO) - 3, 4)
                Else
                    e.Cell.Row.Cells("CUST_CREDIT_CARD_LAST4").Value = ""
                End If

                objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                Try
                    objCCProcessor.ValidateCard()

                Catch ex As Exception

                End Try

                Dim CUST_CREDIT_CARD_TYPE As String = ""

                Select Case objCCProcessor.CreditCardType
                    Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                        CUST_CREDIT_CARD_TYPE = "AMEX"
                    Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                        CUST_CREDIT_CARD_TYPE = "MSTR"
                    Case TAC.ARCCCARD.CreditCardTypes.vctVisa
                        CUST_CREDIT_CARD_TYPE = "VISA"
                    Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                        CUST_CREDIT_CARD_TYPE = "DISC"
                    Case Else
                        CUST_CREDIT_CARD_TYPE = ""
                End Select

                Try
                    If CUST_CREDIT_CARD_TYPE.Length > 0 Then
                        e.Cell.Row.Cells("CUST_CREDIT_CARD_TYPE").Value = (AbsCon.Get_Image(AbsCon.Folders("Images") & "\ABS\CC\", CUST_CREDIT_CARD_TYPE & ".GIF"))
                    End If
                Catch ex As Exception

                End Try


        End Select
    End Sub

    Private Sub grdCC_AfterRowInsert(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdCC.AfterRowInsert
        If InitializingControl Then Exit Sub
        e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value = "A"
        e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value = "US"
    End Sub

    Private Sub grdCC_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdCC.AfterRowUpdate
        If InitializingControl Then Exit Sub
        If e.Row.Cells("CUST_CREDIT_CARD_PREFERRED").Value & "" = "1" Then
            Dim CUST_CREDIT_CARD_NO As String = e.Row.Cells("CUST_CREDIT_CARD_NO").Text
            For Each row As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO <> '" & CUST_CREDIT_CARD_NO & "'", "")
                row.Item("CUST_CREDIT_CARD_PREFERRED") = ""
            Next
        End If
    End Sub

    Private Sub grdCC_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdCC.BeforeRowUpdate
        If InitializingControl Then Exit Sub
        grdCC.RowUpdateCancelAction = Infragistics.Win.UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation

        If CustomerCode.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Cannot Determine the Customer. Please contact ABS.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        With grdCC

            e.Row.Cells("CUST_CODE").Value = CustomerCode
            e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value = (e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value & String.Empty).ToString.Trim.ToUpper
            e.Row.Cells("CUST_CREDIT_CARD_STATE").Value = (e.Row.Cells("CUST_CREDIT_CARD_STATE").Value & String.Empty).ToString.Trim.ToUpper

            'CUST_CREDIT_CARD_STATUS - set to active if no value
            If (e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value & String.Empty).ToString.Trim.Length = 0 Then
                e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value = "A"
            End If

            Dim CUST_CREDIT_CARD_KEY As String = e.Row.Cells("CUST_CREDIT_CARD_KEY").Value & String.Empty
            If CUST_CREDIT_CARD_KEY.Length = 0 Then
                CUST_CREDIT_CARD_KEY = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSTC.CUST_CREDIT_CARD_KEY")
                e.Row.Cells("CUST_CREDIT_CARD_KEY").Value = CUST_CREDIT_CARD_KEY
            End If

            Dim CUST_CREDIT_CARD_NO As String = e.Row.Cells("CUST_CREDIT_CARD_NO").Text
            Dim CUST_CREDIT_CARD_EXP_DATE As String = e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Text
            e.Row.Cells("CUST_CREDIT_CARD_NO").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance.ForeColor = Drawing.Color.Empty
            Try
                objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                objCCProcessor.CustomerCreditCard.CardExpMonth = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 1, 2) & "")
                objCCProcessor.CustomerCreditCard.CardExpYear = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 3, 2) & "")
                objCCProcessor.ValidateCard()
            Catch ex As Exception
                e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("CUST_CREDIT_CARD_NO").Appearance.ForeColor = Drawing.Color.Red
            End Try

        End With
    End Sub

    Private Sub grdCC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdCC.ClickCellButton
        If InitializingControl Then Exit Sub
        Dim sql_where As String = ""
        Select Case grdCC.ActiveCell.Column.Key
            Case "CUST_CREDIT_CARD_COUNTRY"
                AbsCon.grdClickCellButton(grdCC, sql_where, False, "", "COUNTRY_CODE")
                ' sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "' or LENS_DESIGNER_CODE is Null"
        End Select

    End Sub

    Private Sub grdCC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdCC.InitializeRow
        If InitializingControl Then Exit Sub
        If e.Row.IsAddRow Then
        Else
            Dim row As DataRow = AbsCon.dst.Tables("ARTCUSTC").Rows.Find(New Object() {e.Row.Cells("CUST_CODE").Value & "", e.Row.Cells("CUST_CREDIT_CARD_NO").Value & ""})
            If row IsNot Nothing AndAlso row.RowState = DataRowState.Added Then
                e.Row.Cells("CUST_CREDIT_CARD_NO").Hidden = False
            Else
                e.Row.Cells("CUST_CREDIT_CARD_NO").Hidden = True
            End If
        End If

        If e.Row Is Nothing OrElse Not e.Row.IsDataRow Then
            Exit Sub
        End If

        If objCCProcessor Is Nothing Then
            Exit Sub
        End If

        'Dim CUST_CREDIT_CARD_TYPE As String = e.Row.Cells("CUST_CREDIT_CARD_TYPE").Text

        'If CUST_CREDIT_CARD_TYPE <> "" Then
        '    e.Row.Cells("CUST_CREDIT_CARD_TYPE").Appearance.Image = AbsCon.Get_Image(AbsCon.Folders("Images") & "\ABS\CC\", CUST_CREDIT_CARD_TYPE & ".GIF")
        'End If

        Try
            objCCProcessor.CustomerCreditCard.CardNumber = e.Row.Cells("CUST_CREDIT_CARD_NO").Text
            objCCProcessor.ValidateCard()

            Dim CUST_CREDIT_CARD_TYPE As String = ""

            Select Case objCCProcessor.CreditCardType
                Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                    CUST_CREDIT_CARD_TYPE = "AMEX"
                Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                    CUST_CREDIT_CARD_TYPE = "MSTR"
                Case TAC.ARCCCARD.CreditCardTypes.vctVisa
                    CUST_CREDIT_CARD_TYPE = "VISA"
                Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                    CUST_CREDIT_CARD_TYPE = "DISC"
                Case Else
                    CUST_CREDIT_CARD_TYPE = ""
            End Select

            If CUST_CREDIT_CARD_TYPE.Length > 0 Then
                e.Row.Cells("CUST_CREDIT_CARD_TYPE").Value = (AbsCon.Get_Image(AbsCon.Folders("Images") & "\ABS\CC\", CUST_CREDIT_CARD_TYPE & ".GIF"))
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdACH_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdACH.AfterRowActivate
        If InitializingControl Then Exit Sub
        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then Exit Sub

        If Not grdACH.ActiveRow.IsAddRow AndAlso grdACH.ActiveRow.Cells("WEB_IND").Value & String.Empty = "1" Then
            grdACH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Else
            If grdACH.DisplayLayout.Override.AllowUpdate <> Infragistics.Win.DefaultableBoolean.True Then
                grdACH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            End If
        End If
    End Sub

    Private Sub grdACH_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdACH.AfterRowUpdate
        If InitializingControl Then Exit Sub
        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then Exit Sub

        If e.Row.Cells("ACH_DEFAULT_ACCT_IND").Value & "" = "1" Then
            Dim ACH_ROUTING_NO As String = e.Row.Cells("ACH_ROUTING_NO").Text
            For Each row As DataRow In AbsCon.dst.Tables("ARTCUSPA").Select("ACH_ROUTING_NO <> '" & ACH_ROUTING_NO & "'", "")
                row.Item("ACH_DEFAULT_ACCT_IND") = ""
            Next
        End If
    End Sub

    Private Sub grdACH_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdACH.BeforeRowUpdate
        If InitializingControl Then Exit Sub
        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then
            e.Cancel = True
            Exit Sub
        End If

        With grdACH
            Dim errorMessage As String = String.Empty
            If Not ValidateRoutingNumber(e.Row.Cells("ACH_ROUTING_NO").Text, errorMessage) Then
                MessageBox.Show(errorMessage, "ACH", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If

            If e.Row.Cells("ACH_ACCT_STATUS").Text = "" Then
                .ActiveRow.Cells("ACH_ACCT_STATUS").Value = "A"
            End If

            If e.Row.IsAddRow Then

                If e.Row.Cells("CUST_CODE").Text = "" Then
                    .ActiveRow.Cells("CUST_CODE").Value = sCustomerCode
                End If

                .ActiveRow.Cells("ACH_ACCT_ID").Value = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSPA.ACH_ACCT_ID")
            End If
        End With

    End Sub

#End Region

#Region "Private Subs / Functions"

    Private Sub chkCCEdit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCCEdit.CheckedChanged

        If InitializingControl Then Exit Sub

        If AbsCon Is Nothing Then Exit Sub

        Try
            If chkCCEdit.Checked Then
                If sCustomerCode.Length = 0 Then
                    chkCCEdit.Checked = False
                    Exit Sub
                End If

                If Not AbsCon.Logical_Lock("ARTCUST1", sCustomerCode) Then
                    chkCCEdit.Checked = False
                    Exit Sub
                End If

                If Not AbsCon.Logical_Lock("ARTCUSTC", sCustomerCode) Then
                    chkCCEdit.Checked = False
                    Exit Sub
                End If

                ' Need to refresh the data incase it is stale.
                DisplayData(sCustomerCode, False)
                ' As of Today only Web Customers can set ACH; therefore, do not allow it to be changed on this screen
                ' it needs to be changed my the customer on the Web.
                If chkCUST_AUTOQ_WEB.Checked OrElse optCUST_AUTO_CCPA.Value = "3" Then
                    optCUST_AUTO_CCPA.Enabled = False
                    chkCUST_AUTOQ_WEB.Enabled = False
                Else
                    optCUST_AUTO_CCPA.Enabled = True
                    chkCUST_AUTOQ_WEB.Enabled = True
                End If

                chkCUST_AUTO_CC_AUTH.Enabled = sAllowAutoAuthForm
            Else
                AbsCon.MultiTask_Release()
                AbsCon.Fill_Records("ARTCUSTC", sCustomerCode)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            chkCCEdit.Checked = False
        End Try

        ToggleCCEditable()
    End Sub

    Private Sub ToggleCCEditable()
        If InitializingControl Then Exit Sub
        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            If chkCCEdit.Checked Then
                AbsCon.Set_Read_Only(grpCCPA, False)
                AbsCon.Set_Read_Only(grpButtons, False)
                SplitContainer25.Panel2Collapsed = False
                With grdCC.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.True
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                End With
                grdCC.DisplayLayout.Bands(0).Columns("CUST_CREDIT_CARD_NO").Hidden = False

                If allowACHEntries Then
                    With grdACH.DisplayLayout.Override
                        .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = Infragistics.Win.DefaultableBoolean.True
                        .AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                    End With
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ROUTING_NO").Hidden = False
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_NO").Hidden = False
                End If
                chkCUST_AUTO_CC_AUTH.Enabled = AllowAutoAuthForm
                If optCUST_AUTO_CCPA.Value = "3" Then
                    optCUST_AUTO_CCPA.Enabled = False
                Else
                    optCUST_AUTO_CCPA.Enabled = True
                End If

            Else
                AbsCon.Set_Read_Only(grpCCPA, True)
                AbsCon.Set_Read_Only(grpButtons, True)
                SplitContainer25.Panel2Collapsed = True
                With grdCC.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.False
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                End With
                If grdCC.DisplayLayout.Bands(0).Columns.Contains("CUST_CREDIT_CARD_NO") Then
                    grdCC.DisplayLayout.Bands(0).Columns("CUST_CREDIT_CARD_NO").Hidden = True
                End If

                With grdACH.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.False
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub cmdCCUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCCUpdate.Click
        If InitializingControl Then Exit Sub

        Dim beginTransaction As Boolean = False
        Dim rowSOTPARM1 As DataRow = Nothing
        Dim EMsg As String = String.Empty
        Dim sql As String = String.Empty
        Dim CUST_CODE As String = sCustomerCode

        Try
            If chkCCEdit.Checked Then
                Dim CUST_AUTO_CCPA As String = optCUST_AUTO_CCPA.Value & String.Empty
                Dim CUST_AUTO_CCPA_NOTE As String = txtCUST_AUTO_CCPA_NOTE.Text.Trim
                Dim CUST_AUTO_CC_AUTH As String = IIf(chkCUST_AUTO_CC_AUTH.Checked, "1", "0")
                Dim CUST_AUTOQ_WEB As String = IIf(chkCUST_AUTOQ_WEB.Checked, "1", "0")
                Dim CUST_AUTO_CC_AUTH_DATE As String = dteCUST_AUTO_CC_AUTH_DATE.Value & String.Empty
                Dim CUST_AUTO_CC_OPER As String = cbeCUST_AUTO_CC_OPER.Value & String.Empty

                If CUST_AUTO_CCPA.Length = 0 Then CUST_AUTO_CCPA = "0"

                If chkCUST_AUTO_CC_AUTH.Checked Then
                    If Not IsDate(CUST_AUTO_CC_AUTH_DATE) Then
                        CUST_AUTO_CC_AUTH_DATE = DateTime.Now.ToString("dd-MMM-yyyy")
                    End If

                    If CUST_AUTO_CC_OPER.Length = 0 Then
                        CUST_AUTO_CC_OPER = AbsCon.UserID
                    End If
                End If

                If AbsCon.dst.Tables("ARTCUSTC").Rows.Count = 0 _
                        AndAlso ",1,2,".Contains(CUST_AUTO_CCPA) Then
                    EMsg = "There must be at least one credit card on file for Reminder or Auto Charge."
                    MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If AbsCon.dst.Tables("ARTCUSPA").Rows.Count = 0 _
                        AndAlso ",3,".Contains(CUST_AUTO_CCPA) Then
                    EMsg = "There must be at least one Bank Account on file for ACH Processing."
                    MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If ",3,".Contains(CUST_AUTO_CCPA) AndAlso AbsCon.dst.Tables("ARTCUSPA").Rows.Count > 0 _
                        AndAlso AbsCon.dst.Tables("ARTCUSPA").SELECT("ACH_AUTO_PAY_IND = '1'").LENGTH = 0 Then
                    EMsg = "There must be a Bank Account with 'Auto Pay' choosen for ACH Processing."
                    MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If ",2,3,".Contains(CUST_AUTO_CCPA) AndAlso chkCUST_AUTO_CC_AUTH.Checked = False Then
                    EMsg = "You cannot select 'Auto Charge Queue' or 'ACH Queue' if the customer has not provided an Authorization Form."
                    MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim prevent_update As Boolean = False

                For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("", "", DataViewRowState.CurrentRows)
                    Dim CUST_CREDIT_CARD_NO As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & ""
                    Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") & ""
                    Dim CUST_CREDIT_CARD_LAST4 As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_LAST4") & ""

                    Try
                        objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                        objCCProcessor.CustomerCreditCard.CardExpMonth = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 1, 2) & "")
                        objCCProcessor.CustomerCreditCard.CardExpYear = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 3, 2) & "")
                        objCCProcessor.ValidateCard()
                    Catch ex As Exception
                        EMsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ":" & ex.Message
                    End Try

                    If Not objCCProcessor.DateCheckPassed Then
                        'EMsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ":" & "Expiration Date is Not Valid"
                    End If

                    If rowARTCUSTC.Item("CUST_CREDIT_CARD_NAME") & "" = "" Then
                        EMsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ":" & "Name Missing"
                    End If

                    If Not objCCProcessor.DigitCheckPassed Then
                        prevent_update = True
                        EMsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ":" & " CC Check Digit failed"
                    End If
                Next

                If EMsg <> "" Then
                    If prevent_update Then
                        MsgBox(EMsg, MsgBoxStyle.OkOnly, "There are problems with the Credit Card Data")
                        Exit Sub
                    Else
                        If MsgBox(EMsg & vbCr & vbCr & "Proceed with Update Anyway?", MsgBoxStyle.YesNo, "There are problems with the Credit Card Data") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                Dim add_now As Boolean = False

                Dim rowARTCUSTC_pref As DataRow = Nothing

                If CUST_AUTO_CCPA = "1" OrElse CUST_AUTO_CCPA = "2" Then

                    AbsCon.Fill_Records("ARTCCPA1", "", True, "Select * from ARTCCPA1 where CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYPP = '" & AbsCon.Period_Calc(AbsCon.CYP, -1) & "'")

                    Dim row() As DataRow = AbsCon.dst.Tables("ARTCCPA1").Select _
                    ("CCPA_REASON = 'A' and OPS_YYYYPP = '" & AbsCon.Period_Calc(AbsCon.CYP, -1) & "'" _
                     & " AND (CCPA_STATUS = '0' OR CCPA_STATUS = '1' OR CCPA_STATUS = '2')")

                    If row.Length = 0 Then
                        Dim SORT_VALUE As String = ""
                        For Each rowCC As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("CUST_CODE = '" & CUST_CODE & "' and CUST_CREDIT_CARD_STATUS = 'A'")
                            Dim YYMM As String = Mid(rowCC.Item("CUST_CREDIT_CARD_EXP_DATE") & "", 3, 2) & Mid(rowCC.Item("CUST_CREDIT_CARD_EXP_DATE") & "", 1, 2)
                            If YYMM >= Mid(AbsCon.CYM, 3, 4) Then
                                Dim CUST_CREDIT_CARD_PREFERRED As String = rowCC.Item("CUST_CREDIT_CARD_PREFERRED") & ""
                                If CUST_CREDIT_CARD_PREFERRED = "" Then
                                    CUST_CREDIT_CARD_PREFERRED = "0"
                                End If
                                Dim SORT_VALUE_CC As String = CUST_CREDIT_CARD_PREFERRED & YYMM
                                If SORT_VALUE = "" Or SORT_VALUE_CC > SORT_VALUE Then
                                    rowARTCUSTC_pref = AbsCon.dst.Tables("ARTCUSTC").NewRow
                                    rowARTCUSTC_pref.ItemArray = rowCC.ItemArray
                                    SORT_VALUE = SORT_VALUE_CC
                                End If

                            End If
                        Next
                    Else
                        add_now = True
                    End If

                    If rowARTCUSTC_pref IsNot Nothing Then
                        If MsgBox("Add this customer to Queue Now?", MsgBoxStyle.YesNo, "Option to Add Customer to Queue Now") = MsgBoxResult.Yes Then
                            add_now = True
                        End If
                    End If

                End If

                AbsCon.BeginTrans()
                beginTransaction = True

                ' Record event if the user chnages the  CUST_AUTO_CCPA or CUST_AUTO_CCPA_NOTE
                Dim zMsg As String = String.Empty
                If rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty <> CUST_AUTO_CCPA Then
                    zMsg = "CC Queue changed from "

                    Select Case rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty
                        Case "1"
                            zMsg &= optCUST_AUTO_CCPA.Items(1).DisplayText
                        Case "2"
                            zMsg &= optCUST_AUTO_CCPA.Items(2).DisplayText
                        Case Else
                            zMsg &= optCUST_AUTO_CCPA.Items(0).DisplayText
                    End Select

                    zMsg &= " to "
                    zMsg &= optCUST_AUTO_CCPA.CheckedItem.DisplayText

                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                If (rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE", DataRowVersion.Original) & String.Empty).ToString.Trim.ToUpper <> CUST_AUTO_CCPA_NOTE.Trim.ToUpper Then
                    zMsg = "CC Note changed to: " & CUST_AUTO_CCPA_NOTE
                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                If (rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty).ToString.Trim.ToUpper <> CUST_AUTO_CCPA.Trim.ToUpper Then
                    zMsg = "CC Queue changed to: " & CUST_AUTO_CCPA
                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                rowARTCUST1.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE

                Dim rowARTCUST1_orig As DataRow = AbsCon.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
                rowARTCUST1_orig.Table.TableName = "ARTCUST1"

                rowARTCUST1.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE
                rowARTCUST1.Item("CUST_AUTO_CC_AUTH") = CUST_AUTO_CC_AUTH
                rowARTCUST1.Item("CUST_AUTOQ_WEB") = CUST_AUTOQ_WEB

                If CUST_AUTO_CC_AUTH_DATE.Length = 0 Then
                    sql = "Update ARTCUST1 Set"
                    sql &= "  CUST_AUTO_CCPA = :PARM1"
                    sql &= ", CUST_AUTO_CCPA_NOTE = :PARM2"
                    sql &= ", CUST_AUTO_CC_AUTH = :PARM3"
                    sql &= ", CUST_AUTOQ_WEB = :PARM4"
                    sql &= " where CUST_CODE = :PARM5"

                    AbsCon.ExecuteSQL(sql, "VVVVV", New Object() { _
                        CUST_AUTO_CCPA, _
                        CUST_AUTO_CCPA_NOTE, _
                        CUST_AUTO_CC_AUTH, CUST_AUTOQ_WEB, _
                        CUST_CODE})
                Else

                    sql = "Update ARTCUST1 Set"
                    sql &= "  CUST_AUTO_CCPA = :PARM1"
                    sql &= ", CUST_AUTO_CCPA_NOTE = :PARM2"
                    sql &= ", CUST_AUTO_CC_AUTH = :PARM3"
                    sql &= ", CUST_AUTO_CC_AUTH_DATE = :PARM4"
                    sql &= ", CUST_AUTO_CC_OPER = :PARM5"
                    sql &= ", CUST_AUTOQ_WEB = :PARM6"
                    sql &= " where CUST_CODE = :PARM7"

                    AbsCon.ExecuteSQL(sql, "VVVDVVV", New Object() { _
                        CUST_AUTO_CCPA, _
                        CUST_AUTO_CCPA_NOTE, _
                        CUST_AUTO_CC_AUTH, _
                        IIf(CUST_AUTO_CC_AUTH_DATE.Length > 0, CDate(CUST_AUTO_CC_AUTH_DATE).ToString("dd-MMM-yyyy"), ""), _
                        CUST_AUTO_CC_OPER, CUST_AUTOQ_WEB, _
                        CUST_CODE})

                    rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") = IIf(CUST_AUTO_CC_AUTH_DATE.Length > 0, CUST_AUTO_CC_AUTH_DATE, "")
                    rowARTCUST1.Item("CUST_AUTO_CC_OPER") = CUST_AUTO_CC_OPER

                End If

                Dim rowARTCUST1_curr As DataRow = AbsCon.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
                rowARTCUST1_curr.Table.TableName = "ARTCUST1"
                rowARTCUST1_curr.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1_curr.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE
                rowARTCUST1_curr.Item("CUST_AUTO_CC_AUTH") = CUST_AUTO_CC_AUTH
                rowARTCUST1_curr.Item("CUST_AUTOQ_WEB") = CUST_AUTOQ_WEB
                If CUST_AUTO_CC_AUTH_DATE.Length > 0 Then
                    rowARTCUST1_curr.Item("CUST_AUTO_CC_AUTH_DATE") = CUST_AUTO_CC_AUTH_DATE
                End If
                rowARTCUST1_curr.Item("CUST_AUTO_CC_OPER") = CUST_AUTO_CC_OPER

                AbsCon.DATETIME_STAMP = DateTime.Now
                AbsCon.Write_Audit_Trail(rowARTCUST1_curr, rowARTCUST1_orig, "E")

                AbsCon.INIT_LAST("ARTCUSTC", True, "CUST_CODE = '" & CUST_CODE & "'")
                AbsCon.Update_Record_TDA("ARTCUSTC", "CUST_CODE = '" & CUST_CODE & "'")

                AbsCon.INIT_LAST("ARTCUSPA", True, "CUST_CODE = '" & CUST_CODE & "'")
                EncryptDecryptArtpymtw("ARTCUSPA", Encryption.Encrypt)
                AbsCon.Update_Record_TDA("ARTCUSPA", "CUST_CODE = '" & CUST_CODE & "'")

                Record_Customer_Event(CUST_CODE, "CC Info Updated", "C")

                If add_now Or CUST_AUTO_CCPA = "0" Then
                    '                    Dim rowARTCCPA1 As DataRow = Nothing
                    '                    AbsCon.dst.Tables("ARTCCPA1").AcceptChanges()

                    '                    sql = "CUST_CODE = '" & CUST_CODE & "' AND CCPA_STATUS IN ('0', '1', '2') AND CCPA_REASON = 'A' and OPS_YYYYPP = '" & AbsCon.Period_Calc(AbsCon.CYP, -1) & "'"

                    '                    If AbsCon.dst.Tables("ARTCCPA1").Select(sql).Length > 0 Then
                    '                        rowARTCCPA1 = AbsCon.dst.Tables("ARTCCPA1").Select(sql)(0)
                    '                        If CUST_AUTO_CCPA = "1" Then
                    '                            rowARTCCPA1.Item("CCPA_STATUS") = "0"
                    '                        ElseIf CUST_AUTO_CCPA = "2" Then
                    '                            rowARTCCPA1.Item("CCPA_STATUS") = "2"
                    '                        ElseIf CUST_AUTO_CCPA = "0" Then
                    '                            rowARTCCPA1.Item("CCPA_STATUS") = "X"
                    '                        End If

                    '                        AbsCon.INIT_LAST("ARTCCPA1", True, , True)
                    '                        GoTo UpdateArtccpa1
                    '                    End If

                    '                    ' Do not create a new entry if not to be queued
                    '                    If CUST_AUTO_CCPA = "0" Then
                    '                        GoTo UpdateArtccpa1
                    '                    End If

                    '                    rowARTCCPA1 = AbsCon.dst.Tables("ARTCCPA1").NewRow
                    '                    rowARTCCPA1.Item("CCPA_NO") = AbsCon.Next_Control_No("ARTCCPA1.CCPA_NO")
                    '                    AbsCon.dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)
                    '                    rowARTCCPA1.Item("CUST_CODE") = CUST_CODE
                    '                    rowARTCCPA1.Item("CCPA_REASON") = "A"
                    '                    rowARTCCPA1.Item("OPS_YYYYPP") = AbsCon.Period_Calc(AbsCon.CYP, -1)

                    '                    If CUST_AUTO_CCPA = "1" Then
                    '                        rowARTCCPA1.Item("CCPA_STATUS") = "0"
                    '                    Else
                    '                        rowARTCCPA1.Item("CCPA_STATUS") = "2"
                    '                    End If

                    '                    rowARTCCPA1.Item("CCPA_NOTE") = CUST_AUTO_CCPA_NOTE
                    '                    rowARTCCPA1.Item("INIT_DATE") = Now + AbsCon.NowTSD
                    '                    rowARTCCPA1.Item("INIT_OPER") = AbsCon.UserID

                    '                    For Each COLUMN_NAME As String In New String() _
                    '                    {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE" _
                    '                     , "CUST_CREDIT_CARD_VER_CODE", "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
                    '                     , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE", "CUST_CREDIT_CARD_ZIP_CODE" _
                    '                     , "CUST_CREDIT_CARD_LAST4"}
                    '                        rowARTCCPA1.Item(COLUMN_NAME) = rowARTCUSTC_pref.Item(COLUMN_NAME)
                    '                    Next

                    '                    Dim CCPA_AMT As Decimal = 0
                    '                    Dim OPS_YYYYPP As String = AbsCon.Period_Calc(AbsCon.CYP, -1)
                    '                    Dim rowARTSTMT1 As DataRow = AbsCon.LookUp("ARTSTMT1", New String() {OPS_YYYYPP, CUST_CODE})
                    '                    If rowARTSTMT1 IsNot Nothing Then
                    '                        CCPA_AMT = Val(rowARTSTMT1.Item("TOTAL_DUE") & "")
                    '                    End If
                    '                    rowARTCCPA1.Item("CCPA_AMT") = CCPA_AMT
                    '                    AbsCon.INIT_LAST("ARTCCPA1", True, , True)

                    'UpdateArtccpa1:

                    '                    AbsCon.Update_Record_TDA("ARTCCPA1")

                End If

                AbsCon.CommitTrans("CC Info Updated")

                chkCCEdit.Checked = False
            End If
            RaiseEvent UpdateClickEvent(rowARTCUST1)

        Catch ex As Exception
            If beginTransaction = True Then
                AbsCon.Rollback()
            End If
            MessageBox.Show("Error updating Credit Card Data: " & ex.Message, "Error", MessageBoxButtons.OK)
        End Try

        AbsCon.Fill_Records("ARTCUST1", CUST_CODE)
        rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").Rows(0)
        SetupCustomerMasterData()

    End Sub

    Private Sub cmdCCCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCCCancel.Click
        If InitializingControl Then Exit Sub

        AbsCon.dst.Tables("ARTCUST1").RejectChanges()
        AbsCon.dst.Tables("ARTCUSTC").RejectChanges()
        AbsCon.dst.Tables("ARTCUSPA").RejectChanges()
        rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").Rows(0)
        SetupCustomerMasterData()
        chkCCEdit.Checked = False

        RaiseEvent CancelClickEvent()
    End Sub

    Private Sub Record_Customer_Event(ByVal CUST_CODE As String, ByVal EVENT_DESC As String, ByVal EVENT_TYPE As String)

        If InitializingControl Then Exit Sub

        Dim sql As String = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_OPER, INIT_DATE, EVENT_DESC, EVENT_TYPE, FORM_NAME)" _
        & " Select 'ARTCUST1', CUST_CODE, :PARM1, SYSDATE, :PARM2, :PARM3, :PARM4 from ARTCUST1 where CUST_CODE = :PARM5"
        AbsCon.ExecuteSQL(sql, "VVVVV", New Object() {AbsCon.UserID, EVENT_DESC, EVENT_TYPE, AbsCon.ActiveFormName, CUST_CODE})
    End Sub

    Private Sub SetupCustomerMasterData()
        If InitializingControl Then Exit Sub
        Try
            Select Case rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty

                Case "0", "1", "2", "3"
                    optCUST_AUTO_CCPA.Value = rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty
                Case Else
                    optCUST_AUTO_CCPA.Value = "0"
            End Select
            txtCUST_AUTO_CCPA_NOTE.Text = rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") & String.Empty

            If rowARTCUST1.Item("CUST_AUTO_CC_AUTH") & String.Empty = "1" Then
                chkCUST_AUTO_CC_AUTH.Checked = True
            Else
                chkCUST_AUTO_CC_AUTH.Checked = False
            End If

            If rowARTCUST1.Item("CUST_AUTOQ_WEB") & String.Empty = "1" Then
                chkCUST_AUTOQ_WEB.Checked = True
            Else
                chkCUST_AUTOQ_WEB.Checked = False
            End If

            If chkCUST_AUTOQ_WEB.Checked Then
                optCUST_AUTO_CCPA.Enabled = False
                chkCUST_AUTOQ_WEB.Enabled = False
            Else
                optCUST_AUTO_CCPA.Enabled = True
                chkCUST_AUTOQ_WEB.Enabled = True
            End If

            cbeCUST_AUTO_CC_OPER.Value = rowARTCUST1.Item("CUST_AUTO_CC_OPER") & String.Empty

            If IsDate(rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") & String.Empty) Then
                dteCUST_AUTO_CC_AUTH_DATE.Value = CDate(rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") & String.Empty)
            Else
                dteCUST_AUTO_CC_AUTH_DATE.Value = String.Empty
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub EncryptDecryptArtpymtw(ByVal tableName As String, ByVal procType As Encryption)
        If InitializingControl Then Exit Sub

        Dim temp As String = String.Empty
        Dim crypt As New TAC.ASCSCRTY
        For Each row As DataRow In AbsCon.dst.Tables(tableName).Rows

            Select Case procType
                Case Encryption.Encrypt
                    temp = crypt.Encrypt_AES(row.Item("ACH_ROUTING_NO") & String.Empty)
                    row.Item("ACH_ROUTING_NO") = temp
                    temp = crypt.Encrypt_AES(row.Item("ACH_ACCT_NO") & String.Empty)
                    row.Item("ACH_ACCT_NO") = temp
                Case Else
                    temp = crypt.Decrypt_AES(row.Item("ACH_ROUTING_NO"))
                    temp = temp.Replace(Chr(0), String.Empty).Trim()
                    row.Item("ACH_ROUTING_NO") = temp

                    temp = crypt.Decrypt_AES(row.Item("ACH_ACCT_NO") & String.Empty)
                    temp = temp.Replace(Chr(0), String.Empty).Trim()
                    row.Item("ACH_ACCT_NO") = temp
            End Select
        Next

    End Sub

    ''' <summary>
    ''' Validate a Bank Routing Number
    ''' </summary>
    ''' <param name="routingNumber">Routing Number to evaluate</param>
    ''' <param name="resp">If returnsd false, error description is in resp</param>
    ''' <remarks></remarks>
    Public Shared Function ValidateRoutingNumber(ByVal routingNumber As String, ByRef resp As String) As Boolean

        If Not System.Text.RegularExpressions.Regex.IsMatch(routingNumber, "^\d+$") Then
            resp = "Routing number cannot have non-numeric characters"
            Return False

        ElseIf routingNumber.Length <> 9 Then
            resp = "Routing number must be 9 digits long"
            Return False

        Else
            Dim isValid As Boolean
            isValid = (Integer.Parse(routingNumber(0)) * 3 + Integer.Parse(routingNumber(1)) * 7 + Integer.Parse(routingNumber(2)) * 1 + _
                      Integer.Parse(routingNumber(3)) * 3 + Integer.Parse(routingNumber(4)) * 7 + Integer.Parse(routingNumber(5)) * 1 + _
                      Integer.Parse(routingNumber(6)) * 3 + Integer.Parse(routingNumber(7)) * 7 + Integer.Parse(routingNumber(8)) * 1) Mod 10 = 0
            If Not isValid Then
                resp = "Routing number is invalid"
                Return False
            End If
        End If

        Return True
    End Function

#End Region

End Class
