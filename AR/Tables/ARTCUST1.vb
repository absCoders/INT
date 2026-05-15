Public Class ARTCUST1

    ' For now only allow W, WP and IPE
    ' adding Y 03/16 as per emails LM
    ' adding B 04/11 as per emails LM
    ' adding IT 01/16/2019 as per emails LM

    Private IPLB3plOrderCodes As List(Of String) = New List(Of String)({"W", "WP", "IPE", "Y", "B", "IT"})


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)
            .Tables("ARTCUST2").Columns.Add("SHIP_VIA_DESC", GetType(System.String), "CUST_STORE_SHIP_VIA_CODE")

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
            & " from ARTCUSTD " _
            & " where ARTCUSTD.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)
        End With

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")

        grdARTCUST2.DisplayLayout.UseFixedHeaders = True
        With grdARTCUST2.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Header.Fixed = True
            .Columns("CUST_STORE_NAME").Header.Fixed = True
        End With
        Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")

        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        ASCMAIN1.Add_Value_List(grdARTCUST2, "SHIP_VIA_DESC", "SELECT SHIP_VIA_CODE, SHIP_VIA_DESC FROM SOTSVIA1")
        'Call InitializeControls(Me)

        Set_Read_Only_for_ctl(Absx1.chkFor("CUST_SHIP_COMPLETE"), True)
        Set_Read_Only_for_ctl(Absx1.chkFor("CUST_CONS_INV"), True)

        UltraTabControl3.Tabs("Credit Cards").Visible = (ROWs("ARTPARM1").Item("AR_PARM_ENABLE_CC") & "" = "1")

        cmdCFG.Visible = False
        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            If ASCMAIN1.Running_in_VS Then
                cmdCFG.Visible = True
            End If

            Absx1.optFor("CUST_FRT_CHG_CODE").Visible = True
            Absx1.chkFor("CUST_SHIP_TO_MANUAL").Visible = True
        Else
            Absx1.optFor("CUST_FRT_CHG_CODE").Visible = False
            Absx1.chkFor("CUST_SHIP_TO_MANUAL").Visible = False

        End If
        '    Absx1.chkFor("CUST_SHIP_COMPLETE").Enabled = False
        '    Absx1.chkFor("CUST_CONS_INV").Enabled = False

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            btnADSCustomerNo.Visible = False
        End If

        Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from TATCURR1")
        If tbl.Rows.Count < 2 Or (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT") Then
            lblCURR_CODE.Visible = False
            txtCURR_CODE.Visible = False
        Else
            lblCURR_CODE.Visible = True
            txtCURR_CODE.Visible = True
        End If

        If ASCMAIN1.CLIENT = "SLP" Then
            Absx1.chkFor("CUST_PII_IND").Visible = True
            Absx1.chkFor("CUST_ECOM_IND").Visible = True
            Absx1.chkFor("CUST_PII_IND").Top = Absx1.chkFor("CUST_SALES_HOLD").Top
            Absx1.chkFor("CUST_ECOM_IND").Top = Absx1.chkFor("CUST_PO_REQD").Top

            Absx1.txtFor("CUST_NO_3PL2").Visible = True
            Absx1.txtFor("CUST_NO_3PL2").Top = Absx1.txtFor("CUST_STORE_NO_3PL").Top
            Absx1.txtFor("CUST_NO_3PL2").Left = Absx1.txtFor("CUST_STORE_NO_3PL").Left
            Absx1.txtFor("CUST_STORE_NO_3PL").Visible = False
            UltraTabControl3.Tabs("Contacts").Visible = False
            UltraTabControl3.Tabs("Credit").Visible = False
            UltraTabControl3.Tabs("Accounting").Visible = False

            Absx1.chkFor("CUST_SALES_HOLD").Visible = False
            Absx1.chkFor("CUST_PO_REQD").Visible = False
            Absx1.chkFor("CUST_INCL_INV_SHIP").Visible = False
            Absx1.chkFor("CUST_CONS_INV").Visible = False
            Absx1.chkFor("CUST_EDI_COMM_SEP").Visible = False
            Absx1.chkFor("CUST_INCL_CAL").Visible = False

        End If

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    'Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
                    'Dim CUST_ADDR1 As String = Absx1.txtFor("CUST_ADDR1").Text
                    'Dim CUST_ADDR2 As String = Absx1.txtFor("CUST_ADDR2").Text
                    'Dim CUST_CITY As String = Absx1.txtFor("CUST_CITY").Text
                    'Dim rx As String = "^([A-Z0-9#-]+)$" ' Allow Upper case, numbers

                    'Dim r As New System.Text.RegularExpressions.Regex(rx)
                    'If Not r.IsMatch(Replace(Replace(CUST_NAME, ",", ""), " ", "")) Then
                    '    EMsg &= vbCr & "The Customer Name has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(Replace(Replace(CUST_ADDR1, ",", ""), " ", "")) Then
                    '    EMsg &= vbCr & "The Customer Address 1 has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(Replace(Replace(CUST_ADDR2, ",", ""), " ", "")) Then
                    '    EMsg &= vbCr & "The Customer Address 2 has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(Replace(Replace(CUST_CITY, ",", ""), " ", "")) Then
                    '    EMsg &= vbCr & "The City has Special Characters which are not allowed"
                    'End If
                End If

                ' Make Sure Upper Case Code.
                Absx1.txtFor("ORDR_CODE_3PL").Text = Absx1.txtFor("ORDR_CODE_3PL").Text.ToUpper.Trim
                dst.Tables("ARTCUST1").Rows(0).Item("ORDR_CODE_3PL") = Absx1.txtFor("ORDR_CODE_3PL").Text

                If Absx1.txtFor("ORDR_CODE_3PL").TextLength > 0 AndAlso _
                    Not IPLB3plOrderCodes.Contains(Absx1.txtFor("ORDR_CODE_3PL").Text) Then
                    EMsg &= "Currently, only " & String.Join(", ", IPLB3plOrderCodes) & " are permitted as the 3PL Order Code."
                End If


                Dim rowTATSTATE As DataRow = LookUp("TATSTATE", Absx1.txtFor("CUST_STATE").Text)
                If rowTATSTATE Is Nothing Then
                    EMsg &= "Invalid Value Specified for State"
                End If

                If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                    If Absx1.txtFor("CUST_STATE").Text = "US" Or Absx1.txtFor("CUST_STATE").Text = "USA" Then
                        EMsg &= "Leave Country Blank for USA"
                    End If
                End If

                Dim CUST_BILL_TO_CUST As String = Absx1.txtFor("CUST_BILL_TO_CUST").Text
                Dim CUST_CREDIT_GROUP_CUST As String = Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text

                If CUST_BILL_TO_CUST <> "" Then
                    If CUST_BILL_TO_CUST = CUST_CODE Then
                        EMsg &= "No Need to set Customer's Bill-To Customer to itself"
                    ElseIf CUST_CREDIT_GROUP_CUST <> "" Then
                        EMsg &= "Cannot set Credit Group Customer and Bill-To Customer - set one or the other"
                    Else
                        If LookUp("ARTCUST1", CUST_BILL_TO_CUST) Is Nothing Then
                            EMsg &= "Invalid Value Specified for Bill-To Customer"
                        End If
                    End If
                Else
                    If CUST_CREDIT_GROUP_CUST <> "" Then
                        If CUST_CREDIT_GROUP_CUST = CUST_CODE Then
                            EMsg &= "No Need to set Customer's Credit Group Customer to itself"
                        Else
                            If LookUp("ARTCUST1", CUST_CREDIT_GROUP_CUST) Is Nothing Then
                                EMsg &= "Invalid Value Specified for Credit Group Customer"
                            End If
                        End If
                    End If
                End If

                ' DO THE FOLLOWING FOR MANDATORY CODES
                'For Each COLUMN_NAME As String In New String() _
                '    {"TERM_CODE", "SREP_CODE", "POST_CODE", "STAX_CODE", "TRADE_CLASS_CODE", "PRICE_CLASS_CODE", "PRICE_LIST_CODE", _
                '     "POST_CODE", "CUST_CLASS_CODE", "SHIP_VIA_CODE", "ROUTING_CODE", "WHSE_CODE", "CUST_BILL_TO_CUST", "CUST_CREDIT_GROUP_CUST", "VEND_CODE"}
                '    Validate_Code(COLUMN_NAME)
                'Next

                Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                If rowSOTSREP1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value specified for Sales Rep Code"
                End If

                Dim rowSOTTCLS1 = LookUp("SOTTCLS1", Absx1.txtFor("TRADE_CLASS_CODE").Text)
                If rowSOTTCLS1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value specified for Trade Class Code"
                End If

                Dim rowSOTPCLS1 = LookUp("SOTPCLS1", Absx1.txtFor("PRICE_CLASS_CODE").Text)
                If rowSOTPCLS1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value specified for Price Class Code"
                End If

                For Each ROW As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTCUSTD"), New String() {"CONTACT_TYPE"}).Rows
                    Dim CONTACT_TYPE As String = ROW.Item("CONTACT_TYPE") & ""
                    Dim sqlw As String = "CONTACT_TYPE = '" & CONTACT_TYPE & "'"
                    Dim c As Integer = Val(dst.Tables("ARTCUSTD").Compute("COUNT(CONTACT_NO)", sqlw & " and CONTACT_PRIMARY = '1'") & "")
                    If c > 1 Then
                        EMsg &= vbCr & "Cannot have > 1 Primary Contact of any Type (see Type " & CONTACT_TYPE & ")"
                    ElseIf c = 0 Then
                        Dim rows() As DataRow = dst.Tables("ARTCUSTD").Select(sqlw)
                        If rows.Length = 1 Then
                            rows(0).Item("CONTACT_PRIMARY") = "1"
                        Else
                            EMsg &= vbCr & "You must select a Primary Contact for each Type of Contact (see Type " & CONTACT_TYPE & ")"
                        End If
                    End If
                Next

                If EMsg.Length = 0 Then
                    ' validate City, State Zip.

                    Dim sqlCityState As String = String.Empty

                    MyBase.Absx1.txtFor("CUST_STATE").Text = MyBase.Absx1.txtFor("CUST_STATE").Text.Trim.ToUpper
                    MyBase.Absx1.txtFor("CUST_CITY").Text = MyBase.Absx1.txtFor("CUST_CITY").Text.Trim.ToUpper
                    MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text = MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text.Trim.ToUpper

                    Dim tblSOTZIPLK As DataTable = _
                        ASCDATA1.GetDataTable("Select * from SOTZIPLK Where STATE_CODE = :PARM1 AND CITY = :PARM2 AND ZIP_CODE = :PARM3", _
                                              "SOTZIPLK", "VVV", _
                                              MyBase.Absx1.txtFor("CUST_STATE").Text, _
                                              MyBase.Absx1.txtFor("CUST_CITY").Text, _
                                              Mid(MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text, 1, 5))


                    If tblSOTZIPLK.Rows.Count = 0 Then
                        tblSOTZIPLK = _
                            ASCDATA1.GetDataTable("Select * from SOTZIPLK Where STATE_CODE = :PARM1 AND ZIP_CODE = :PARM2", _
                                                  "SOTZIPLK", "VV", _
                                                  MyBase.Absx1.txtFor("CUST_STATE").Text, _
                                                  Mid(MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text, 1, 5))

                        If tblSOTZIPLK.Rows.Count > 0 Then
                            If MessageBox.Show("There were no entries found for the provided City/State/Zip Code combination. However, there are entires for the" _
                                                & " provided State/Zip Code." & Environment.NewLine & "Would you like to see the City options?", "Address Validation", MessageBoxButtons.YesNo, _
                                                  MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                                sqlCityState = "Select * from SOTZIPLK Where STATE_CODE = '" & MyBase.Absx1.txtFor("CUST_STATE").Text & "' AND ZIP_CODE = '" & Mid(MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text, 1, 5) & "'"
                            End If
                        Else
                            tblSOTZIPLK = _
                                ASCDATA1.GetDataTable("Select * from SOTZIPLK Where ZIP_CODE = :PARM1", _
                                                      "SOTZIPLK", "V", _
                                                      Mid(MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text, 1, 5))
                            If tblSOTZIPLK.Rows.Count > 0 Then
                                If MessageBox.Show("There were no entries found for the provided City/State/Zip Code combination. However, there are entires for the" _
                                                    & " provided Zip Code." & Environment.NewLine & "Would you like to see the City/State options?", "Address Validation", MessageBoxButtons.YesNo, _
                                                      MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                                    sqlCityState = "Select * from SOTZIPLK Where ZIP_CODE = '" & Mid(MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text, 1, 5) & "'"
                                End If
                            Else
                                If MessageBox.Show("There were no City/State/Zip Code entries found for the provided Zip Code. Do you want to continue updating the record?", _
                                                     "Address Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                                    EMsg = vbCr & "Update cancelled by user."
                                End If
                            End If
                        End If
                    End If

                    If sqlCityState.Length > 0 AndAlso tblSOTZIPLK.Rows.Count > 0 Then
                        ASCMAIN1.CodeSelector.Get_SQL("")
                        ASCMAIN1.CodeSelector.VIEW_NAME = String.Empty
                        ASCMAIN1.CodeSelector.SQL = sqlCityState
                        ASCMAIN1.CodeSelector.UseDataFromTable = tblSOTZIPLK
                        ASCMAIN1.CodeSelector.MultipleSelections = False
                        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                        Using F As New ASFCODE1
                            F.ShowDialog()
                        End Using
                        If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                            MyBase.Absx1.txtFor("CUST_STATE").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STATE_CODE")
                            MyBase.Absx1.txtFor("CUST_CITY").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CITY")
                            MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ZIP_CODE")

                            dst.Tables("ARTCUST1").Rows(0).Item("CUST_STATE") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STATE_CODE")
                            dst.Tables("ARTCUST1").Rows(0).Item("CUST_CITY") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CITY")
                            dst.Tables("ARTCUST1").Rows(0).Item("CUST_ZIP_CODE") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ZIP_CODE")

                        End If
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        'Update_Record_TDA("ARTCUST2")
        Update_Record_TDA("ARTCUSTD")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        Else
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            rowASFBASE1.Item("CUST_CREDIT_LIMIT") = Val(ROWs("ARTPARM1").Item("AR_PARM_INITIAL_CR_LIMIT") & "")
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_CRED_LIMIT_EST") = DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CUST_CREDIT_LIMIT_NOTES") = "Initial Credit Limit"
            rowASFBASE1.Item("CUST_STMT_IND") = "M"
            rowASFBASE1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE")
            rowASFBASE1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
            rowASFBASE1.Item("CUST_STATUS") = "A"
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_STATUS_DATE") = Now.Date ' DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            rowASFBASE1.Item("CUST_NO_3PL") = DBNull.Value
            rowASFBASE1.Item("CUST_STORE_NO_3PL") = DBNull.Value
        End If

        EnforceConstraints(False)
        Fill_Records("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text})
        Fill_Records("ARTCUSTD", New String() {Absx1.txtFor("CUST_CODE").Text})
        EnforceConstraints(True)

        CreditCardQueue1.ClearData()
        CreditCardQueue1.AllowAutoAuthForm = True
        CreditCardQueue1.DisplayData(Absx1.txtFor("CUST_CODE").Text)
        'CreditCardQueue1.AllowEdit = (EntryMode = "Edit" AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("CL"))
        CreditCardQueue1.AllowEdit = (EntryMode = "Edit" AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains(ROWs("ARTPARM1").Item("AR_PARM_SEC_ISSUE_CRD") & String.Empty))

    End Sub

    Overrides Sub Clear_Record_Special()
        If SELECTION_NO = 0 Then Exit Sub
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ARTCUST2").Rows.Clear()
            dst.Tables("ARTCUSTD").Rows.Clear()
            EnforceConstraints(True)

            CreditCardQueue1.ClearData()
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_NAME"), Not tf)
        ' Set_Read_Only(grpCreditLimit, IIf(Not tf, ASCMAIN1.USER_SECURITY_CODEs.Contains("CL"), True))
        Set_Read_Only(grpCreditLimit, True)
        Absx1.chkFor("CUST_SHIP_TO_MANUAL").Enabled = (EntryMode = "New")
        Absx1.txtFor("CURR_CODE").Enabled = (EntryMode = "New")

        If ASCMAIN1.CLIENT = "INT" And (EntryMode = "New" Or EntryMode = "Edit") Then
            Absx1.txtFor("PRICE_CLASS_CODE").ReadOnly = Not (EntryMode = "New" Or ASCMAIN1.USER_SECURITY_CODEs.Contains(ROWs("SOTPARM1").Item("SO_PARM_SEC_PRICE") & ""))
            Absx1.txtFor("PRICE_LIST_CODE").ReadOnly = Not (EntryMode = "New" Or ASCMAIN1.USER_SECURITY_CODEs.Contains(ROWs("SOTPARM1").Item("SO_PARM_SEC_PRICE") & ""))
        End If
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCUSTD, grdARTCUST2}
            With grd.DisplayLayout.Override
                If (EntryMode = "New" Or EntryMode = "Edit") And grd.Name <> "grdARTCUST2" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next

        If (EntryMode = "Edit" OrElse EntryMode = "New") Then
            CreditCardQueue1.SetReadOnly()
            MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = True
            MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = True
            btnADSCustomerNo.Enabled = True
            If (EntryMode = "Edit" And Absx1.txtFor("CUST_NO_3PL").Text = "") And ASCMAIN1.CLIENT = "INT" Then
                MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = False
                MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = False
            End If
        Else
            MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = True
            MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = True
            btnADSCustomerNo.Enabled = False
        End If

    End Sub

    Public Overrides Sub isDeleteAllowed()
        If SELECTION_NO = 0 Then Exit Sub

        MyBase.isDeleteAllowed()
        If EMsg = "" Then
            isDeleteAllowed_Check_Aliased_Columns _
            (New String() {"ARTCUST1.CUST_BILL_TO_CUST"})
        End If
    End Sub

    Public Overrides Function Set_Contact_Info() As Boolean
        If SELECTION_NO = 0 Then Exit Function
        If ScreenMode Then
            CONTACT_ENTITY_KEY = Absx1.txtFor("CUST_CODE").Text
            CONTACT_ENTITY_NAME = rowASFBASE1.Item("CUST_NAME") & "" ' .txtFor("CUST_NAME").Text
        End If
        Return True
    End Function

#End Region

#Region "grdARTCUST2"

    Private Sub grdARTCUST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.AfterCellUpdate
        'Select Case e.Cell.Column.Key
        '    Case "CUST_CODE"
        '        Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
        '        If row IsNot Nothing Then
        '            grdARTCUST2.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
        '        End If
        'End Select
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterRowActivate
        'With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
        '    If grdARTCUST2.ActiveRow.IsAddRow Then
        '        .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdARTCUST2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUST2.BeforeRowsDeleted

    End Sub

    Private Sub grdARTCUST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        'If row Is Nothing Then
        '    e.Cancel = True
        'End If
    End Sub

    Private Sub grdARTCUST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        'Dim sql_where As String = ""
        'Call grdClickCellButton(grdARTCUST2, sql_where, True)
    End Sub

    Private Sub grdARTCUST2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUST2.Error
        grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUST2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST2.InitializeRow
    End Sub

#End Region

#Region "grdARTCUSTD"

    Private Sub grdARTCUSTD_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "CUST_CODE"
            '    Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
            '    If row IsNot Nothing Then
            '        grdARTCUSTD.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
            '    End If
        End Select
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUSTD.AfterRowActivate
        With grdARTCUSTD.DisplayLayout.Bands("ARTCUSTD")
            'If grdARTCUSTD.ActiveRow.IsAddRow Then
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            'Else
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            'End If
        End With
    End Sub

    Private Sub grdARTCUSTD_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUSTD.BeforeRowsDeleted

    End Sub

    Private Sub grdARTCUSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTD.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        'If row Is Nothing Then
        '    e.Cancel = True
        'End If
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            e.Row.Cells("CONTACT_NO").Value = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        End If

        If e.Row.Cells("CONTACT_TYPE").Value & String.Empty = String.Empty Then
            MessageBox.Show("Contact Type is required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub grdARTCUSTD_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdARTCUSTD, sql_where, True)
    End Sub

    Private Sub grdARTCUSTD_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTD.Error
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUSTD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTD.InitializeRow
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        MyBase.Load_Popup_Menus()
        Load_Popup_Menu(CreditCardQueue1.UserControlGrid, "B", "Use Customer Address")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If SELECTION_NO = 0 Then Exit Sub
        If e.SourceControl.Name = CreditCardQueue1.UserControlGrid.Name Then
            If CreditCardQueue1.UserControlGrid.ActiveRow Is Nothing Then
                e.Cancel = True
            End If
        Else
            ' e.Cancel = True
        End If

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case e.Tool.Key

            Case "Use Customer Address"

                With CreditCardQueue1.UserControlGrid
                    If .ActiveRow IsNot Nothing Then
                        If Not ScreenMode Then Exit Sub
                        .ActiveRow.Cells("CUST_CREDIT_CARD_ADDR1").Value = MyBase.Absx1.txtFor("CUST_ADDR1").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_CITY").Value = MyBase.Absx1.txtFor("CUST_CITY").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_STATE").Value = MyBase.Absx1.txtFor("CUST_STATE").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_ZIP_CODE").Value = MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text
                    End If
                End With
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub CreditCardQueue1_UpdateClickEvent(ByRef rowARTCUST1 As System.Data.DataRow) Handles CreditCardQueue1.UpdateClickEvent
        If SELECTION_NO = 0 Then Exit Sub

        If 1 = 1 Then
            Exit Sub
        End If

        ' Sync changes from control incase the user clicks the Update Menu Option
        rowASFBASE1.Item("CUST_AUTO_CCPA") = rowARTCUST1.Item("CUST_AUTO_CCPA")
        rowASFBASE1.Item("CUST_AUTO_CC_AUTH") = rowARTCUST1.Item("CUST_AUTO_CC_AUTH")
        rowASFBASE1.Item("CUST_AUTO_CC_AUTH_DATE") = rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE")
        rowASFBASE1.Item("CUST_AUTO_CCPA_NOTE") = rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE")
        rowASFBASE1.Item("CUST_AUTO_CC_OPER") = rowARTCUST1.Item("CUST_AUTO_CC_OPER")

        If rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty <> rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty _
            OrElse _
            rowARTCUST1.Item("CUST_AUTO_CC_AUTH", DataRowVersion.Original) & String.Empty <> rowARTCUST1.Item("CUST_AUTO_CC_AUTH") & String.Empty Then

            Try
                Dim Note As String = String.Empty
                ASCMAIN1.Progress("Emailing Rep", "")
                MyBase.Fill_Records("ARTCUST1", rowARTCUST1.Item("CUST_CODE"))

                Note = "Queue: "
                Select Case Val(rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty)
                    Case "0"
                        Note &= "None"
                    Case "1"
                        Note &= "Reminder Queue"
                    Case "2"
                        Note &= "Auto Charge Queue"
                    Case Else
                        Note &= "Unknown"
                End Select

                Note &= Environment.NewLine & Environment.NewLine
                Note &= " Auto Charge Queue Authorization Form: "

                Select Case Val(rowARTCUST1.Item("CUST_AUTO_CC_AUTH") & String.Empty)
                    Case "0"
                        Note &= "No form on file"
                    Case "1"
                        Note &= "Form on file"
                    Case Else
                        Note &= "Unknown"
                End Select

                'Dim objASCNOTE1 As New TAC.ASCNOTE1("CUSTQUEUE", dst, rowARTCUST1.Item("CUST_CODE"))
                'objASCNOTE1.Note = Note.Trim
                'objASCNOTE1.CreateComponents()
                'objASCNOTE1.EmailDocument()

            Catch ex As Exception
                MessageBox.Show("Error while sending email: " & ex.Message, "Send Email", MessageBoxButtons.OK)
            Finally
                ASCMAIN1.Progress(String.Empty, String.Empty)
            End Try

        End If
    End Sub

    Private Sub btnADSCustomerNo_Click(sender As System.Object, e As System.EventArgs) Handles btnADSCustomerNo.Click

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            btnADSCustomerNo.Visible = False
            Exit Sub
        End If
        If MyBase.Absx1.txtFor("CUST_NO_3PL").TextLength > 0 Then
            MessageBox.Show("This customer already has an assigned 3PL Customer Number.", "Assign ADS Customer Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        MyBase.Absx1.txtFor("CUST_NO_3PL").Text = ASCMAIN1.Next_Control_No("ARTCUST1.CUST_NO_3PL")

    End Sub

#End Region


    Private Sub cmdCFG_Click(sender As Object, e As EventArgs) Handles cmdCFG.Click

        If ASCMAIN1.Running_in_VS Then
            If MsgBox("Send Customer Master File to Clarins?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
            End If

            ' Change for ADS 07/16/2025, force developer to supply the LP CODE
            Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC999O1", "AC", "", "CLA")

            MsgBox("File Sent", MsgBoxStyle.OkOnly, "Verification")
        End If


        'ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, SELL_CODE, CUST_STORE_NAME from ARTCUST2 where SELL_CODE is Not Null"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "SELL_CODE,CUST_CODE,CUST_STORE_NO")
        '    Dim SELL_CODE As String = row.Item("SELL_CODE")
        '    Dim CUST_CODE As String = row.Item("CUST_CODE")
        '    Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
        '    Dim CUST_STORE_NAME As String = row.Item("CUST_STORE_NAME")
        '    Dim SCS As String = "S:\INT\ALEX\" & SELL_CODE & "\" & CUST_CODE & "-" & CUST_STORE_NO
        '    SCS = "S:\INT\ALEX2\" & SELL_CODE & "\" & CUST_STORE_NAME
        '    My.Computer.FileSystem.CreateDirectory(SCS)
        'Next
    End Sub


End Class