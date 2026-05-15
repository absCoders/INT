Public Class ICTITEM1
    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim ITEM_ALT_SORT As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")
        SO_PARM_UPC_VENDOR_ID = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""

        With dst
            Create_TDA(.Tables.Add, "TATALRT1", "*")

            ASCMAIN1.sql = "Select * from ICTRETLC where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTRETLC", "**", 0, False, "V")
        End With

        Me.loadFromXls = LoadFromXlsType.Grid
        xls_pending_table_name = "ICTITEMI"
        SplitContainer4.Panel2Collapsed = True

        ASCMAIN1.sql = ""
        ASCMAIN1.Add_Value_List(Absx1.cbeFor("ITEM_ORDR_REL_CODE"), "ICTITEM1.ITEM_ORDR_REL_CODE")

        grdICTRETLC.DataSource = dst.Tables("ICTRETLC")
        grdICTRETLC.Visible = (ASCMAIN1.DBS_COMPANY = "SLP")

        With grdICTRETLC.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTRETLC.DisplayLayout.Bands(0).Columns
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        Next

        Absx1.numFor("ITEM_COST_STD").Focus()
        Absx1.txtFor("ITEM_CODE").Focus()

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            lblITEM_ALT_SORT.Text = "3PL"
            btnUPC.Visible = False
        End If

        lblITEM_CODE_ALT.Visible = (ASCMAIN1.CLIENT = "INT") And (ASCMAIN1.DBS_COMPANY <> "SLP")
        txtITEM_CODE_ALT.Visible = (ASCMAIN1.CLIENT = "INT") And (ASCMAIN1.DBS_COMPANY <> "SLP")
        'SplitContainer2.SplitterDistance = SplitContainer2.Width / 3
        'SplitContainer1.SplitterDistance = SplitContainer1.Width / 2

        grpHazmat.Visible = (ASCMAIN1.CLIENT = "INT")

        Me.lblUnitWtUM.Visible = (ASCMAIN1.DBS_COMPANY = "SLP")
        Me.lblUnitWtUM.Text = "kg"
        Me.lblCaseWtUM.Visible = (ASCMAIN1.DBS_COMPANY = "SLP")
        Me.lblCaseWtUM.Text = "kg"

        Me.lblUnitDimUM.Visible = (ASCMAIN1.DBS_COMPANY = "SLP")
        Me.lblUnitDimUM.Text = "cm"
        Me.lblCaseDimUM.Visible = (ASCMAIN1.DBS_COMPANY = "SLP")
        Me.lblCaseDimUM.Text = "cm"

        Me.lblITEM_CODE_COMPARE_TO.Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")
        Me.lblITEM_CODE_COMPARE_TO_2.Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")
        Me.lblITEM_CODE_COMPARE_TO_3.Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")
        Absx1.txtFor("ITEM_CODE_COMPARE_TO").Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")
        Absx1.txtFor("ITEM_CODE_COMPARE_TO_2").Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")
        Absx1.txtFor("ITEM_CODE_COMPARE_TO_3").Visible = (ASCMAIN1.DBS_COMPANY <> "SLP")

        Create_Lookup("ICTSEAS1")
        Create_Lookup("ICTCOLL1")

        Me.xls_allow_edit = True


    End Sub

#Region "Overrides"

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME

            'Case "ITEM_SUB_VARIETY_CODE"
            '    sql_where = "ITEM_VARIETY_CODE = '" & MyBase.Absx1.txtFor("ITEM_VARIETY_CODE").Text.Trim & "'"
        End Select
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Edit"

            Case "Update"
                If ASCMAIN1.CLIENT = "INT" Then

                End If

                Dim ITEM_ABC_CODE As String = Absx1.txtFor("ITEM_ABC_CODE").Text
                If ITEM_ABC_CODE <> "" Then
                    Dim rowDPTABCP1 As DataRow = LookUp("DPTABCP1", ITEM_ABC_CODE)
                    If rowDPTABCP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid value for ABC Code"
                    End If
                End If

                Dim ITEM_POS_MAX As Decimal = Val(Absx1.numFor("ITEM_POS_MAX").Value & "")
                Dim ITEM_POS_MIN As Decimal = Val(Absx1.numFor("ITEM_POS_MIN").Value & "")

                If ITEM_POS_MIN > ITEM_POS_MAX Then
                    EMsg &= vbCr & "Min Position cannot be greater than Max Position"
                End If

                If Absx1.dteFor("ITEM_PLAN_QUIET_ZONE_DATE").Value & "" <> "" Then

                    ' rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") = "2"
                    If rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" <> "2" Then
                        rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_DATE") = DBNull.Value
                        ' EMsg &= vbCr & "You Must Specify 'Date' if you are going to provide a Do Not Plan Before Date"
                    End If

                Else
                    'If Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & "" = "2" Then
                    '    ' rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") = "0"
                    '    EMsg &= vbCr & "You Must Specify a real date if using the Do Not Plan Before feature"
                    'End If
                End If
                Dim ITEM_PLAN_QUIET_ZONE_TYPE As String = Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & "" ' rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" '  Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & ""
                If ITEM_PLAN_QUIET_ZONE_TYPE = "2" Then
                    If Absx1.dteFor("ITEM_PLAN_QUIET_ZONE_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Specify a real date if using the Do Not Plan Before feature"
                    End If


                End If

                Dim ITEM_SO_QTY_MULT As Int64 = Val(Absx1.numFor("ITEM_SO_QTY_MULT").Value & "")
                Dim ITEM_SO_QTY_MIN As Int64 = Val(Absx1.numFor("ITEM_SO_QTY_MIN").Value & "")
                Dim ITEM_STD_PACK_SLS As Int64 = Val(Absx1.numFor("ITEM_STD_PACK_SLS").Value & "")

                Dim INT_only As Boolean = (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT")

                If ITEM_SO_QTY_MULT = 0 Then
                    If INT_only Then EMsg &= vbCr & "SO Multiple may not be 0"
                    ITEM_SO_QTY_MULT = 1
                End If

                If ITEM_SO_QTY_MIN = 0 Then
                    If INT_only Then EMsg &= vbCr & "SO Minimum may not be 0"
                    ITEM_SO_QTY_MIN = 1
                End If

                If ITEM_STD_PACK_SLS = 0 Then ITEM_STD_PACK_SLS = 1

                If ITEM_SO_QTY_MIN <> 1 Then
                    If ITEM_SO_QTY_MIN Mod ITEM_SO_QTY_MULT <> 0 Then
                        If Absx1.chkFor("ITEM_ALLOW_HALF_PACK").Checked And (ITEM_SO_QTY_MIN * 2) Mod ITEM_SO_QTY_MULT = 0 Then
                            'ALLOWED
                        Else
                            EMsg &= vbCr & "SO Minimum is not congruous with SO Multiple"
                        End If
                    End If
                    If ITEM_SO_QTY_MIN Mod ITEM_STD_PACK_SLS <> 0 Then
                        If Absx1.chkFor("ITEM_ALLOW_HALF_PACK").Checked And (ITEM_SO_QTY_MIN * 2) Mod ITEM_STD_PACK_SLS = 0 Then
                            'ALLOWED
                        Else
                            EMsg &= vbCr & "SO Minimum is not congruous with Standard (Inner) Pack"
                        End If
                    End If

                End If

                If ITEM_SO_QTY_MULT <> 1 And ITEM_SO_QTY_MULT Mod ITEM_STD_PACK_SLS <> 0 Then EMsg &= vbCr & "SO Multiple is not congruous with Standard (Inner) Pack"

                Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(Absx1.numFor("ITEM_SHELF_LIFE_YRS").Value & "")
                Dim ITEM_LOT_CONTROL As String = ITEM_SHELF_LIFE_YRS & ""

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    If Absx1.chkFor("ITEM_LOT_CONTROL").Checked And ITEM_SHELF_LIFE_YRS = 0 Then
                        If INT_only Then EMsg &= vbCr & "Cannot have Lot Control if Shelf Life is 0"
                    End If
                    If Absx1.chkFor("ITEM_ALLOW_HALF_PACK").Checked And (ITEM_SO_QTY_MULT = 0 Or ITEM_SO_QTY_MULT Mod 2 <> 0) Then
                        ' If INT_only Then EMsg &= vbCr & "Cannot have Half Pack unless Inner Pack is >0 and an Even Number"
                        If INT_only Then EMsg &= vbCr & "Cannot have Half Pack unless Order Multiple is >0 and an Even Number"
                    End If
                    If Absx1.chkFor("ITEM_CONTAINS_ALCOHOL").Checked And Not Absx1.chkFor("ITEM_LOT_CONTROL").Checked Then
                        If INT_only Then EMsg &= vbCr & "Cannot set Hazardous (Contains Alcohol) without Lot Control"
                    End If
                End If

                ' check for special characters

                If Absx1.txtFor("ITEM_DESC").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Item Description"
                Else
                    If ASCMAIN1.CLIENT = "INT" Then
                        If ASCMAIN1.DBS_COMPANY = "SLP" Then
                            ' NO CHECKS
                        Else
                            'Dim rx As String = "[^a-zA-Z0-9 .,_$-]" ' Allow Upper/Lower case, numbers, space, dot, comma, underscore and dash
                            'Dim rx As String = "[^a-zA-Z0-9 .,$-]" ' Allow Upper/Lower case, numbers, space, dot, comma and dash
                            Dim rx As String = "[^a-zA-Z0-9 .$]" ' Allow Upper/Lower case, numbers, space, dot
                            Dim r As New System.Text.RegularExpressions.Regex(rx)
                            If r.IsMatch(Absx1.txtFor("ITEM_DESC").Text) Then
                                EMsg &= vbCr & "Item Description has Special Characters which are not allowed"
                            End If
                        End If
                    End If
                End If

                If Absx1.txtFor("COLLECTION_CODE").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Collection Code"
                End If
                If Absx1.txtFor("PROD_CODE").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Product Code"
                End If
                If Absx1.txtFor("COST_CATGY_CODE").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Cost Category Code"
                Else
                    Dim ITEM_SNU_CODE As String = Absx1.optFor("ITEM_SNU_CODE").Value & ""
                    If ITEM_SNU_CODE = "" Then
                        ITEM_SNU_CODE = rowASFBASE1.Item("ITEM_SNU_CODE") & ""
                    End If
                    Select Case ITEM_SNU_CODE
                        Case "S"
                            If Val(Absx1.numFor("ITEM_RETAIL_PRICE").Value & "") = 0 Then
                                EMsg &= vbCr & "Invalid Value for Retail Price for a Saleable Item"
                            End If
                        Case "N", "U"
                            If Val(Absx1.numFor("ITEM_RETAIL_PRICE").Value & "") <> 0 Then
                                EMsg &= vbCr & "Invalid Value for Retail Price for a Non-Saleable Item"
                            End If
                        Case Else
                            EMsg &= vbCr & "Unable to determine SNU"
                    End Select

                End If
                If Absx1.txtFor("ITEM_UOM").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Item Unit of Measure"
                End If

                Dim ITEM_EAN_CODE As String = Trim(Absx1.txtFor("ITEM_EAN_CODE").Text)
                Dim ITEM_UPC_CODE As String = Trim(Absx1.txtFor("ITEM_UPC_CODE").Text)

                If ITEM_EAN_CODE <> "" And ITEM_UPC_CODE <> "" Then
                    EMsg &= vbCr & "An Item Cannot have both a UPC and an EAN"
                Else
                    If ITEM_EAN_CODE <> "" Then
                        Dim rowEAN As DataRow = ASCDATA1.GetDataRow("Select ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE = :PARM1 and ITEM_CODE <> :PARM2", "VV", New Object() {ITEM_EAN_CODE, Absx1.txtFor("ITEM_CODE").Text})
                        If rowEAN IsNot Nothing Then
                            EMsg &= vbCr & "EAN " & ITEM_EAN_CODE & " already used for Item " & rowEAN.Item("ITEM_CODE")
                        End If
                    End If
                    If ITEM_UPC_CODE <> "" Then
                        Dim rowUPC As DataRow = ASCDATA1.GetDataRow("Select ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE = :PARM1 and ITEM_CODE <> :PARM2", "VV", New Object() {ITEM_UPC_CODE, Absx1.txtFor("ITEM_CODE").Text})
                        If rowUPC IsNot Nothing Then
                            EMsg &= vbCr & "UPC " & ITEM_UPC_CODE & " already used for Item " & rowUPC.Item("ITEM_CODE")
                        End If
                    End If
                End If

                If ASCMAIN1.CLIENT = "AHA" Then ' DON'T ENABLE FOR AHA UNLESS WE PRE-VALIDATE THE DATA
                Else
                    If ITEM_EAN_CODE <> "" Then
                        If ITEM_EAN_CODE.Length <> 13 Or Format(Val(ITEM_EAN_CODE), "0000000000000") <> ITEM_EAN_CODE Then
                            EMsg &= vbCr & "EAN " & ITEM_EAN_CODE & " must be 13 numeric digits"
                        End If
                    End If
                    If ITEM_UPC_CODE <> "" Then
                        If ITEM_UPC_CODE.Length <> 12 Or Format(Val(ITEM_UPC_CODE), "000000000000") <> ITEM_UPC_CODE Then
                            EMsg &= vbCr & "UPC " & ITEM_UPC_CODE & " must be 12 numeric digits"
                        End If
                    End If
                End If


                If Absx1.txtFor("ITEM_CODE_COMPARE_TO").Text <> "" Then
                    If LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE_COMPARE_TO").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Item Code Specified for Compare-To Item"
                    End If
                End If

                If Absx1.txtFor("ITEM_CODE_COMPARE_TO_2").Text <> "" Then
                    If LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE_COMPARE_TO_2").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Item Code Specified for Compare-To Item 2"
                    End If
                End If

                If Absx1.txtFor("ITEM_CODE_COMPARE_TO_3").Text <> "" Then
                    If LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE_COMPARE_TO_3").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Item Code Specified for Compare-To Item 3"
                    End If
                End If

                Dim ITEM_BUFFER_QTY As Int32 = Val(Absx1.numFor("ITEM_BUFFER_QTY").Value & "")
                Dim ITEM_BUFFER_PCT As Int32 = Val(Absx1.numFor("ITEM_BUFFER_PCT").Value & "")
                If ITEM_BUFFER_QTY <> 0 And ITEM_BUFFER_PCT <> 0 Then
                    EMsg &= vbCr & "Cannot Specify both Buffer Qty and Pct"
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.txtFor("ITEM_ALT_SORT").Text = "" Then
                        If Absx1.txtFor("ITEM_EAN_CODE").Text <> "" Then
                            rowASFBASE1.Item("ITEM_ALT_SORT") = Mid(Absx1.txtFor("ITEM_EAN_CODE").Text, 7, 6)
                        Else

                        End If
                    End If

                    Dim ITEM_ALT_SORT As String = txtITEM_ALT_SORT.Text

                    If ASCMAIN1.DBS_COMPANY = "SLP" Then
                    Else

                        If ITEM_ALT_SORT = "" Or ITEM_ALT_SORT.Contains(" ") Then
                            EMsg &= vbCr & "3PL Item Code is mandatory, and may not contain blanks"
                        Else
                            '/^[a-z0-9]+$/i
                            '/^([a-zA-Z0-9 _-]+)$/
                            Dim rx As String = "[^A-Z0-9]" ' Allow Upper case, numbers

                            Dim r As New System.Text.RegularExpressions.Regex(rx)
                            If r.IsMatch(ITEM_ALT_SORT) Then
                                EMsg &= vbCr & "3PL Item Code has Special Characters which are not allowed"
                            End If

                            'If EntryMode = "New" Then
                            Dim COLLECTION_CODE As String = Absx1.txtFor("COLLECTION_CODE").Text
                            Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
                            ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1,ICTCOLL1" & vbCrLf _
                                & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                                & "   and ICTITEM1.ITEM_CODE <> '" & ITEM_CODE & "'" & vbCrLf _
                                & "   and ICTCOLL1.BRAND_CODE = " & vbCrLf _
                                & " (Select BRAND_CODE from ICTCOLL1 where COLLECTION_CODE = '" & COLLECTION_CODE & "')" & vbCrLf _
                                & "   and ICTITEM1.ITEM_ALT_SORT = '" & ITEM_ALT_SORT & "'"

                            ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1" & vbCrLf _
                                & " where ICTITEM1.ITEM_CODE <> '" & ITEM_CODE & "'" & vbCrLf _
                                & "   and ICTITEM1.ITEM_ALT_SORT = '" & ITEM_ALT_SORT & "'"

                            Dim row As DataRow = ASCDATA1.GetDataRow
                            If row IsNot Nothing Then
                                EMsg &= vbCr & "3PL Item Code " & ITEM_ALT_SORT & " already on file for Item " & row.Item("ITEM_CODE")
                            End If
                            'End If
                        End If

                    End If


                    If ASCMAIN1.CLIENT = "INT" Then
                        Dim ITEM_CODE_ALT As String = txtITEM_CODE_ALT.Text
                        If ITEM_CODE_ALT = "" Then
                            ' this is ok
                        ElseIf ITEM_CODE_ALT = Absx1.txtFor("ITEM_CODE").Text Then
                            EMsg &= vbCr & "Component Alternate Item Code " & ITEM_CODE_ALT & " may not be same as Item Code"
                        Else
                            Dim row As DataRow = LookUp("ICTITEM1", ITEM_CODE_ALT)
                            If row Is Nothing Then
                                EMsg &= vbCr & "Invalid Alternate Item Code " & ITEM_CODE_ALT & " - must be an actual, existing item"
                            End If
                        End If
                    End If

                    If EntryMode <> "Edit" Then
                        ' Cannot change EAN or ITEM_ALT_SORT
                    End If

                    If Absx1.txtFor("ITEM_CLASS_CODE").Text = "" Then
                        EMsg &= vbCr & "You must enter a value for Item Class"
                    Else
                        If LookUp("ICTCLAS1", Absx1.txtFor("ITEM_CLASS_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value entered for Item Class"
                        End If
                    End If

                    If Absx1.txtFor("ITEM_TYPE_CODE").Text = "" Then
                        EMsg &= vbCr & "You must enter a value for Item Type"
                    Else
                        If LookUp("ICTTYPE1", Absx1.txtFor("ITEM_TYPE_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value entered for Item Type"
                        End If
                    End If
                End If

                If ASCMAIN1.CLIENT = "INT" AndAlso EMsg.Length = 0 Then
                    Dim warningMessage As String = String.Empty

                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(Absx1.txtFor("ITEM_CODE").Text)

                    If (rowICTITEM1.Item("HMAT_CODE") & String.Empty).ToString.Trim.Length > 0 AndAlso rowICTITEM1.Item("ITEM_CONTAINS_ALCOHOL") & String.Empty <> "1" Then
                        warningMessage &= vbCr & "A Hazardous Material Code was provided; however, 'Item Contains Alcohol (HAZMAT)' is not checked."
                    ElseIf (rowICTITEM1.Item("HMAT_CODE") & String.Empty).ToString.Trim.Length = 0 AndAlso rowICTITEM1.Item("ITEM_CONTAINS_ALCOHOL") & String.Empty = "1" Then
                        warningMessage &= vbCr & "A Hazardous Material Code was NOT provided; however, 'Item Contains Alcohol (HAZMAT)' is checked."
                    End If

                    If (rowICTITEM1.Item("COMMODITY_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                        warningMessage &= vbCr & "You did not provide a Commodity Code."
                    End If

                    If upload_XLS_mode = "" AndAlso warningMessage.Length > 0 Then
                        warningMessage &= vbCr & vbCr & "Update Anyway?"
                        Select Case MessageBox.Show(warningMessage, "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                            Case Windows.Forms.DialogResult.Yes

                            Case Windows.Forms.DialogResult.No
                                EMsg &= vbCr & "User cancelled Update."

                        End Select
                    End If
                End If

                If EMsg = "" Then
                    If EntryMode = "New" Then
                        Dim PROD_CODE As String = Absx1.txtFor("PROD_CODE").Text
                        If PROD_CODE <> "" Then
                            Dim rowICTPROD1 As DataRow = LookUp("ICTPROD1", PROD_CODE)
                            If rowICTPROD1 IsNot Nothing Then
                                Dim COST_CATGY_CODE As String = rowICTPROD1.Item("COST_CATGY_CODE") & ""
                                Dim PROD_MAX_POS As Decimal = Val(rowICTPROD1.Item("PROD_MAX_POS") & "")
                                Dim PROD_MIN_POS As Decimal = Val(rowICTPROD1.Item("PROD_MIN_POS") & "")
                                Dim PROD_MIN_DAYS_SUPPLY As Decimal = Val(rowICTPROD1.Item("PROD_MIN_DAYS_SUPPLY") & "")
                                If COST_CATGY_CODE = "N" AndAlso PROD_MIN_DAYS_SUPPLY >= 0 And PROD_MAX_POS > 0 And PROD_MIN_POS > 0 And PROD_MAX_POS > PROD_MIN_POS Then
                                    rowASFBASE1.Item("ITEM_POS_MAX") = PROD_MAX_POS
                                    rowASFBASE1.Item("ITEM_POS_MIN") = PROD_MIN_POS
                                    rowASFBASE1.Item("ITEM_MIN_DAYS_SUPPLY") = PROD_MIN_DAYS_SUPPLY
                                End If
                            End If
                        End If
                    End If

                    If EntryMode = "Edit" Then

                        If ASCMAIN1.USER_ID = "wjz" Then
                            ' enable this section if ASCMAIN1.USER_ID = 'wjz'
                            ' need to create J/Es for Albina - see email 07/13/2024 for an example
                        Else
                            If rowASFBASE1.Item("COST_CATGY_CODE") & "" <> rowASFBASE1.Item("COST_CATGY_CODE", DataRowVersion.Original) & "" Then
                                EMsg &= vbCr & "Cannot Change the Cost Category of an Item using Item Master File Maintenance."
                            End If
                        End If

                    End If
                End If

                ' CHECK SALES DIVISION
                'If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                'Else
                '    Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text)
                '    Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", rowICTCOLL1.Item("BRAND_CODE") & "")
                '    Absx1.txtFor("COLLECTION_CODE").Text = rowICTBRAN1.Item("SALES_DIVISION_CODE")

                'End If
        End Select
    End Sub


    Overrides Sub Proceed_Update_Special_Pre()
        If EntryMode = "New" Then
            rowASFBASE1.Item("ITEM_COST_STATUS") = "P"

        End If

        Dim rowICTTYPE1 As DataRow = LookUp("ICTTYPE1", Absx1.txtFor("ITEM_TYPE_CODE").Text)
        If rowICTTYPE1 IsNot Nothing Then
            rowASFBASE1.Item("ITEM_PLAN_WASTE_PCT") = rowICTTYPE1.Item("ITEM_WASTE_PCT")
        End If

        If EntryMode = "Edit" Then
            If rowASFBASE1.Item("ITEM_ALT_SORT") & "" <> ITEM_ALT_SORT Then
                Throw New Exception("Data Loss - please call ABS (do not click OK)")
            End If
        End If

        If EntryMode = "Edit" Then

            Dim ITEM_RETAIL_PRICE_OLD As Decimal = Val(rowASFBASE1.Item("ITEM_RETAIL_PRICE", DataRowVersion.Original) & "")
            Dim ITEM_RETAIL_PRICE_NEW As Decimal = Val(rowASFBASE1.Item("ITEM_RETAIL_PRICE", DataRowVersion.Current) & "")

            If ITEM_RETAIL_PRICE_OLD <> ITEM_RETAIL_PRICE_NEW Then
                Dim ITEM_CODE As String = rowASFBASE1.Item("ITEM_CODE")
                If ASCMAIN1.Running_in_VS Then Stop
                Dim ALERT_MESSAGE As String = ""
                Dim ALERT_SUBJECT As String = ""
                Dim ALERT_EMAIL = ROWs("SOTPARM1").Item("SO_PARM_EML_PRICE") & ""

                Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
                With rowTATALRT1
                    Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
                    .Item("ALERT_NO") = ALERT_NO
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("FORM_NAME") = "ICTITEM1"
                    .Item("FORM_KEY") = ALERT_NO
                    .Item("ALERT_EMAIL") = ALERT_EMAIL
                    .Item("ALERT_EML") = "1"

                    .Item("ALERT_EML_DATE") = DATETIME_STAMP
                    ALERT_SUBJECT = "Retail Price Change for Item " & ITEM_CODE
                    .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
                    ALERT_MESSAGE = "Control No: " & ALERT_NO & vbCrLf & $"Retail Price Change for Item {ITEM_CODE} from {Format(ITEM_RETAIL_PRICE_OLD, "###.00")} To {Format(ITEM_RETAIL_PRICE_NEW, "###.00")}"
                    .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
                End With
                dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                EMAIL_ADDRESSs.Add(ALERT_EMAIL, "Price Auditor")

                Dim SEND_NO As String = ""
                If ASCMAIN1.Running_in_VS Then
                    SEND_NO = "TESTING"
                    Stop
                Else
                    SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                    ALERT_SUBJECT, "PC_PRCEXC", True, False, ITEM_CODE, ITEM_CODE, "Price List Code", ALERT_MESSAGE)
                End If

                rowTATALRT1.Item("SEND_NO") = SEND_NO
                Update_Record_TDA("TATALRT1")

                'TAC.TACMAIN1.Record_Event("SOTPRIC1", PRICE_LIST_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "PRCEXC", "Price Exception Alert emailed", SEND_NO, "SOTPRIC1")
            End If
        End If

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        If EntryMode = "New" Then
            rowASFBASE1.Item("ITEM_UOM") = "EA"
            rowASFBASE1.Item("ITEM_STATUS") = "A"
            rowASFBASE1.Item("ITEM_STATUS_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            rowASFBASE1.Item("ITEM_CATGY_CODE") = "E"

            For Each COLUMN_NAME As String In New String() {"ITEM_UPC_CODE", "ITEM_COST_STATUS", "ITEM_COST_MAKE_BUY", "ITEM_COST_CURR_CODE", "ITEM_COST_FRT_CLASS", "ITEM_COST_STD",
"ITEM_COST_STD_FUTURE", "ITEM_YYYYPP_CUR_COST", "ITEM_YYYYPP_PRV_COST", "ITEM_COST_WASTE_PCT", "ITEM_PLAN_WASTE_PCT",
"ITEM_RSV_PCT", "ITEM_ORDR_REL_CODE", "ITEM_EAN_CODE", "INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE", "ORDR_REL_BATCH_NO",
"XMIT_BATCH_3PL", "XMIT_DATE_3PL", "ITEM_PICTURE_FILENAME", "ITEM_SPEC_NO", "ITEM_CAT_CREATED", "ITEM_CAT_CHANGED",
"ITEM_CAT_DELETED", "ITEM_NEW_LIST_PRICE", "ITEM_NEW_LIST_PRICE_DATE", "ITEM_ALT_SORT"}
                rowASFBASE1.Item(COLUMN_NAME) = DBNull.Value
            Next
        End If

        Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text

        If EntryMode = "Edit" Then
            Dim rowCheck As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            ITEM_ALT_SORT = rowCheck.Item("ITEM_ALT_SORT") & ""
        End If

        Dim IMAGE_NAME As String = ITEM_CODE

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imgba() As Byte = Nothing
        picItemImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            If EntryMode = "Edit" Then
                Absx1.txtFor("ITEM_ALT_SORT").Enabled = False
                Absx1.txtFor("ITEM_EAN_CODE").Enabled = True ' LBM says she will be careful ' False
            Else
                Absx1.txtFor("ITEM_ALT_SORT").Enabled = True
                Absx1.txtFor("ITEM_EAN_CODE").Enabled = True
            End If
        End If

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            cmdSendToCUSA.Visible = False
            cmdSendItemToCUSA.Visible = True
        End If

        If ASCMAIN1.CLIENT = "SLP" Then
            EnforceConstraints(False)
            Fill_Records("ICTRETLC", New String() {ITEM_CODE})
            grdICTRETLC.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdICTRETLC.DisplayLayout.Bands(0).SortedColumns.Add("CURR_CODE", False, True)
            grdICTRETLC.DisplayLayout.Bands(0).SortedColumns.Add("OPS_YYYYPP", True)
            'Sort_grdColumns(grdICTRETLC, "CURR_CODE," & "OPS_YYYYPP".ToLower)

            EnforceConstraints(True)
            'grdICTRETLC.Text = ""
        End If

        Set_ABC_Parameters_ReadOnly()

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ICTRETLC").Rows.Clear()
            EnforceConstraints(True)
        End If

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            cmdSendToCUSA.Visible = True
            cmdSendItemToCUSA.Visible = False
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTRETLC.Enabled = tf
        If tf Then
            Set_ABC_Parameters_ReadOnly()
        End If
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        grdXLS_PENDING.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        grdXLS_PENDING.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        'With grdARTCUSTX.DisplayLayout.Override
        '    If EntryMode = "New" Or EntryMode = "Edit" Then
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        '        .AllowUpdate = DefaultableBoolean.True
        '        .AllowDelete = DefaultableBoolean.True
        '        cmdCustomers.Visible = True
        '    Else
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '        .AllowUpdate = DefaultableBoolean.False
        '        .AllowDelete = DefaultableBoolean.False

        '        cmdCustomers.Visible = False
        '    End If
        'End With
        Set_Read_Only_for_ctl(Absx1.optFor("ITEM_SNU_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.dteFor("ITEM_STATUS_DATE"), Not tf Or EntryMode <> "Edit")
        Set_Read_Only(grpCosting, True)

        btnUPC.Visible = ScreenMode And (EntryMode = "New" Or EntryMode = "Edit")
        If ASCMAIN1.CLIENT = "INT" Then
            btnUPC.Visible = False
        End If

        If ASCMAIN1.CLIENT = "INT" Then
            If EntryMode = "Edit" Then
                If ASCMAIN1.USER_SECURITY_CODEs.Contains("RT") Then
                Else
                    Set_Read_Only_for_ctl(Absx1.numFor("ITEM_RETAIL_PRICE"), True)
                End If

                Set_Read_Only_for_ctl(Absx1.numFor("ITEM_NEW_RETAIL_PRICE"), True)
                Set_Read_Only_for_ctl(Absx1.dteFor("ITEM_NEW_RETAIL_PRICE_DATE"), True)
            End If
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "CUST_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then

        '        End If

        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "PROD_CODE"
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    If Absx1.optFor("ITEM_BASIC_PROMO").Value & "" = "" Then
                        Dim row As DataRow = LookUp("ICTPROD1", Absx1.txtFor("PROD_CODE").Text)
                        If row IsNot Nothing Then
                            Dim PROD_BASIC_PROMO As String = row.Item("PROD_BASIC_PROMO") & ""
                            If PROD_BASIC_PROMO = "B" Or PROD_BASIC_PROMO = "P" Then
                                Absx1.optFor("ITEM_BASIC_PROMO").Value = PROD_BASIC_PROMO
                            End If
                        End If
                    End If
                End If

            Case "COLLECTION_CODE"
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text)
                        If rowICTCOLL1 IsNot Nothing AndAlso rowICTCOLL1.Item("BRAND_CODE") & "" <> "" Then
                            Dim row As DataRow = LookUp("ICTBRAN1", rowICTCOLL1.Item("BRAND_CODE") & "")
                            If row IsNot Nothing Then
                                Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
                                If SALES_DIVISION_CODE <> "" Then
                                    Absx1.txtFor("SALES_DIVISION_CODE").Text = SALES_DIVISION_CODE
                                End If
                            End If
                        End If
                    End If
                End If
        End Select
    End Sub

    Sub Set_ABC_Parameter_Values(ITEM_ABC_CODE As String)
        If Not Absx1.chkFor("ITEM_ABC_PARMS_LOCKED").Checked Then
            If ITEM_ABC_CODE <> "" Then
                Dim row As DataRow = LookUp("DPTABCP1", ITEM_ABC_CODE)
                If row IsNot Nothing Then
                    Dim ABC_MAX_POS As Decimal = Val(row.Item("ABC_MAX_POS") & "")
                    Dim ABC_MIN_POS As Decimal = Val(row.Item("ABC_MIN_POS") & "")
                    Dim ABC_MIN_DAYS_SUPPLY As Integer = Val(row.Item("ABC_MIN_DAYS_SUPPLY") & "")
                    Absx1.numFor("ITEM_POS_MAX").Value = ABC_MAX_POS
                    Absx1.numFor("ITEM_POS_MIN").Value = ABC_MIN_POS
                    Absx1.numFor("ITEM_MIN_DAYS_SUPPLY").Value = ABC_MIN_DAYS_SUPPLY
                End If
            End If
        End If
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)

            Case "ITEM_ABC_CODE"
                Dim ITEM_ABC_CODE As String = Absx1.txtFor("ITEM_ABC_CODE").Text

                If EntryMode = "Edit" Then
                    Set_ABC_Parameter_Values(ITEM_ABC_CODE)
                End If

                'Case "PROD_CODE"
                '    If EntryMode = "Edit" Then
                '        If Absx1.optFor("ITEM_BASIC_PROMO").Value & "" = "" Then
                '            Dim row As DataRow = LookUp("ICTPROD1", Absx1.txtFor("PROD_CODE").Text)
                '            If row IsNot Nothing Then
                '                Dim PROD_BASIC_PROMO As String = row.Item("PROD_BASIC_PROMO") & ""
                '                If PROD_BASIC_PROMO = "B" Or PROD_BASIC_PROMO = "P" Then
                '                    Absx1.optFor("ITEM_BASIC_PROMO").Value = PROD_BASIC_PROMO
                '                End If
                '            End If
                '        End If
                '    End If

                'Case "COLLECTION_CODE"
                '    If EntryMode = "Edit" Then
                '        If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                '            Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text)
                '            If rowICTCOLL1 IsNot Nothing AndAlso rowICTCOLL1.Item("BRAND_CODE") & "" <> "" Then
                '                Dim row As DataRow = LookUp("ICTBRAN1", rowICTCOLL1.Item("BRAND_CODE") & "")
                '                If row IsNot Nothing Then
                '                    Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
                '                    If SALES_DIVISION_CODE <> "" Then
                '                        Absx1.txtFor("SALES_DIVISION_CODE").Text = SALES_DIVISION_CODE
                '                    End If
                '                End If
                '            End If
                '        End If
                '    End If
        End Select

    End Sub

    'Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    'Select Case Absx1.GetABSColumnName(txtctl)
    '    Case "CUST_CODE"

    '    Case "ORDR_NO"

    'End Select
    'End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        'Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        'Select Case COLUMN_NAME
        '    Case "ORDR_ADDR_TYPE_ST"
        '        If Not Me.IsLoading Then

        '        End If
        'End Select
    End Sub

    Public Overrides Sub chk_CheckedChanged(sender As Object, e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ITEM_ABC_PARMS_LOCKED"
                If Not Me.IsLoading Then
                    Set_ABC_Parameters_ReadOnly()
                    If ScreenMode Then
                        Set_ABC_Parameter_Values(Absx1.txtFor("ITEM_ABC_CODE").Text)
                    End If
                End If
        End Select
    End Sub

    Sub Set_ABC_Parameters_ReadOnly()
        Dim TF As Boolean = Not Absx1.chkFor("ITEM_ABC_PARMS_LOCKED").Checked
        Absx1.numFor("ITEM_POS_MAX").ReadOnly = TF
        Absx1.numFor("ITEM_POS_MIN").ReadOnly = TF
        Absx1.numFor("ITEM_MIN_DAYS_SUPPLY").ReadOnly = TF
    End Sub
#End Region

    Private Sub btnUPC_Click(sender As System.Object, e As System.EventArgs) Handles btnUPC.Click
        If SO_PARM_UPC_VENDOR_ID = "" Then
            MsgBox("No UPC Vendor Prefix (see SO Paramters Table)", MsgBoxStyle.OkOnly, "Cannot Generate a UPC")
            Exit Sub
        End If
        If Absx1.txtFor("ITEM_UPC_CODE").Text <> "" Then
            If MsgBox("UPC Already Exists for this Item - Are you Sure that you want to Replace it?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
        End If
        Dim UPC_CODE As String = ASCMAIN1.Next_Control_No("ICTITEM1.ITEM_UPC_CODE")
        Absx1.txtFor("ITEM_UPC_CODE").Text = TAC.SOCMAIN1.UPC(Me, UPC_CODE, SO_PARM_UPC_VENDOR_ID, True)
    End Sub

    Private Sub cmdSendToCUSA_Click(sender As Object, e As EventArgs) Handles cmdSendToCUSA.Click

        ' Need to prompt the user what 3PL we are to send to.
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LP_CODE", "WHTTPLP1", "", "")
        Dim LP_CODE As String = String.Empty

        Dim F As New ASFCODE1
        F.ShowDialog()
        F.Dispose()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            LP_CODE = ASCMAIN1.CodeSelector.SelectedCode
        Else
            Exit Sub
        End If

        If MsgBox($"Send Item Master File to {LP_CODE}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        ' Change for ADS 07/16/2025, force developer to supply the LP CODE
        Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "All", LP_CODE)

        MsgBox("File Sent", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Private Sub cmdSendItemToCUSA_Click(sender As Object, e As EventArgs) Handles cmdSendItemToCUSA.Click
        ' Need to prompt the user what 3PL we are to send to.
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LP_CODE", "WHTTPLP1", "", "")
        Dim LP_CODE As String = String.Empty

        Dim F As New ASFCODE1
        F.ShowDialog()
        F.Dispose()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            LP_CODE = ASCMAIN1.CodeSelector.SelectedCode
        Else
            Exit Sub
        End If

        If MsgBox($"Send Item {Absx1.txtFor("ITEM_CODE").Text} to 3PL {LP_CODE}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        Select Case LP_CODE
            Case "ADS"
                ' Change for ADS 07/16/2025, force developer to supply the LP CODE
                Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", $"IT:{Absx1.txtFor("ITEM_CODE").Text}", LP_CODE)
                MsgBox("Item Sent", MsgBoxStyle.OkOnly, "Verification")

            Case "CLA"
                Dim ITEM_ALT_SORT As String = txtITEM_ALT_SORT.Text
                If MsgBox("Forcibly Include 3PL Item " & ITEM_ALT_SORT & " on Next Send to Clarins?" _
                  & vbCrLf & vbCrLf & "Note - item must still meet normal criteria to be eligible for send",
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
                ASCDATA1.ExecuteSQL("Update CONV.CFG_ITMST Set ITDES2 = '.' where ITITEM = '" & ITEM_ALT_SORT & "'")
                MsgBox("Item has been queued for next Send", MsgBoxStyle.OkOnly, "Verification")
        End Select


    End Sub



#Region "Excel Import"

    Protected Overrides Sub InitializeXlsImport_Special()
        ASCMAIN1.sql = "SELECT ITEM_CODE,ITEM_DESC,NUM_CHARS_DUMMY, ITEM_EAN_CODE,ITEM_RETAIL_PRICE,ITEM_COST_STD,LAUNCH_DATE,PROD_CODE,
COLLECTION_CODE,COLLECTION_DESC,ITEM_TYPE_CODE,ITEM_HIDE_FROM_CAT,ITEM_CLASS_CODE,HC_CODE,COST_CATGY_CODE,
SEASON_CODE,APPR_SAMPLE_DUMMY, ITEM_BASIC_PROMO,COUNTRY_CODE,ITEM_CONTAINS_ALCOHOL,ITEM_LOT_CTL,ITEM_WEIGHT_CHECK,ITEM_APPR_1ST_REC,ITEM_CRITICAL_TO_SHIP,
ITEM_SHELF_LIFE_YRS,VEND_CODE FROM ICTITEMI"
        Create_TDA(dst.Tables.Add, xls_pending_table_name, "**", 0,, "")

        'grdICTITEMI.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        'grdICTITEMI.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show


        'TODO: rename grdICTITEMI in ASFCODEM
        grdXLS_PENDING.DataSource = dst.Tables(xls_pending_table_name)
        ASCMAIN1.grdInitializeLayout(grdXLS_PENDING)
        grdXLS_PENDING.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        grdXLS_PENDING.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        Create_Summary(grdXLS_PENDING, "ITEM_CODE", "Count")

        Create_TDA(dst.Tables.Add, "ICTITEMT", "*", 0, True)

        UnsubscribeBeforeRowUpdate(grdXLS_PENDING)


        Fill_Records(xls_pending_table_name, ClearBeforeFilling:=True)
        Fill_Records("ICTITEMT", ClearBeforeFilling:=True)

    End Sub

    Protected Overrides Function LoadXlsWorksheetForGridImport_Special(ByVal FILENAME As String) As DataTable
        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing


        oWB = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        oSheet = oWB.Sheets(0)

        Dim usedCells As SpreadsheetGear.IRange = oSheet.UsedRange
        Dim rowCount = usedCells.RowCount

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE ICTITEMI")
        dst.Tables(xls_pending_table_name).Rows.Clear()

        Dim dtSheetData As DataTable = oSheet.UsedRange.Cells($"A2:AD{rowCount}").GetDataTable(SpreadsheetGear.Data.GetDataFlags.NoColumnHeaders Or SpreadsheetGear.Data.GetDataFlags.FormattedText)
        dtSheetData.Columns.Add("EXCEL_ROW", GetType(Integer))
        oSheet.UsedRange.Cells($"I2:I{rowCount}").NumberFormat = "m/d/yyyy"

        For i As Integer = dtSheetData.Rows.Count - 1 To 0 Step -1
            Dim dr As DataRow = dtSheetData.Rows(i)
            dr("EXCEL_ROW") = i + 2
            If (dr.Item(0) & "").ToString().ToUpper() <> "YES" Then
                dr.Delete()
            End If
        Next
        dtSheetData.Columns.RemoveAt(0)
        dtSheetData.Columns.RemoveAt(0)


        Return dtSheetData
    End Function


    Protected Overrides Sub ValidateImportedXls_Special(ByVal dr As DataRow, ByVal excelRow As Integer, dtErrorRows As DataTable)

        'MaxLength
        For i As Integer = 0 To dst.Tables("ICTITEMT").Columns.Count - 1

            Dim col As DataColumn = dst.Tables("ICTITEMT").Columns(i)

            Dim colName As String = col.ColumnName
            If colName = "ITEM_UPC_CODE" Then Continue For
            Dim colValue As String = dr(colName) & ""
            If String.IsNullOrEmpty(colValue) Then Continue For


            'Convert lookups and booleans first

            'Boolean columns
            If {"ITEM_HIDE_FROM_CAT", "ITEM_LOT_CTL", "ITEM_WEIGHT_CHECK", "ITEM_APPR_1ST_REC", "ITEM_CRITICAL_TO_SHIP", "ITEM_CONTAINS_ALCOHOL"}.Contains(colName) Then
                'Convert value
                Select Case colValue.ToString.Trim().ToUpper()
                    Case "YES"
                        colValue = "1"
                    Case "NO"
                        colValue = "0"
                End Select

                If Not {"1", "0"}.Contains(colValue) Then
                    dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided -- must be ""Yes"" or ""No""")
                End If
                Continue For
            End If

            If colName = "COST_CATGY_CODE" Then
                Continue For
                'If Not {"S", "UN", "NC"}.Contains(colValue) Then
                '    dtErrorRows.Rows.Add(dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided")
                'End If
                'Continue For
            End If

            If colName = "ITEM_BASIC_PROMO" Then
                If Not {"BASIC", "PROMO"}.Contains(colValue.ToUpper()) Then
                    dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided")
                End If
                Continue For
            End If

            If {"ITEM_EAN_CODE"}.Contains(colName) Then
                Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(colValue, "^\d{12,13}$")
                If Not match.Success Then
                    dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided for EAN/UPC (must be 12 or 13 digits)")
                End If
                Continue For
            End If

            'Lookup Values
            Dim tableLookups As Dictionary(Of String, String) = New Dictionary(Of String, String) From {{"SEASON_CODE", "ICTSEAS1"}, {"COLLECTION_CODE", "ICTCOLL1"}}
            If tableLookups.Keys.Contains(colName) Then

                If Not String.IsNullOrEmpty(colValue) Then
                    Dim lookupRow As DataRow = LookUp(tableLookups(colName), colValue)
                    If lookupRow Is Nothing Then
                        'Invalid value provided
                        dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided")
                    End If
                End If
                Continue For
            End If

            'Custom Lookups
            Dim customLookups As Dictionary(Of String, String) = New Dictionary(Of String, String) From {{"PROD_CODE", "SELECT PROD_CODE,PROD_DESC FROM ICTPROD1 WHERE PROD_DESC=:PARM1"},
                {"ITEM_TYPE_CODE", "SELECT ITEM_TYPE_CODE,ITEM_TYPE_DESC FROM ICTTYPE1 WHERE ITEM_TYPE_DESC=:PARM1"},
                {"ITEM_CLASS_CODE", "SELECT ITEM_CLASS_DESC,ITEM_CLASS_DESC FROM ICTCLAS1 WHERE ITEM_CLASS_DESC=:PARM1"},
                {"COUNTRY_CODE", "SELECT COUNTRY_CODE,COUNTRY_NAME FROM TATCNTRY WHERE COUNTRY_NAME=:PARM1"},
                {"VEND_CODE", "SELECT VEND_CODE,VEND_NAME FROM APTVEND1 WHERE VEND_NAME=:PARM1"}}
            If customLookups.Keys.Contains(colName) Then

                If Not String.IsNullOrEmpty(colValue) Then
                    Dim lookupRow As DataRow = ASCDATA1.GetDataRow(customLookups(colName), "V", colValue)

                    If lookupRow Is Nothing Then
                        'Invalid value provided
                        dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, $"Invalid value provided")
                    End If
                End If
                Continue For
            End If


            If {"LAUNCH_DATE"}.Contains(colName) Then
                CanConvert(Of Date)(dr, excelRow, colName, colValue, dtErrorRows, "Invalid date format")
                Continue For
            End If

            If {"ITEM_COST_STD", "ITEM_RETAIL_PRICE"}.Contains(colName) Then
                CanConvert(Of Decimal)(dr, excelRow, colName, colValue, dtErrorRows, "Invalid number format")
                Continue For
            End If

            If {"ITEM_SHELF_LIFE_YRS", "ITEM_SO_QTY_MIN", "ITEM_SO_QTY_MULT"}.Contains(colName) Then
                CanConvert(Of Integer)(dr, excelRow, colName, colValue, dtErrorRows, "Invalid number format")
                Continue For
            End If

            If col.DataType Is Type.GetType("System.String") Then
                If colValue.Length > col.MaxLength Then
                    dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, dr(colName), $"Data exceeds max length ({col.MaxLength}), Current: {colValue.ToString().Length} ")
                End If
            End If

            'DataType
            Dim expectedType As Type = col.DataType
            Dim actualValue As Object = dr.Item(i)

            If actualValue IsNot DBNull.Value AndAlso actualValue.GetType() IsNot expectedType Then
                dtErrorRows.Rows.Add(
                    excelRow,
                    dr.Item("ITEM_CODE"),
                    col.ColumnName,
                    dr(col.ColumnName),
                    $"Data type mismatch (Expected: {expectedType.Name}, Found: {actualValue.GetType().Name})"
                )
            End If
        Next

        'Lookup

    End Sub


    Private Sub CanConvert(Of T)(ByVal dr As DataRow, ByVal excelRow As Integer, ByVal colName As String, ByVal colValue As String, ByRef dtErrorRows As DataTable, ByVal errorMessage As String)
        Dim convertedValue As T

        colValue = colValue.ToUpper().Replace("$", "").Replace("N/A", "")

        If Not String.IsNullOrEmpty(colValue) Then
            Dim isValid As Boolean = False

            If GetType(T) Is GetType(Date) Then
                Dim temp As Date
                isValid = Date.TryParse(colValue, temp)
            ElseIf GetType(T) Is GetType(Decimal) Then
                Dim temp As Decimal
                isValid = Decimal.TryParse(colValue, temp)
            ElseIf GetType(T) Is GetType(Integer) Then
                Dim temp As Integer
                isValid = Integer.TryParse(colValue, temp)
            End If

            If Not isValid Then
                dtErrorRows.Rows.Add(excelRow, dr.Item("ITEM_CODE"), colName, colValue, errorMessage)
            End If
        End If
    End Sub



    Private Sub btnFinalize_Click(sender As Object, e As EventArgs) Handles btnFinalize.Click

        Try
            Update_Record_TDA(xls_pending_table_name)
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        ASCDATA1.ExecuteSQL("DELETE FROM ICTITEMT")
        ASCDATA1.ExecuteSQL("INSERT INTO ICTITEMT (ITEM_CODE,ITEM_DESC,ITEM_EAN_CODE,ITEM_UPC_CODE,ITEM_RETAIL_PRICE,ITEM_COST_STD,LAUNCH_DATE,PROD_CODE,COLLECTION_CODE,ITEM_TYPE_CODE,ITEM_HIDE_FROM_CAT,ITEM_CLASS_CODE,HC_CODE,COST_CATGY_CODE,SEASON_CODE,ITEM_BASIC_PROMO,COUNTRY_CODE,ITEM_CONTAINS_ALCOHOL,ITEM_LOT_CTL,ITEM_WEIGHT_CHECK,ITEM_APPR_1ST_REC,ITEM_CRITICAL_TO_SHIP,ITEM_SHELF_LIFE_YRS,VEND_CODE,ITEM_SO_QTY_MIN,ITEM_SO_QTY_MULT)
SELECT ITEM_CODE,ITEM_DESC,
CASE WHEN LENGTH(ITEM_EAN_CODE) = 13 THEN ITEM_EAN_CODE END ITEM_EAN_CODE,
CASE WHEN LENGTH(ITEM_EAN_CODE) = 12 THEN ITEM_EAN_CODE END ITEM_UPC_CODE,
TO_NUMBER(REPLACE(ITEM_RETAIL_PRICE,'$','')) ITEM_RETAIL_PRICE,
TO_NUMBER(REPLACE(ITEM_COST_STD,'$','')) ITEM_COST_STD,
TO_DATE(LAUNCH_DATE,'FMMM/DD/YYYY') LAUNCH_DATE,
(SELECT PROD_CODE FROM ICTPROD1 WHERE II.PROD_CODE=PROD_DESC) PROD_CODE,
COLLECTION_CODE,
(SELECT ITEM_TYPE_CODE FROM ICTTYPE1 WHERE II.ITEM_TYPE_CODE=ITEM_TYPE_DESC) ITEM_TYPE_CODE,
DECODE(UPPER(ITEM_HIDE_FROM_CAT),'YES','1','NO','0') ITEM_HIDE_FROM_CAT,
(SELECT ITEM_CLASS_CODE FROM ICTCLAS1 WHERE II.ITEM_CLASS_CODE=ITEM_CLASS_DESC) ITEM_CLASS_CODE,
HC_CODE,
(SELECT COST_CATGY_CODE FROM ICTPROD1 WHERE II.PROD_CODE=PROD_DESC) COST_CATGY_CODE,
SEASON_CODE,
DECODE(UPPER(ITEM_BASIC_PROMO),'BASIC','B','PROMO','P') ITEM_BASIC_PROMO,
(SELECT COUNTRY_CODE FROM TATCNTRY WHERE II.COUNTRY_CODE=COUNTRY_NAME) COUNTRY_CODE,
DECODE(UPPER(ITEM_CONTAINS_ALCOHOL),'YES','1','NO','0') ITEM_CONTAINS_ALCOHOL,
DECODE(UPPER(ITEM_LOT_CTL),'YES','1','NO','0') ITEM_LOT_CTL,
DECODE(UPPER(ITEM_WEIGHT_CHECK),'YES','1','NO','0') ITEM_WEIGHT_CHECK,
DECODE(UPPER(ITEM_APPR_1ST_REC),'YES','1','NO','0') ITEM_APPR_1ST_REC,
DECODE(UPPER(ITEM_CRITICAL_TO_SHIP),'YES','1','NO','0') ITEM_CRITICAL_TO_SHIP,
TO_NUMBER(ITEM_SHELF_LIFE_YRS) ITEM_SHELF_LIFE_YRS,
(SELECT VEND_CODE FROM APTVEND1 WHERE II.VEND_CODE=VEND_NAME) VEND_CODE,
TO_NUMBER(ITEM_SO_QTY_MIN) ITEM_SO_QTY_MIN,
TO_NUMBER(ITEM_SO_QTY_MULT) ITEM_SO_QTY_MULT
FROM ICTITEMI II")

        Fill_Records("ICTITEMT")

        splImport.Visible = False

        dtXLS = dst.Tables(TABLE_NAME).Clone
        dtXLS.Columns.Add("ERROR MESSAGE")
        dtXLS.Columns("ERROR MESSAGE").SetOrdinal(0)
        dtXLS.DefaultView.Sort = "[ERROR MESSAGE]"
        dtXLS = dtXLS.DefaultView.ToTable

        For Each rowICTITEMT As DataRow In dst.Tables("ICTITEMT").Rows
            Dim row As DataRow = dtXLS.NewRow

            'Dim ITEM_CODE As String = rowICTITEMT("ITEM_CODE")

            'Absx1.txtFor("ITEM_CODE").Value = ITEM_CODE

            For Each col As DataColumn In dst.Tables("ICTITEM1").Columns
                If rowICTITEMT.Table.Columns.Contains(col.ColumnName) Then
                    If Not XLS_PROVIDED_COLUMNS.Contains(col.ColumnName) Then XLS_PROVIDED_COLUMNS.Add(col.ColumnName)
                    row(col.ColumnName) = rowICTITEMT(col.ColumnName)
                End If
            Next

            dtXLS.Rows.Add(row)
            ' dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
        Next

        ASCMAIN1.Progress("Validating import data...")

        PreProcess_XLS(dtXLS)


        splImport.Visible = True
        'Mode_Settings_XLS(False)
        ASCMAIN1.Progress("")

    End Sub


#End Region
End Class