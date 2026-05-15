Public Class SOFSHIPB

#Region "Declarations"

    Dim WHSE_CODE As String

    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name

    Dim BOL_NO As String
    Dim rowSOTSHIPB As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To

    Dim sqlSOTSHIPX As String
    Dim sqlSOTCART1 As String

    Dim SHIP_BOL_NOs As New List(Of String)
    Private finalizing As Boolean = False

    Private CreditCardProcessor As TAC.TAFCARDF
    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Private sqlSOTPICK1 As String = String.Empty
    Private isEdiCustomer As Boolean = False

    Private validDates() As Date = TAC.SOCMAIN1.Validate_Invoice_Date(Nothing, 0, 1, Nothing)

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTSHIPB.* from SOTSHIPB where BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIPS", "**", 0, False, "V", 1)
            .Tables("SOTSHIPS").Columns.Add("SEL")
            .Tables("SOTSHIPS").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "SOTSHIPB", "*", 1)

            ' No Common carriers on the form (No Fedex, UPS, USPS - carrier type = 'U')
            ' This was changed on 1/11/2013 - Debbie needs to see Amazon UPS orders.
            ' Do not show Web Orders SOTORDR0.ORDR_SOURCE <> 'W'
            ' 4/9/2013 - need to show web orders wher ethe state code is an armed forces state code. They ship Parcel Post
            sqlSOTSHIPX = "Select DISTINCT SOTSHIP1.*" & vbCrLf _
                & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ",ARTCUST1.CUST_NAME" & vbCrLf _
                & " from SOTSHIP1,SOTORDR0,ARTCUST1,SOTSVIA1,SOTCARR1, SOTORDR1" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)" & vbCrLf _
                & "   and SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE (+)" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO (+)" & vbCrLf _
                & "   and (SOTORDR0.ORDR_SOURCE <> 'W' OR (SOTORDR0.ORDR_SOURCE = 'W' AND NVL(SOTORDR1.STAX_CODE, '*') IN ('AE', 'AA', 'AP')))"

            ASCMAIN1.sql = sqlSOTSHIPX _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                & "   and SOTSHIP1.BILL_OF_LADING_NO IS NULL" _
                & "   and SOTSHIP1.WHSE_CODE = :PARM1" _
                & "   and SOTORDR0.CUST_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "VV", 1)
            .Tables("SOTSHIPX").Columns.Add("SEL")
            .Tables("SOTSHIPX").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = sqlSOTSHIPX _
                & "   and SOTSHIP1.BILL_OF_LADING_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "V", 1)
            .Tables("SOTSHIP1").Columns.Add("SEL")
            .Tables("SOTSHIP1").Columns("SEL").DefaultValue = "0"
            .Tables("SOTSHIP1").Columns.Add("SHIP_TOTAL_WGT_CALC", GetType(System.Decimal))
            .Tables("SOTSHIP1").Columns("SHIP_TOTAL_WGT_CALC").DefaultValue = 0
            .Tables("SOTSHIP1").Columns.Add("SHIP_CNT_CARTONS_CALC", GetType(System.Int16))
            .Tables("SOTSHIP1").Columns("SHIP_CNT_CARTONS_CALC").DefaultValue = 0


            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.SREP_CODE" & vbCrLf _
                & ", SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FREIGHT" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND" & vbCrLf _
                & ", SOTORDR5.CUST_NAME, SOTORDR5.CUST_ADDR1, SOTORDR5.CUST_ADDR2, SOTORDR5.CUST_ADDR3" & vbCrLf _
                & ", SOTORDR5.CUST_CITY, SOTORDR5.CUST_STATE, SOTORDR5.CUST_ZIP_CODE, SOTORDR5.CUST_COUNTRY" & vbCrLf _
                & ", SOTORDR5.CUST_CONTACT, SOTORDR5.CUST_PHONE, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
                & " from SOTPICK1, SOTORDR1, SOTORDR5" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR5.ORDR_NO (+) = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" & vbCrLf
            sqlSOTPICK1 = ASCMAIN1.sql

            ASCMAIN1.sql &= "and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTPICK1, SOTPICK2, SOTORDR2" & vbCrLf _
                & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & " and SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO and SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO" & vbCrLf _
                & " and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            'With .Tables.Add("SOTBOLGT")
            '    .Columns.Add("HANDLING_TYPE", GetType(System.String))
            '    .Columns.Add("HANDLING_UNITS", GetType(System.Int32))
            '    .Columns.Add("SHIP_CNT_CARTONS", GetType(System.Int32))
            '    .Columns.Add("SHIP_TOTAL_WGT", GetType(System.Decimal))
            '    .PrimaryKey = New DataColumn() {.Columns("HANDLING_TYPE")}
            'End With

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)
            Create_TDA(.Tables.Add, "SOTSHIPA", "*")
            'Create_TDA(.Tables.Add, "SOTCART1", "*")

            ASCMAIN1.sql = " SELECT SOTORDR1.CUST_STORE_NO, SOTCART1.* " _
             & " FROM SOTCART1, SOTPICK1, SOTORDR1" _
             & " WHERE SOTCART1.PICK_NO = SOTPICK1.PICK_NO" _
             & " AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
            sqlSOTCART1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTCART1", "**")

            ' Credit Card Processing
            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(.Tables.Add, "ARTCCPDA", "*")

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "TATEVNT1", "*")


            .Relations.Add("SOTSHIP1_SOTPICK1", dst.Tables("SOTSHIP1").Columns("SHIP_BOL_NO"), dst.Tables("SOTPICK1").Columns("SHIP_BOL_NO"))
            .Relations.Add("SOTPICK1_SOTCART1", dst.Tables("SOTPICK1").Columns("PICK_NO"), dst.Tables("SOTCART1").Columns("PICK_NO"))
        End With

        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdSOTCART1.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage

        Fill_Records("ICTWHSE1")

        grdSOTSHIPS.DataSource = dst.Tables("SOTSHIPS")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        'grdSOTBOLGT.DataSource = dst.Tables("SOTBOLGT")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")

        grdSOTSHIPS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTSHIPS.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        grdSOTSHIPS.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL", "BOL_NO", "WHSE_CODE", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"SEL"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With

        With grdSOTSHIP1.DisplayLayout
            .UseFixedHeaders = True
            With .Bands(0)
                For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "ORDR_CUST_PO", "ORDR_DEPT"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Header.Appearance.BackColor = Drawing.Color.White

                    If New String() {"SHIP_TOTAL_WGT", "SHIP_CNT_CARTONS", "HANDLING_TYPE", "HANDLING_UNITS"}.Contains(gcol.Key) Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"SHIP_TOTAL_WGT_CALC"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                Next
            End With
        End With

        With grdSOTCART1.DisplayLayout
            .UseFixedHeaders = True
            With .Bands(0)
                For Each COLUMN_NAME As String In New String() {"PICK_NO", "CUST_STORE_NO", "CART_NO", "CART_TOTAL_UNITS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Header.Appearance.BackColor = Drawing.Color.White

                    If New String() {"CART_TOTAL_WGT_ACTUAL", "PKG_CODE", "CART_MEMO"}.Contains(gcol.Key) Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If

                    If New String() {"CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"}.Contains(gcol.Key) Then
                        Create_Summary(grdSOTCART1, gcol.Key, "Sum")
                    End If
                Next
            End With
            Create_Summary(grdSOTCART1, "CART_NO", "Count")

        End With


        With grdSOTSHIPX.DisplayLayout
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            .UseFixedHeaders = True
            With .Bands(0)
                For Each COLUMN_NAME As String In New String() {"SEL", "SHIP_BOL_NO", "ORDR_CUST_PO", "ORDR_DEPT"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    If New String() {"SEL"}.Contains(gcol.Key) Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    End If
                Next
            End With
        End With


        Create_Summary(grdSOTSHIPS, "BOL_NO", "Count")
        Create_Summary(grdSOTSHIPS, "SEL")

        Create_Summary(grdSOTSHIP1, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIP1, New String() {"SHIP_TOTAL_WGT", "SHIP_CNT_CARTONS", "HANDLING_UNITS"})

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, "SEL")


        ASCMAIN1.Add_Value_List(grdSOTCART1, "PACKAGING_TYPE", Nothing, New String() {":", "P:Pallet", "B:Box"})

        'With dst.Tables("SOTBOLGT").Rows
        '    .Add(New Object() {"P"})
        '    .Add(New Object() {"C"})
        'End With
        'Sort_grdColumns(grdSOTBOLGT, "HANDLING_TYPE", True)

        Show_Filter(grdSOTSHIPS, True)
        grdSOTSHIPS.DisplayLayout.GroupByBox.Hidden = False

        '  splBOL.Panel2Collapsed = True

        Check_InquiryMode()

        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "SHIP_STATUS", Nothing, New String() {":", "P:In Pick", "F:Shipped", "C:Cancelled"})
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "HANDLING_TYPE", Nothing, New String() {":", "P:Pallet", "C:Carton", "S:Skid", "T:Totes", "L:Loose", "O:Other"})

        dteStart.MinDate = CDate("01/01/2000")
        dteStart.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)

        dteEnd.MinDate = dteStart.MinDate
        dteEnd.MaxDate = dteStart.MaxDate

        If InquiryMode Then
            MyBase.Absx1.dteFor("SHIPPED_ACTUAL").MinDate = CDate("01/01/2013")
            MyBase.Absx1.dteFor("SHIPPED_ACTUAL").MaxDate = validDates(1)
        Else
            MyBase.Absx1.dteFor("SHIPPED_ACTUAL").MinDate = validDates(0)
            MyBase.Absx1.dteFor("SHIPPED_ACTUAL").MaxDate = validDates(1)
        End If

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFBOLGI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                If chkMasterBOL.Checked Then
                    If dst.Tables("SOTSHIPS").Select("SEL = '1'").Length = 0 Then
                        MsgBox("To Start a New Master BOL, Select at least 1 BOL and then click New", _
                               MsgBoxStyle.OkOnly, "Cannot Start a New Master BOL without Selecting at least 1 Regular BOL")
                        Exit Sub
                    Else

                        If dst.Tables("SOTSHIPS").Select("SEL = '1' and BOL_STATUS <> 'O'").Length > 0 Then
                            EMsg &= vbCr & "You may only view BOLs with a status other than Open"
                            Exit Select
                        End If

                        WHSE_CODE = ""
                        CUST_CODE = ""
                        Dim SHIP_VIA_CODE As String = ""
                        Dim SHIP_ADDR_TYPE As String = ""
                        Dim SHIP_ADDR_CODE As String = ""

                        Dim BOL_DATE As Date
                        Dim SHIPPED_ACTUAL As Date
                        For Each row As DataRow In dst.Tables("SOTSHIPS").Select("SEL = '1'")
                            If row.Item("MASTER_BOL") & "" = "1" Then
                                EMsg &= vbCr & "Cannot select a Master BOL (" & row.Item("BOL_NO") & ") as part of a new Master BOL"
                            End If

                            If WHSE_CODE = "" Then
                                WHSE_CODE = row.Item("WHSE_CODE")
                                CUST_CODE = row.Item("CUST_CODE")
                                SHIP_VIA_CODE = row.Item("SHIP_VIA_CODE")
                                BOL_DATE = row.Item("BOL_DATE")
                                SHIPPED_ACTUAL = row.Item("SHIPPED_ACTUAL")
                                SHIP_ADDR_TYPE = row.Item("SHIP_ADDR_TYPE")
                                SHIP_ADDR_CODE = row.Item("SHIP_ADDR_CODE")
                            Else
                                If WHSE_CODE <> row.Item("WHSE_CODE") _
                                    Or CUST_CODE <> row.Item("CUST_CODE") _
                                    Or SHIP_VIA_CODE <> row.Item("SHIP_VIA_CODE") _
                                    Or BOL_DATE <> row.Item("BOL_DATE") _
                                    Or SHIPPED_ACTUAL <> row.Item("SHIPPED_ACTUAL") _
                                    Or SHIP_ADDR_TYPE <> row.Item("SHIP_ADDR_TYPE") Then
                                    EMsg &= "Warehouse, Customer, Ship Via, BOL Date, Ship-To, and Ship Date must match for all BOLs selected into a Master BOL"
                                End If
                            End If
                        Next
                    End If
                Else
                    If dst.Tables("SOTSHIPX").Select("SEL = '1'").Length = 0 Then
                        MsgBox("To Start a New BOL, Select at least 1 Shipment and then click New", _
                               MsgBoxStyle.OkOnly, "Cannot Start a New BOL without Selecting at least 1 Shipment")
                        Exit Sub
                    Else
                        WHSE_CODE = ""
                        CUST_CODE = ""

                        If dst.Tables("SOTSHIPS").Select("SEL = '1' and BOL_STATUS <> 'O'").Length > 0 Then
                            EMsg &= vbCr & "You may only view BOLs with a status other than Open"
                            Exit Select
                        End If

                        SHIP_BOL_NOs.Clear()

                        Dim SHIP_VIA_CODE As String = ""
                        Dim SHIP_ADDR_TYPE As String = ""
                        Dim SHIP_ADDR_CODE As String = ""

                        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("SEL = '1'")
                            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                            Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                            If rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                                EMsg &= "Shipment " & SHIP_BOL_NO & " is No Longer In Pick"
                            ElseIf rowSOTSHIP1.Item("BILL_OF_LADING_NO") & "" <> "" Then
                                EMsg &= "Shipment " & SHIP_BOL_NO & " is already part of BOL " & rowSOTSHIP1.Item("BILL_OF_LADING_NO")
                            End If

                            SHIP_BOL_NOs.Add(SHIP_BOL_NO)

                            If rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                                EMsg &= vbCr & "Cannot select a Shipment which is not in Pick"
                            ElseIf rowSOTSHIP1.Item("BILL_OF_LADING_NO") & "" <> "" Then
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is already part of a BOL"
                            End If

                            If rowSOTSHIP1.Item("SHIP_VIA_CODE") & "" = "" Then
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " has not been Routed (ie, assigned a Ship Via)"
                            End If

                            If WHSE_CODE = "" Then
                                WHSE_CODE = rowSOTSHIP1.Item("WHSE_CODE") & ""
                                CUST_CODE = row.Item("CUST_CODE") & ""
                                SHIP_VIA_CODE = rowSOTSHIP1.Item("SHIP_VIA_CODE") & ""
                                SHIP_ADDR_TYPE = rowSOTSHIP1.Item("SHIP_ADDR_TYPE") & ""
                                SHIP_ADDR_CODE = rowSOTSHIP1.Item("SHIP_ADDR_CODE") & ""
                            Else
                                If WHSE_CODE <> rowSOTSHIP1.Item("WHSE_CODE") & "" _
                                    Or CUST_CODE <> row.Item("CUST_CODE") & "" _
                                    Or SHIP_VIA_CODE <> rowSOTSHIP1.Item("SHIP_VIA_CODE") & "" _
                                    Or SHIP_ADDR_TYPE <> rowSOTSHIP1.Item("SHIP_ADDR_TYPE") & "" _
                                    Or SHIP_ADDR_CODE <> rowSOTSHIP1.Item("SHIP_ADDR_CODE") & "" Then
                                    EMsg &= "Warehouse, Customer, Ship-To, Ship Via must match for all Shipments selected into a BOL"
                                End If
                            End If
                        Next
                    End If
                End If

                'If Absx1.txtFor("WHSE_CODE").Text = "" Then
                '    EMsg &= vbCr & "You Must First Specify a Warehouse"
                'Else
                '    rowARTCUST1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                '    If rowARTCUST1 IsNot Nothing Then
                '        CUST_CODE = Absx1.txtFor("WHSE_CODE").Text
                '    Else
                '        EMsg &= vbCr & "No Record of Warehouse " & Absx1.txtFor("WHSE_CODE").Text
                '    End If
                'End If

                'If Absx1.txtFor("CUST_CODE").Text = "" Then
                '    EMsg &= vbCr & "You Must First Specify a Customer"
                'Else
                '    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                '    If rowARTCUST1 IsNot Nothing Then
                '        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                '    Else
                '        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                '    End If
                'End If

                'If EMsg = "" Then
                '    ' Customer must have a Sales Rep assigned
                'End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTSHIPB", CUST_CODE) Then Exit Sub
                Else
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Edit", "View"

                WHSE_CODE = ""
                CUST_CODE = ""
                BOL_NO = ""

                If Absx1.txtFor("BOL_NO").Text = "" Then
                    EMsg &= vbCr & "No BOL No Specified"
                Else
                    BOL_NO = Absx1.txtFor("BOL_NO").Text
                    rowSOTSHIPB = LookUp("SOTSHIPB", BOL_NO)
                    If rowSOTSHIPB Is Nothing Then
                        EMsg &= vbCr & "No Record of BOL No " & BOL_NO
                    Else
                        WHSE_CODE = rowSOTSHIPB.Item("WHSE_CODE")
                        CUST_CODE = rowSOTSHIPB.Item("CUST_CODE")
                        If rowSOTSHIPB.Item("BOL_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowSOTSHIPB.Item("BOL_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "BOL No " & BOL_NO & " has been Cancelled"
                                Case "F"
                                    EMsg &= vbCr & "BOL No " & BOL_NO & " has been Shipped"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "BOL No " & BOL_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And EntryMode = "E" Then
                    If Not ASCMAIN1.Logical_Lock("SOTSHIPB", BOL_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTSHIPB", CUST_CODE) Then Exit Sub
                End If

            Case "Update", "Finalize"

                If Absx1.dteFor("BOL_DATE").Value & "" = "" _
                    Or Absx1.dteFor("SHIPPED_ACTUAL").Value & "" = "" _
                    Or Absx1.dteFor("SCHED_DELIV_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "BOL Date, Date Shipped and Delivery Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("BOL_DATE").Value, "yyyyMMdd") _
                        > Format(Absx1.dteFor("SHIPPED_ACTUAL").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Ship Date cannot be Prior to BOL Date"
                    End If
                    If Format(Absx1.dteFor("SHIPPED_ACTUAL").Value, "yyyyMMdd") _
                        > Format(Absx1.dteFor("SCHED_DELIV_DATE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Delivery Date cannot be Prior to Ship Date"
                    End If
                End If

                If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                    EMsg &= vbCr & "Ship Via is required"
                Else
                    If LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Ship Via"
                    End If
                End If

                If EMsg.Length = 0 Then
                    If eItemKey = "Finalize" Then
                        ' No carriers of carrier_type 'U' can ship from here Unless Carrier Code Freight
                        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("SELECT SOTCARR1.*, SOTSVIA1.CARRIER_PROD_CODE " _
                                                      & " FROM SOTCARR1, SOTSVIA1" _
                                                      & " WHERE SOTCARR1.CARRIER_CODE =SOTSVIA1.CARRIER_CODE" _
                                                      & " AND SOTSVIA1.SHIP_VIA_CODE = :PARM1", "V", New Object() {Absx1.txtFor("SHIP_VIA_CODE").Text})
                        If rowSOTCARR1 Is Nothing OrElse rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty = "U" Then
                            If rowSOTCARR1 Is Nothing OrElse rowSOTCARR1.Item("CARRIER_PROD_CODE") & String.Empty <> "FRT" Then
                                EMsg &= "The selected Fedex or UPS ship via must be setup as a Freight shipment."
                            End If
                        End If

                        If Format(Absx1.dteFor("SHIPPED_ACTUAL").Value, "yyyyMMdd") < Format(DateTime.Now, "yyyyMMdd") Then
                            EMsg &= vbCr & "Ship Date which is used as the invoice date cannot be Prior to today."
                        ElseIf Format(Absx1.dteFor("SHIPPED_ACTUAL").Value, "yyyyMM") < Format(validDates(0), "yyyyMM") _
                            OrElse Format(Absx1.dteFor("SHIPPED_ACTUAL").Value, "yyyyMM") > Format(validDates(1), "yyyyMM") Then
                            EMsg &= vbCr & "Ship Date which is used as the invoice date must be between " & validDates(0) & " and " & validDates(1)
                        End If
                    End If
                End If

                If Absx1.txtFor("FRT_TERMS").Text = "" Then
                    EMsg &= vbCr & "Freight Terms is required"
                Else
                    If Not New String() {"PPD", "PPA", "COL"}.Contains(Absx1.txtFor("FRT_TERMS").Text) Then
                        EMsg &= vbCr & "Invalid Freight Terms"
                    End If
                End If

                If grdSOTSHIP1.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Shipments on BOL"
                Else
                    If Val(dst.Tables("SOTSHIP1").Compute("COUNT(SHIP_BOL_NO)", "ISNULL(SHIP_CNT_CARTONS,0) <= 0 or ISNULL(SHIP_TOTAL_WGT,0) <= 0") & "") <> 0 Then
                        EMsg &= vbCr & "There are Shipments on this BOL with invalid or 0 values for Cartons or Weight"
                    End If
                End If

                If isEdiCustomer Then
                    Dim MissingWeightError As Boolean = False
                    For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                        CalculateShipmentWeightAndCartons(rowSOTSHIP1.Item("SHIP_BOL_NO"))

                        Dim SHIP_CNT_CARTONS As Int16 = Val(rowSOTSHIP1.Item("SHIP_CNT_CARTONS") & String.Empty)
                        Dim SHIP_TOTAL_WGT As Decimal = Val(rowSOTSHIP1.Item("SHIP_TOTAL_WGT") & String.Empty)

                        If SHIP_CNT_CARTONS <= 0 OrElse SHIP_TOTAL_WGT <= 0 Then
                            EMsg &= "Shipment (" & rowSOTSHIP1.Item("SHIP_BOL_NO") & ") requires Number of Cartons and Total Weight be greater > 0"
                        End If

                        Dim SHIP_CNT_CARTONS_CALC As Int16 = Val(rowSOTSHIP1.Item("SHIP_CNT_CARTONS_CALC") & String.Empty)
                        Dim SHIP_TOTAL_WGT_CALC As Decimal = Val(rowSOTSHIP1.Item("SHIP_TOTAL_WGT_CALC") & String.Empty)

                        ' On 2/4/2013 as per Debbie do not enforce cartons - Take all day to process a shipment
                        If SHIP_CNT_CARTONS > Val(rowSOTSHIP1.Item("SHIP_CNT_CARTONS_CALC") & String.Empty) Then
                            'EMsg &= vbCr & "Total Cartons for shipment " & rowSOTSHIP1.Item("SHIP_BOL_NO") & " must be less equal total carton count"
                         End If

                        If SHIP_TOTAL_WGT < SHIP_TOTAL_WGT_CALC Then
                            'EMsg &= vbCr & "Total Weight for shipment " & rowSOTSHIP1.Item("SHIP_BOL_NO") & " must be greater equal total carton weight"
                        End If
                    Next

                    ' All cartons must have weights
                    Dim totalCartonsMissingWeight As Int16 = 0
                    For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                        If Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) <= 0 Then
                            totalCartonsMissingWeight += 1
                        End If
                    Next

                    ' On 2/4/2013 as per Debbie do not enforce cartons - Take all day to process a shipment
                    If totalCartonsMissingWeight > 0 Then
                        'EMsg &= vbCr & "There are " & totalCartonsMissingWeight & " cartons missing the weight."
                    End If

                    ' ASN customers require a Pro Number
                    Dim SHIP_REF As String = MyBase.Absx1.txtFor("SHIP_REF").Text.Trim
                    If EMsg.Length = 0 AndAlso eItemKey = "Finalize" AndAlso SHIP_REF.Length = 0 Then
                        If dst.Tables("SOTSHIP1").Select("SHIP_856_IND = '1'").Length > 0 Then
                            EMsg &= vbCr & "ASN customers require a Pro No."
                        End If
                    End If
                End If

                If EMsg.Length = 0 AndAlso eItemKey = "Finalize" Then
                    If MessageBox.Show("Are you sure you want to finalize this Bill of Lading?", "Finalize", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where BILL_OF_LADING_NO = '" & "" & "' and SHIP_STATUS <> 'P'"


                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "BOL has Shipments that have been Updated"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Delete this BOL?", _
                                    MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

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

            Case "Update", "Finalize"
                finalizing = eItemKey = "Finalize"
                If finalizing Then
                    If Not CaptureCreditCard() Then
                        Exit Sub
                    End If
                End If
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Refresh List"
                Setup_SOTSHIPX()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTSHIPB.Item("BOL_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        .Items("Finalize").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        .Items("Finalize").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Finalize").Settings.Enabled = DefaultableBoolean.False
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                .Items("Refresh List").Settings.Enabled = iScreenMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Finalize").Visible = (Not (EntryMode = "V") Or Not ScreenMode Or Not (EntryMode = "N"))
                .Items("Delete").Visible = (EntryMode = "E")
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Refresh List").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
            End With

            '.Groups("Totals").Visible = ScreenMode
        End With

        lblStatus.Visible = ScreenMode

        splBOLs.Visible = Not ScreenMode
        splBOL.Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(grpSHIPTO, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(grpTHIRD_PARTY, Not (EntryMode = "E" Or EntryMode = "N"))

        If rowSOTSHIPB IsNot Nothing AndAlso rowSOTSHIPB.Item("BOL_STATUS") = "F" Then
            tabSOTSHIPB.Tabs("Shipments").Visible = False
        Else
            tabSOTSHIPB.Tabs("Shipments").Visible = True
        End If

        If ScreenMode Then
            grdSOTSHIPX.Parent = tabSOTSHIPB.Tabs("Shipments").TabPage
            grdSOTSHIPX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            Show_Filter(grdSOTSHIPX, False)
            With grdSOTSHIPX.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = True
                .Columns("CUST_CODE").Hidden = True
                .Columns("CUST_NAME").Hidden = True
                .Columns("WHSE_CODE").Hidden = True
            End With
            grdSOTSHIPX.Text = "Other Shipments in Pick for Customer " & CUST_CODE & " in Warehouse " & WHSE_CODE
        Else
            grdSOTSHIPX.Parent = splBOLs.Panel2
            Show_Filter(grdSOTSHIPX, True)
            With grdSOTSHIPX.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = False
                .Columns("CUST_CODE").Hidden = False
                .Columns("CUST_NAME").Hidden = False
                .Columns("WHSE_CODE").Hidden = False
            End With
        End If

        If ScreenMode Then
            With grdSOTSHIP1.DisplayLayout.Override
                If EntryMode = "V" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                End If

            End With

            With grdSOTCART1.DisplayLayout.Override
                If EntryMode = "V" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False
                End If

            End With

            btnSaveShipTo.Visible = EntryMode <> "V"
            btnSaveFrt3PY.Visible = EntryMode <> "V"
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("BOL_NO").Text = ""
        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        WHSE_CODE = ""
        CUST_CODE = ""
        BOL_NO = ""
        finalizing = False
        isEdiCustomer = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPB", "SOTSHIP1", "SOTPICK1", "SOTPICK2", "ARTCUST1", "SOTSHIPS", "SOTSHIPX", "SOTCART1", _
             "ARTCCPA1", "ARTCCPA2", "ARTCCPDA", "SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTINVH1", "SOTINVH2", "ARTOPEN1", "TATEVNT1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")

        Load_SOTSHIPS()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Dim SHIP_BOL_NO_SEL As String = String.Empty

        If EntryMode = "N" Then
            BOL_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTSHIPB.BOL_NO"), "0" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

            rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
            rowSOTSHIPB = dst.Tables("SOTSHIPB").NewRow
            With rowSOTSHIPB
                .Item("BOL_NO") = BOL_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("MASTER_BOL") = "0"
                .Item("BOL_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("BOL_INST") = rowARTCUST1.Item("CUST_BOL_INST")
                .Item("THIRD_PARTY") = "0"

                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIPX").Select("SEL='1'")(0)
                SHIP_BOL_NO_SEL = rowSOTSHIP1.Item("SHIP_BOL_NO") & String.Empty

                .Item("FRT_TERMS") = rowSOTSHIP1.Item("FRT_TERMS")
                .Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF")
                .Item("SHIP_TRAILER_NO") = rowSOTSHIP1.Item("SHIP_TRAILER_NO")
                .Item("SHIP_SEAL_NO") = rowSOTSHIP1.Item("SHIP_SEAL_NO")
                dst.Tables("SOTSHIPB").Rows.Add(rowSOTSHIPB)

                .Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")
                .Item("SHIP_LOAD_NO") = rowSOTSHIP1.Item("SHIP_LOAD_NO") & String.Empty
                .Item("SHIP_APPT_NO") = rowSOTSHIP1.Item("SHIP_APPT_NO") & String.Empty
                .Item("SHIP_FOB") = rowSOTSHIP1.Item("SHIP_FOB") & String.Empty

            End With

            For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
                ASCMAIN1.sql = sqlSOTSHIPX & " AND SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                Fill_Records("SOTSHIP1", String.Empty, False, ASCMAIN1.sql)
            Next

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select()
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = BOL_NO
            Next
        Else
            rowSOTSHIPB = Fill_Record("SOTSHIPB", BOL_NO)
            'Fill_Records("SOTSHIP1", BOL_NO)

            If rowSOTSHIPB.Item("MASTER_BOL") & String.Empty = "1" Then
                chkMasterBOL.Checked = True
                ASCMAIN1.sql = "SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO IN (SELECT BOL_NO FROM SOTSHIPB WHERE MASTER_BOL_NO = '" & BOL_NO & "')"
                Fill_Records("SOTSHIP1", String.Empty, False, ASCMAIN1.sql)
            Else
                Fill_Records("SOTSHIP1", BOL_NO)
            End If

            If rowSOTSHIPB.Item("BOL_STATUS") = "F" Then
                tabSOTSHIPB.Tabs("Shipments").Visible = False
            End If

        End If

        CUST_CODE = rowSOTSHIPB.Item("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)

        isEdiCustomer = False
        If dst.Tables("SOTSHIP1").Select("SHIP_856_IND = '1' OR SHIP_810_IND = '1'").Length > 0 Then
            isEdiCustomer = True
        End If

        Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & ""
        CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
            CURR_EXCH_RATE = 1
        Else
            Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", rowARTCUST1.Item("CURR_CODE"))
            CURR_CODE = rowTATCURR1.Item("CURR_CODE")
            CURR_EXCH_RATE = rowTATCURR1.Item("CURR_EXCH_CUR")
        End If

        Sort_grdColumns(grdSOTSHIP1, "SHIP_BOL_NO")

        Setup_SOTSHIPX()
        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO")

        ' Load the SOTPICK1, SOTCART1, SOTORDR1, SOTORDR2 for the SOTSHIP1 records
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Rows
            ASCMAIN1.sql = sqlSOTPICK1 & " AND SOTPICK1.SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'"
            Fill_Records("SOTPICK1", String.Empty, False, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO = '" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'"
            Fill_Records("SOTORDR1", String.Empty, False, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR2 WHERE ORDR_NO IN (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "')"
            Fill_Records("SOTORDR2", String.Empty, False, ASCMAIN1.sql)
        Next

        If EntryMode = "N" AndAlso rowSOTSHIPB.Item("SHIP_FOB") & String.Empty = String.Empty Then
            rowSOTSHIPB.Item("SHIP_FOB") = dst.Tables("SOTORDR1").Compute("MAX(ORDR_FOB)", "") & String.Empty
        End If

        ASCMAIN1.sql = String.Empty
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Rows
            ASCMAIN1.sql &= ", '" & rowSOTPICK1.Item("PICK_NO") & "'"
        Next
        If ASCMAIN1.sql.Length > 0 Then
            ASCMAIN1.sql = ASCMAIN1.sql.Substring(2).Trim
            ASCMAIN1.sql = sqlSOTCART1 & " AND SOTCART1.PICK_NO IN (" & ASCMAIN1.sql & ")"
            Fill_Records("SOTCART1", String.Empty, True, ASCMAIN1.sql)
            ' default to our packaging
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                If rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = String.Empty Then
                    rowSOTCART1.Item("PACKAGING_TYPE") = 31
                End If
            Next
            Sort_grdColumns(grdSOTCART1, "PICK_NO,CUST_STORE_NO,CART_NO")
        End If

        lblINIT_DATE.Text = "Entered on " & Format(rowSOTSHIPB.Item("INIT_DATE"), "MM/dd/yyyy")

        If EntryMode = "N" Then
            lblStatus.Text = "New BOL"
            ' Pre Populate Ship to Fields with SOTORDR5 for the 1st Pick Ticket of 1st Shipment 

            Dim rowSOTPICK1 As DataRow = Nothing
            If dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO_SEL & "'").Length > 0 Then
                rowSOTPICK1 = dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO_SEL & "'")(0)
            End If

            If rowSOTPICK1 IsNot Nothing Then
                With rowSOTSHIPB
                    .Item("SHIP_TO_NAME") = rowSOTPICK1.Item("CUST_NAME") & String.Empty
                    .Item("SHIP_TO_ADDR1") = rowSOTPICK1.Item("CUST_ADDR1") & String.Empty
                    .Item("SHIP_TO_ADDR2") = rowSOTPICK1.Item("CUST_ADDR2") & String.Empty
                    .Item("SHIP_TO_ADDR3") = rowSOTPICK1.Item("CUST_ADDR3") & String.Empty
                    .Item("SHIP_TO_CITY") = rowSOTPICK1.Item("CUST_CITY") & String.Empty
                    .Item("SHIP_TO_STATE") = rowSOTPICK1.Item("CUST_STATE") & String.Empty
                    .Item("SHIP_TO_ZIP_CODE") = rowSOTPICK1.Item("CUST_ZIP_CODE") & String.Empty
                    .Item("SHIP_TO_COUNTRY") = rowSOTPICK1.Item("CUST_COUNTRY") & String.Empty
                    .Item("SHIP_TO_CONTACT") = rowSOTPICK1.Item("CUST_CONTACT") & String.Empty
                    .Item("SHIP_TO_PHONE") = rowSOTPICK1.Item("CUST_PHONE") & String.Empty
                End With
            End If

            Dim SHIP_CNT_CARTONS As Int16 = 0
            Dim CART_TOTAL_WGT_ACTUAL As Decimal = 0
            Dim CART_TOTAL_WGT_CALC As Decimal = 0
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Rows
                SHIP_BOL_NO_SEL = dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_LOAD_NO") & String.Empty
                SHIP_CNT_CARTONS = 0
                CART_TOTAL_WGT_ACTUAL = 0
                CART_TOTAL_WGT_CALC = 0
                For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO_SEL & "'", "")
                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & String.Empty
                    SHIP_CNT_CARTONS &= Val(dst.Tables("SOTCART1").Compute("COUNT(CART_NO)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                    CART_TOTAL_WGT_ACTUAL &= Val(dst.Tables("SOTCART1").Compute("SUM(ISNULL(CART_TOTAL_WGT_ACTUAL, 0))", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                    CART_TOTAL_WGT_CALC &= Val(dst.Tables("SOTCART1").Compute("SUM(ISNULL(CART_TOTAL_WGT_CALC, 0))", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                Next

                rowSOTSHIP1.Item("SHIP_TOTAL_WGT") = IIf(CART_TOTAL_WGT_ACTUAL > 0, CART_TOTAL_WGT_ACTUAL, CART_TOTAL_WGT_CALC)
                rowSOTSHIP1.Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
            Next
        Else
            Select Case rowSOTSHIPB.Item("BOL_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "F"
                    lblStatus.Text = "Shipped"
            End Select
        End If

        With grdSOTSHIP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"SHIP_TOTAL_WGT", "SHIP_CNT_CARTONS", "HANDLING_TYPE", "HANDLING_UNITS"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdSOTSHIP1.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
        Else
            With grdSOTSHIP1.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        End If

        grpTHIRD_PARTY.Visible = rowSOTSHIPB.Item("THIRD_PARTY") & "" = "1"

        'Display_Totals()
        EnforceConstraints(True)

        If EntryMode = "N" Then
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                CalculateShipmentWeightAndCartons(rowSOTSHIP1.Item("SHIP_BOL_NO"))
            Next
        End If

        If grdSOTSHIP1.Rows.Count > 0 Then
            grdSOTSHIP1.Rows(0).Activate()
            grdSOTSHIP1_AfterRowActivate(Nothing, Nothing)
        End If

        ' sometimes the shia via desc does not show
        Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
        MyBase.Absx1.txtFor("SHIP_VIA_CODE").Clear()
        MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text = SHIP_VIA_CODE

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            CalculateShipmentWeightAndCartons(rowSOTSHIP1.Item("SHIP_BOL_NO"))
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, BOL_NO)
        For Each TABLE_NAME As String In New String() {"SOTSHIPB"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where BOL_NO = '" & BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        Dim inTransaction As Boolean = False

        Try
            '*******************************************************************************
            ' If finalizing then update Status Flag and SOTSHIP1 records
            ' Will not have credit card processing in this screen 01/12/2102
            '*******************************************************************************

            If finalizing Then
                'Dim rowSOTSHIPB As DataRow = dst.Tables("SOTSHIPB").Rows(0)
                rowSOTSHIPB.Item("BOL_STATUS") = "F"
                rowSOTSHIPB.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPB.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIPB.Item("SHIPPED_ACTUAL") = MyBase.Absx1.dteFor("SHIPPED_ACTUAL").Value

                ' If there is additional freight then add it to one of the Pick Tickets
                If Val(rowSOTSHIPB.Item("SHIP_FREIGHT") & String.Empty) > 0 Then
                    Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("", "SHIP_BOL_NO, ORDR_NO")(0)
                    rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowSOTSHIPB.Item("SHIP_FREIGHT") & String.Empty)
                End If

                dst.Tables("SOTSHIP1").AcceptChanges()
                For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                    Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                    With rowSOTSHIP1
                        .Item("FRT_TERMS") = rowSOTSHIPB.Item("FRT_TERMS")
                        .Item("SHIP_REF") = rowSOTSHIPB.Item("SHIP_REF")
                        .Item("SHIP_TRAILER_NO") = rowSOTSHIPB.Item("SHIP_TRAILER_NO")
                        .Item("SHIP_SEAL_NO") = rowSOTSHIPB.Item("SHIP_SEAL_NO")
                        .Item("SHIP_VIA_CODE") = rowSOTSHIPB.Item("SHIP_VIA_CODE")
                        .Item("SHIP_LOAD_NO") = rowSOTSHIPB.Item("SHIP_LOAD_NO") & String.Empty
                        .Item("SHIP_APPT_NO") = rowSOTSHIPB.Item("SHIP_APPT_NO") & String.Empty
                        .Item("BILL_OF_LADING_NO") = rowSOTSHIPB.Item("BOL_NO") & String.Empty
                        .Item("WHSE_CODE") = rowSOTSHIPB.Item("WHSE_CODE") & String.Empty
                        .Item("THIRD_PARTY") = rowSOTSHIPB.Item("THIRD_PARTY") & String.Empty
                        .Item("SHIP_DATE_SHIPPED") = rowSOTSHIPB.Item("SHIPPED_ACTUAL") & String.Empty
                        .Item("INV_DATE") = rowSOTSHIPB.Item("SHIPPED_ACTUAL") & String.Empty
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                        .Item("SHIP_FOB") = rowSOTSHIPB.Item("SHIP_FOB")
                        .Item("SHIP_STATUS") = "F"
                    End With

                    ' These table are used only from finalization.
                    Fill_Records("SOTPICK2", New Object() {SHIP_BOL_NO}, False)
                    ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                    Fill_Records("SOTORDR5", String.Empty, False, ASCMAIN1.sql)
                Next

                UpdateOrderPicksAndCreateInvoices()

            End If
            '*******************************************************************************
            ' End If finalizing then update Status Flag and SOTSHIP1 records
            '*******************************************************************************

            BeginTrans()
            inTransaction = True

            If EntryMode <> "N" Then Delete_Records()

            INIT_LAST("SOTSHIPB", False, , True)
            Dim sqldelete As String = "BOL_NO = '" & BOL_NO & "'"
            Update_Record_TDA("SOTSHIPB", sqldelete)

            INIT_LAST("SOTSHIP1", False, , True)
            Update_Record_TDA("SOTSHIP1")
            Update_Record_TDA("SOTCART1")

            ' Need to make sure PT's have weights
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select()
                Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
            Next
            Update_Record_TDA("SOTPICK1")

            If finalizing Then
                Update_Record_TDA("SOTORDR1")
                Update_Record_TDA("SOTORDR2")
                Update_Record_TDA("SOTINVH1")
                Update_Record_TDA("SOTINVH2")
                Update_Record_TDA("SOTPICK1")
                Update_Record_TDA("ARTOPEN1")
                Update_Record_TDA("TATEVNT1")

                For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows

                    ASCMAIN1.Progress("", rowSOTINVH1.Item("INV_NO"))

                    ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                    ASCDATA1.ExecuteSQL()

                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
                        TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))
                    End If

                    ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
                           New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")}, _
                           New String() {"INV_TYPE_IN", "INV_NO_IN"})

                Next

                ' Process each BOL, now that it is in Oracle
                For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                    Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                    ASCMAIN1.sql = "Update SOTORDR1" _
                           & " Set ORDR_STATUS = 'F'" & vbCrLf _
                           & " where ORDR_NO in " & vbCrLf _
                           & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                           & " where SHIP_BOL_NO = :PARM1)" & vbCrLf
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {SHIP_BOL_NO})

                    ASCMAIN1.sql = "Update SOTPICK1" _
                            & " Set PICK_STATUS = 'F'" & vbCrLf _
                            & " where PICK_NO in " & vbCrLf _
                            & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
                            & " where SHIP_BOL_NO = :PARM1)" & vbCrLf
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {SHIP_BOL_NO})

                    ASCMAIN1.sql = "Update SOTORDR1 " _
                             & "Set TERM_CODE = :PARM1, ORDR_DEPT = :PARM2" & vbCrLf _
                             & " where ORDR_NO in " & vbCrLf _
                             & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                             & " where SHIP_BOL_NO = :PARM3)" & vbCrLf
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", _
                                        New Object() {rowSOTSHIP1.Item("TERM_CODE"), _
                                                      rowSOTSHIP1.Item("ORDR_DEPT"), _
                                                      rowSOTSHIP1.Item("SHIP_BOL_NO")})

                    ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'); END;"
                    ASCDATA1.ExecuteSQL()
                Next

            End If


            CommitTrans("Update Complete")

        Catch ex As Exception
            If inTransaction Then Rollback(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "BOL_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("WHSE_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code or a Warehouse", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTSHIPB.BOL_STATUS = 'O' "
                End If

                If Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    sql_where &= " and SOTSHIPB.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
                End If
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTSHIPB.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If

            Case "SHIP_TO_CODE"
                sql_where &= " and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "' and CUST_ADDR_TYPE = 'ST'"

            Case "FRT_3PY_CODE"
                sql_where &= " and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "' and CUST_ADDR_TYPE = '3P'"

        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("BOL_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTSHIPB"
            E.COLUMN_NAME = "BOL_NO"
            E.CODE_VALUE = Absx1.txtFor("BOL_NO").Text
            E.DESC_VALUE = "Reservation"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTSHIPB.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPS, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSHIPX, "B", "Add Shipment")
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
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTSHIPX"
                    tlb_btn = DirectCast(tlb.Tools("Add Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = (EntryMode = "N" OrElse EntryMode = "E")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Add Shipment"
                If Not (EntryMode = "N" OrElse EntryMode = "E") Then
                    MessageBox.Show("Invalid form State.", "Add Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim WHSE_CODE As String = MyBase.Absx1.txtFor("WHSE_CODE").Text
                Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value

                Dim sql As String = sqlSOTSHIPX
                Dim rowSOTSHIPX As DataRow = ASCDATA1.GetDataRow(sqlSOTSHIPX & " AND SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")

                If rowSOTSHIPX Is Nothing Then
                    MessageBox.Show("The selected Shipment cannot be located.", "Add Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim emsg As String = String.Empty
                If rowSOTSHIPX.Item("SHIP_STATUS") & String.Empty <> "P" Then
                    emsg = "Invalid Ship Status"
                ElseIf rowSOTSHIPX.Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
                    emsg = "Shipment already assigned to a BOL"
                ElseIf rowSOTSHIPX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                    emsg = "Invalid Ship Via"
                ElseIf rowSOTSHIPX.Item("SHIP_VIA_CODE") & String.Empty <> SHIP_VIA_CODE Then
                    emsg = "Invalid Ship Via"
                ElseIf rowSOTSHIPX.Item("WHSE_CODE") & String.Empty <> WHSE_CODE Then
                    emsg = "Invalid Warehouse"
                ElseIf rowSOTSHIPX.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                    emsg = "Invalid Customer"
                End If

                If emsg.Length = 0 AndAlso dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows).Length > 0 Then
                    Dim SHIP_ADDR_TYPE As String = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)(0).Item("SHIP_ADDR_TYPE") & String.Empty
                    Dim SHIP_ADDR_CODE As String = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)(0).Item("SHIP_ADDR_CODE") & String.Empty

                    If rowSOTSHIPX.Item("SHIP_ADDR_TYPE") & String.Empty <> SHIP_ADDR_TYPE Then
                        emsg = "Invalid Ship Address Type"
                    ElseIf rowSOTSHIPX.Item("SHIP_ADDR_CODE") & String.Empty <> SHIP_ADDR_CODE Then
                        emsg = "Invalid Ship Adddress Code"
                    End If
                End If

                If emsg.Length > 0 Then
                    MessageBox.Show("The selected Shipment no longer meets the current BOL criteria: " & emsg, "Add Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                ' Lock the shipment
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO, False, True, False, 5) Then
                    ASCMAIN1.MultiTask_Release(, , 5)
                    Exit Sub
                End If

                sql = sqlSOTSHIPX & " and SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                Fill_Records("SOTSHIP1", String.Empty, False, sql)
                'Fill_Records("SOTPICK1", New Object() {SHIP_BOL_NO}, False)
                ASCMAIN1.sql = sqlSOTPICK1 & " AND SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                Fill_Records("SOTPICK1", String.Empty, False, ASCMAIN1.sql)

                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "")
                    sql = sqlSOTCART1 & "  AND SOTCART1.PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'"
                    Fill_Records("SOTCART1", String.Empty, False, sql)
                    ' default to our packaging
                    For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                        If rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = String.Empty Then
                            rowSOTCART1.Item("PACKAGING_TYPE") = 31
                        End If
                    Next
                    Sort_grdColumns(grdSOTCART1, "PICK_NO,CUST_STORE_NO,CART_NO")
                Next

                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = rowSOTSHIPB.Item("BOL_NO") & String.Empty
                ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO = '" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'"
                Fill_Records("SOTORDR1", String.Empty, False, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select * from SOTORDR2 WHERE ORDR_NO IN (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "')"
                Fill_Records("SOTORDR2", String.Empty, False, ASCMAIN1.sql)

                Setup_SOTSHIPX()
                CalculateShipmentWeightAndCartons(SHIP_BOL_NO)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPS()
                End If

            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPS()
                End If

            Case "BOL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPS()

                    'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    'If WHSE_CODE <> "" Then
                    '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    '    If rowICTWHSE1 IsNot Nothing Then

                    '    End If
                    'End If
                End If

            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPS()

                    'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    'If CUST_CODE <> "" Then
                    '    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    '    If rowARTCUST1 IsNot Nothing Then

                    '    End If
                    'End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Load_SOTSHIPS()
            Case "CUST_CODE"
                Load_SOTSHIPS()
            Case "BOL_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub chk_CheckedChanged(sender As Object, e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "THIRD_PARTY"
                grpTHIRD_PARTY.Visible = (Absx1.chkFor("THIRD_PARTY").Checked)
        End Select
    End Sub
    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTSHIPS()

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        Dim sqlw As String = ""
        Dim caption As String = ""

        If WHSE_CODE <> "" Then
            sqlw &= " and WHSE_CODE = '" & WHSE_CODE & "'"
            caption = " and Warehouse " & WHSE_CODE
        End If
        If CUST_CODE <> "" Then
            sqlw &= " and CUST_CODE = '" & CUST_CODE & "'"
            caption = " and Customer " & CUST_CODE
        End If

        If caption <> "" Then caption = " associated with " & Mid(caption, 6)

        If optStatus.Value = "O" Then
            ASCMAIN1.sql = "Select * from SOTSHIPB where BOL_STATUS = 'O'" & sqlw
            grdSOTSHIPS.Text = "Open BOLs" & caption
            splBOLs.Panel2Collapsed = False
        Else
            If Not IsDate(dteStart.Value) OrElse Not IsDate(dteEnd.Value) Then
                MessageBox.Show("Start and End date must be a valid date.", "Shipment Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ASCMAIN1.sql = "Select * from SOTSHIPB where BOL_STATUS = 'F'" & sqlw _
             & " and BOL_DATE between '" & CDate(dteStart.Value).ToString("dd-MMM-yyyy") & "' and '" & CDate(dteEnd.Value).ToString("dd-MMM-yyyy") & "'"
            grdSOTSHIPS.Text = "Finalized BOLs" & caption & " between " & dteStart.Value & " and " & dteStart.Value
            splBOLs.Panel2Collapsed = True
        End If

        Fill_Records("SOTSHIPS", "", , ASCMAIN1.sql)

        Sort_grdColumns(grdSOTSHIPS, "BOL_NO".ToLower)

        grdSOTSHIPS.Visible = True


        ASCMAIN1.sql = sqlSOTSHIPX & Replace(Replace(sqlw, "CUST_CODE", "SOTORDR0.CUST_CODE"), "WHSE_CODE", "SOTSHIP1.WHSE_CODE")
        ASCMAIN1.sql &= "and SOTSHIP1.BILL_OF_LADING_NO IS NULL and SOTSHIP1.SHIP_STATUS = 'P'"

        Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
        grdSOTSHIPX.Text = "Shipments in Pick" & caption
        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO")

        grdSOTSHIPS.Visible = True
    End Sub

    Sub Print_Record()

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "SORSHIPB"
        Generate_Report(RPT, "Bill of Lading")

        ' Need to make sure PT's have weights
        'SOTCART1.CART_TOTAL_WGT_ACTUAL, SOTPICK1.PICK_TOTAL_WGT
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(PICK_TOTAL_WGT,0) = 0")
            Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
        Next

        CR_params.Add("SUBT", "")
        RPT = "SORSHIPM"
        Generate_Report(RPT, "Consolidation Manifest")

        Print_Report_End()
    End Sub

    Private Sub grdSOTSHIPS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPS.DoubleClickRow
        Absx1.txtFor("BOL_NO").Text = e.Row.Cells("BOL_NO").Value
        Click_Command("View")
    End Sub

    Sub Dependent_Updates(S As Integer, BOL_NO As String)
        If S = -1 Then
            ASCMAIN1.sql = "Update SOTSHIP1 Set " _
                & "  BILL_OF_LADING_NO = Null" _
                & ", SHIP_TOTAL_WGT = Null" _
                & ", SHIP_CNT_CARTONS = Null" _
                & ", HANDLING_TYPE = Null" _
                & ", HANDLING_UNITS = Null" _
                & " where BILL_OF_LADING_NO = '" & BOL_NO & "'"
            ASCDATA1.ExecuteSQL()
        Else
            Update_Record_TDA("SOTSHIP1")
        End If
    End Sub

    Sub Display_Totals()
        'Dim KEY As Int32 = 0
        'For Each HANDLING_TYPE As String In New String() {"P", "C"}
        '    Dim rowSOTBOLGT As DataRow = dst.Tables("SOTBOLGT").Rows.Find(HANDLING_TYPE)
        '    rowSOTBOLGT.Item("HANDLING_UNITS") = Val(dst.Tables("SOTSHIP1").Compute("SUM(HANDLING_UNITS)", "") & "")
        '    rowSOTBOLGT.Item("SHIP_TOTAL_WGT") = Val(dst.Tables("SOTSHIP1").Compute("SUM(SHIP_TOTAL_WGT)", "") & "")
        '    rowSOTBOLGT.Item("SHIP_CNT_CARTONS") = Val(dst.Tables("SOTSHIP1").Compute("SUM(SHIP_CNT_CARTONS)", "") & "")
        'Next
    End Sub

#Region "grdSOTSHIP1"

    Private Sub grdSOTSHIP1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSHIP1.AfterCellUpdate
        With grdSOTSHIP1.ActiveRow
            Select Case e.Cell.Column.Key
            End Select
        End With
    End Sub


    Private Sub grdSOTSHIP1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIP1.AfterRowActivate

        Dim SHIP_BOL_NO As String = grdSOTSHIP1.ActiveRow.Cells("SHIP_BOL_NO").Text

        Dim viewSOTPICK1 As DataView = New DataView(dst.Tables("SOTPICK1"))
        viewSOTPICK1.RowFilter = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        grdSOTPICK1.DataSource = viewSOTPICK1
        grdSOTPICK1.Text = "Pick Tickets for Shipment " & SHIP_BOL_NO

        Dim PICK_NOS As String = String.Empty
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            PICK_NOS &= " OR PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'"
        Next

        If PICK_NOS.Length > 0 Then
            PICK_NOS = PICK_NOS.Substring(3).Trim
        End If

        Dim viewSOTCART1 As DataView = New DataView(dst.Tables("SOTCART1"))
        viewSOTCART1.RowFilter = PICK_NOS
        grdSOTCART1.DataSource = viewSOTCART1
        grdSOTCART1.Text = "Cartons for Shipment " & SHIP_BOL_NO

    End Sub

#End Region

    Private Sub Setup_SOTSHIPX()


        Dim WHSE_CODE As String = MyBase.Absx1.txtFor("WHSE_CODE").Text
        Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
        Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text

        Dim sql As String = sqlSOTSHIPX

        sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
        sql &= " and SOTSHIP1.BILL_OF_LADING_NO IS NULL"
        sql &= " and SOTSHIP1.SHIP_VIA_CODE IS NOT NULL"

        sql &= " and SOTSHIP1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'"
        sql &= " and SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"
        sql &= " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"

        If dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
            Fill_Records("SOTSHIPX", String.Empty, True, sql)
            Exit Sub
        End If

        Dim SHIP_ADDR_TYPE As String = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)(0).Item("SHIP_ADDR_TYPE") & String.Empty
        Dim SHIP_ADDR_CODE As String = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)(0).Item("SHIP_ADDR_CODE") & String.Empty

        sql &= " and SOTSHIP1.SHIP_ADDR_TYPE = '" & SHIP_ADDR_TYPE & "'"

        If SHIP_ADDR_CODE.Length = 0 Then
            sql &= " and SOTSHIP1.SHIP_ADDR_CODE IS NULL"
        Else
            sql &= " and SOTSHIP1.SHIP_ADDR_CODE = '" & SHIP_ADDR_CODE & "'"
        End If

        Dim SHIP_BOL_NOs As String = String.Empty
        For Each row As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
            SHIP_BOL_NOs &= ", '" & row.Item("SHIP_BOL_NO") & "'"
        Next

        If SHIP_BOL_NOs.Length > 0 Then
            SHIP_BOL_NOs = SHIP_BOL_NOs.Substring(2).Trim
            sql &= " and SOTSHIP1.SHIP_BOL_NO NOT IN (" & SHIP_BOL_NOs & ")"
        End If

        Fill_Records("SOTSHIPX", String.Empty, True, sql)

    End Sub

    Private Sub btnSaveShipTo_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveShipTo.Click
        SaveShipAddress("SHIP_TO")
    End Sub

    Private Sub btnSaveFrt3PY_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveFrt3PY.Click
        SaveShipAddress("FRT_3PY")
    End Sub

    Private Sub SaveShipAddress(ByVal prefix As String)

        'SHIP_TO_CODE, FRT_3PY_CODE
        Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text
        Dim CUST_ADDR_CODE As String = MyBase.Absx1.txtFor(prefix & "_CODE").Text
        CUST_ADDR_CODE = CUST_ADDR_CODE.ToUpper.Trim

        If CUST_ADDR_CODE.Length = 0 Then
            MessageBox.Show("You must provide an Address Code.", "Add Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        MyBase.Absx1.txtFor(prefix & "_CODE").Text = CUST_ADDR_CODE

        Dim CUST_ADDR_TYPE As String = String.Empty
        Select Case prefix
            Case "SHIP_TO"
                CUST_ADDR_TYPE = "ST"
            Case "FRT_3PY"
                CUST_ADDR_TYPE = "3P"
            Case Else
                MessageBox.Show("Unable to determine the Address Type.", "Add Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
        End Select

        ' Clean up data Look for minimum requirements
        For Each suffix As String In New String() {"NAME", "ADDR1", "ADDR2", "CITY", "STATE", "ZIP_CODE", "COUNTRY", "CONTACT"}
            MyBase.Absx1.txtFor(prefix & "_" & suffix).Text = MyBase.Absx1.txtFor(prefix & "_" & suffix).Text.Trim
        Next

        MyBase.Absx1.txtFor(prefix & "_" & "NAME").Text = StrConv(MyBase.Absx1.txtFor(prefix & "_" & "NAME").Text, VbStrConv.ProperCase)
        MyBase.Absx1.txtFor(prefix & "_" & "ADDR1").Text = StrConv(MyBase.Absx1.txtFor(prefix & "_" & "ADDR1").Text, VbStrConv.ProperCase)
        MyBase.Absx1.txtFor(prefix & "_" & "ADDR2").Text = StrConv(MyBase.Absx1.txtFor(prefix & "_" & "ADDR2").Text, VbStrConv.ProperCase)
        MyBase.Absx1.txtFor(prefix & "_" & "CITY").Text = StrConv(MyBase.Absx1.txtFor(prefix & "_" & "CITY").Text, VbStrConv.ProperCase)
        MyBase.Absx1.txtFor(prefix & "_" & "STATE").Text = MyBase.Absx1.txtFor(prefix & "_" & "STATE").Text.ToUpper
        MyBase.Absx1.txtFor(prefix & "_" & "COUNTRY").Text = MyBase.Absx1.txtFor(prefix & "_" & "COUNTRY").Text.ToUpper
        MyBase.Absx1.txtFor(prefix & "_" & "CONTACT").Text = StrConv(MyBase.Absx1.txtFor(prefix & "_" & "CONTACT").Text, VbStrConv.ProperCase)

        If MyBase.Absx1.txtFor(prefix & "_" & "NAME").TextLength = 0 _
            OrElse MyBase.Absx1.txtFor(prefix & "_" & "ADDR1").TextLength = 0 _
            OrElse MyBase.Absx1.txtFor(prefix & "_" & "CITY").TextLength = 0 _
            OrElse MyBase.Absx1.txtFor(prefix & "_" & "STATE").TextLength = 0 _
            OrElse MyBase.Absx1.txtFor(prefix & "_" & "ZIP_CODE").TextLength = 0 Then
            MessageBox.Show("Name, Address Line 1, City, State and Zip Code are required.", "Add Address", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Fill_Records("SOTSHIPA", New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
        If dst.Tables("SOTSHIPA").Rows.Count > 0 Then
            If MessageBox.Show("Do you want to Update the existing address data with the data on the screen?", "Add Address", _
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        Dim rowSOTSHIPA As DataRow = Nothing
        If dst.Tables("SOTSHIPA").Rows.Count = 0 Then
            rowSOTSHIPA = dst.Tables("SOTSHIPA").NewRow
            rowSOTSHIPA.Item("CUST_CODE") = CUST_CODE
            rowSOTSHIPA.Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            rowSOTSHIPA.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            dst.Tables("SOTSHIPA").Rows.Add(rowSOTSHIPA)
        Else
            rowSOTSHIPA = dst.Tables("SOTSHIPA").Rows(0)
        End If

        rowSOTSHIPA.Item("CUST_NAME") = MyBase.Absx1.txtFor(prefix & "_" & "NAME").Text
        rowSOTSHIPA.Item("CUST_ADDR1") = MyBase.Absx1.txtFor(prefix & "_" & "ADDR1").Text
        rowSOTSHIPA.Item("CUST_ADDR2") = MyBase.Absx1.txtFor(prefix & "_" & "ADDR2").Text
        rowSOTSHIPA.Item("CUST_CITY") = MyBase.Absx1.txtFor(prefix & "_" & "CITY").Text
        rowSOTSHIPA.Item("CUST_STATE") = MyBase.Absx1.txtFor(prefix & "_" & "STATE").Text
        rowSOTSHIPA.Item("CUST_ZIP_CODE") = MyBase.Absx1.txtFor(prefix & "_" & "ZIP_CODE").Text
        rowSOTSHIPA.Item("CUST_COUNTRY") = MyBase.Absx1.txtFor(prefix & "_" & "COUNTRY").Text
        rowSOTSHIPA.Item("CUST_CONTACT") = MyBase.Absx1.txtFor(prefix & "_" & "CONTACT").Text
        rowSOTSHIPA.Item("CUST_PHONE") = MyBase.Absx1.medFor(prefix & "_" & "PHONE").Text

        Try
            BeginTrans()
            Update_Record_TDA("SOTSHIPA")
            CommitTrans("Address Updated")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        grpDateRange.Enabled = optStatus.Value = "F"
    End Sub

    Private Sub cmdRefresh_Click(sender As System.Object, e As System.EventArgs) Handles cmdRefresh.Click
        Load_SOTSHIPS()
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate
        Dim SHIP_BOL_NO As String = grdSOTSHIP1.ActiveRow.Cells("SHIP_BOL_NO").Text
        CalculateShipmentWeightAndCartons(SHIP_BOL_NO)
    End Sub

    Private Sub CalculateShipmentWeightAndCartons(ByVal SHIP_BOL_NO As String)
        Dim SHIP_TOTAL_WGT_CALC As Decimal = 0
        Dim SHIP_CNT_CARTONS_CALC As Int16 = 0

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        If rowSOTSHIP1 Is Nothing Then
            Exit Sub
        End If

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "", DataViewRowState.CurrentRows)
            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "", DataViewRowState.CurrentRows)
                SHIP_TOTAL_WGT_CALC += Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty)
                SHIP_CNT_CARTONS_CALC += 1
            Next
        Next

        rowSOTSHIP1.Item("SHIP_TOTAL_WGT_CALC") = SHIP_TOTAL_WGT_CALC
        rowSOTSHIP1.Item("SHIP_CNT_CARTONS_CALC") = SHIP_CNT_CARTONS_CALC

    End Sub

    Private Sub grdSOTSHIP1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSHIP1.AfterRowUpdate
        Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Text
        CalculateShipmentWeightAndCartons(SHIP_BOL_NO)
    End Sub


    Private Function CaptureCreditCard() As Boolean

        Dim CreditCardProcessed As Boolean = True
        Dim shippingFreight As Decimal = 0

        Dim xtraFreight As Double = Val(rowSOTSHIPB.Item("SHIP_FREIGHT") & String.Empty)

        Try
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(CCPA_NO_ORDR, '') <> ''", "SHIP_BOL_NO")

                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim CCPA_NO_ORDR As String = rowSOTPICK1.Item("CCPA_NO_ORDR")
                Dim chargeAmount As Decimal = 0
                Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty

                ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
                Fill_Records("SOTORDR5", String.Empty, False, ASCMAIN1.sql)


                CCPA_NO_ORDR = CCPA_NO_ORDR.Trim
                If CCPA_NO_ORDR.Length = 0 Then Continue For

                If rowSOTPICK1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                    ' Credit card against pick ticket already processed. May have been an error in the code
                    Continue For
                End If

                chargeAmount = Val(ASCDATA1.GetDataValue("SELECT SUM(NVL(PICK_QTY_CONF, 0) * NVL(PICK_UNIT_PRICE, 0)) FROM SOTPICK2 WHERE PICK_NO = :PARM1", "V", New Object() {PICK_NO}) & String.Empty)

                shippingFreight = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty)
                chargeAmount += shippingFreight + xtraFreight
                ' Apply Only Once
                xtraFreight = 0

                If chargeAmount > 0 Then
                    Try
                        ' This is done to preserve credit card transactions if the code causes an error after this point
                        Dim CCPA_NO As String = ProcessCreditCardAuthorization(CCPA_NO_ORDR, chargeAmount, shippingFreight, 0)
                        CreditCardProcessed = CCPA_NO.Length > 0 AndAlso CreditCardProcessed

                        rowSOTPICK1.Item("CCPA_NO") = CCPA_NO

                        ' Flag Pick Ticket as getting charged. This prevents a duplicate charge of an error occurs
                        If CCPA_NO.Length > 0 Then
                            MyBase.BeginTrans()
                            ASCDATA1.ExecuteSQL("Update SOTPICK1 set CCPA_NO = '" & CCPA_NO & "' where PICK_NO = '" & PICK_NO & "'")
                            MyBase.CommitTrans()
                        End If

                    Catch ex As Exception
                        MyBase.Rollback(ex.Message)
                    End Try
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("The following error occurred when processing the credit card: " & ex.Message, "Credit Card Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        If CreditCardProcessed = False Then
            ' try to 
            MessageBox.Show("Credit card could not be authorized.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            Return True
        End If

    End Function

    Private Function ProcessCreditCardAuthorization(ByVal AUTH_CCPA_NO As String, _
                                                ByVal ChargeAmount As Double, _
                                                ByVal freightAmount As Decimal, _
                                                ByVal salesTax As Decimal) As String

        Dim sql As String = String.Empty
        Dim ORDR_UNIT_PRICE As Decimal = 0

        AUTH_CCPA_NO = AUTH_CCPA_NO.Trim
        If AUTH_CCPA_NO.Length = 0 Then Return String.Empty

        If ChargeAmount <= 0 Then Return String.Empty
        ChargeAmount = Math.Round(ChargeAmount, 2)

        ASCMAIN1.Progress("Processing Credit Card", String.Empty)

        MyBase.Fill_Records("ARTCCPA1", AUTH_CCPA_NO)
        MyBase.Fill_Records("ARTCCPDA", AUTH_CCPA_NO)

        If dst.Tables("ARTCCPA1").Rows.Count <> 1 Then Return String.Empty
        If dst.Tables("ARTCCPDA").Rows.Count <> 1 Then Return String.Empty

        Dim rowARTCCPA1_AUTH As DataRow = dst.Tables("ARTCCPA1").Rows(0)

        Dim AUTH_RESPONSE_APPROVAL_CODE As String = (rowARTCCPA1_AUTH.Item("RESPONSE_APPROVAL_CODE") & String.Empty).ToString.Trim
        If AUTH_RESPONSE_APPROVAL_CODE.Length = 0 Then Return String.Empty

        Dim CCPA_NO As String = String.Empty

        Try
            Me.CreditCardProcessor = New TAC.TAFCARDF(Me)

            '******************************************************************************************************************
            ' Default to Authorize
            ' If it is the case the Invoice Total Amount is greater than the Approved Amount
            ' Then Void the Original Auth and process as a Sale.
            ' If the Void Authorize fails, continue and process as a sale.
            Dim chargeType As String = "A"

            Dim OriginalAuthAmount As Decimal = Val(rowARTCCPA1_AUTH.Item("CCPA_AMT") & String.Empty)
            If ChargeAmount > OriginalAuthAmount Then
                chargeType = "S"
                Try
                    CreditCardProcessor.objCCProcessor.VoidTransaction(rowARTCCPA1_AUTH.Item("TRANS_ID") & String.Empty, "1")
                Catch ex As Exception
                End Try
            End If
            '******************************************************************************************************************

            With Me.CreditCardProcessor
                .ORDR_NO = rowARTCCPA1_AUTH.Item("ORDR_NO") & String.Empty

                .objCCProcessor.TransactionNumber = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
                .objCCProcessor.TransactionAmount = ChargeAmount
                .objCCProcessor.CustomerCreditCard.CardNumber = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO") & String.Empty

                Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                CUST_CREDIT_CARD_EXP_DATE = CUST_CREDIT_CARD_EXP_DATE.PadRight(4, "0")
                .objCCProcessor.CustomerCreditCard.CardExpMonth = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2)
                .objCCProcessor.CustomerCreditCard.CardExpYear = CUST_CREDIT_CARD_EXP_DATE.Substring(2)
                .objCCProcessor.ValidateCard()


                If chargeType = "S" Then
                    .objCCProcessor.CreditCardProcessingNo = CCPA_NO
                    .objCCProcessor.InternalReference = "Customer: " & rowARTCCPA1_AUTH.Item("CUST_CODE") & ", TransType: " & "S"

                    .objCCProcessor.CustomerCreditCard.CardHolderFirstName = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderLastName = "" 'Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
                    .objCCProcessor.CustomerCreditCard.CardHolderAddress = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ADDR1") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderCity = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_CITY") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderState = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_STATE") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderZipCode = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ZIP_CODE") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderCountry = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_COUNTRY") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderTelephone = "" 'Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
                    .objCCProcessor.CustomerCreditCard.CardCVVData = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty
                End If

                With .objCCProcessor.Level2Data
                    .Clear()

                    .CardType = CreditCardProcessor.objCCProcessor.CreditCardType

                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO")) ' dst.Tables("SOTORDR1").Rows(0)
                    Dim rowSHIPTO As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)

                    If rowSHIPTO IsNot Nothing Then
                        .DestinationZip = rowSHIPTO.Item("CUST_ZIP_CODE") & String.Empty
                        .DestinationState = rowSHIPTO.Item("CUST_STATE") & String.Empty
                    End If

                    .DiscountAmount = 0
                    .FreightAmount = freightAmount
                    .InvoiceNumber = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .OrderDate = rowSOTORDR1.Item("ORDR_DATE") & String.Empty
                    .PurchaseIdentifier = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .TaxAmount = salesTax

                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 IsNot Nothing Then
                        .ShipFromZip = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                    End If

                End With

                .rowARTCCPA1 = rowARTCCPA1_AUTH
                If chargeType = "S" Then
                    CCPA_NO = .CC_Sale(ChargeAmount)
                Else
                    CCPA_NO = .CC_Capture(ChargeAmount)
                End If

            End With

        Catch ex As Exception
            MessageBox.Show("The following error occurred processing a credit card: " & ex.Message, "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CreditCardProcessor.Dispose()
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

        Return CCPA_NO

    End Function

    Private Sub grdSOTSHIP1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSHIP1.ClickCellButton
        If grdSOTSHIP1.ActiveCell.Column.Key = "SHIP_TOTAL_WGT_CALC" Then
            If MessageBox.Show("Do you want to copy the Calculated Weight to the Actual Weight?", "Weight", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            grdSOTSHIP1.ActiveRow.Cells("SHIP_TOTAL_WGT").Value = grdSOTSHIP1.ActiveRow.Cells("SHIP_TOTAL_WGT_CALC").Value

        End If
    End Sub

    Private Sub UpdateOrderPicksAndCreateInvoices()

        ' Update Sotordr1 and Sortordr2 and Possiblt SOTPICK1,2 
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "", DataViewRowState.CurrentRows)
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                    Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO")

                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) - Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                    If rowSOTORDR2.Item("ORDR_QTY_PICK") < 0 Then
                        rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    ElseIf rowSOTORDR2.Item("ORDR_QTY_PICK") > 0 Then
                        ' No back Orders
                        rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                        rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    End If
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    rowSOTORDR2.Item("ORDR_QTY_SHIP") = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + (Val(rowSOTPICK2.Item("PICK_QTY_CANC") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_BACK") & String.Empty))
                    rowSOTORDR2.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                    rowSOTORDR2.Item("ORDR_STATUS") = "F"
                Next

                ' Finalize Sales Order and Pick Ticket
                rowSOTORDR1.Item("ORDR_STATUS") = "F"
                rowSOTORDR1.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTORDR1.Item("ORDR_FOB") = rowSOTSHIP1.Item("SHIP_FOB")

                If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text Then
                    Dim row As DataRow = dst.Tables("TATEVNT1").Rows.Add
                    row.Item("TABLE_NAME") = "SOTORDR1"
                    row.Item("TABLE_KEY") = ORDR_NO
                    row.Item("INIT_DATE") = DATETIME_STAMP
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("EVENT_TYPE") = "SHPMTC"
                    row.Item("EVENT_DESC") = "Ship Via was changed from " _
                        & rowSOTORDR1.Item("SHIP_VIA_CODE") & " to " & MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                    row.Item("EVENT_KEY") = ""
                End If
            Next
        Next

        ' Create Invoice Records
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)

            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            rowSOTSHIP1.Item("INV_DATE") = MyBase.Absx1.dteFor("SHIPPED_ACTUAL").Value
            rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = MyBase.Absx1.dteFor("SHIPPED_ACTUAL").Value
            rowSOTSHIP1.Item("SHIP_FOB") = txtSHIP_FOB.Text.Trim

            Dim SOCINVH1 As New TAC.SOCINVH1(dst.Tables("SOTINVH1"), dst.Tables("SOTINVH2"), _
                                              dst.Tables("SOTPICK1"), dst.Tables("SOTPICK2"), _
                                              dst.Tables("ARTOPEN1"), dst.Tables("SOTSHIP1"), _
                                              dst.Tables("SOTORDR5"))
            SOCINVH1.CreateInvoices(SHIP_BOL_NO)
        Next

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            rowSOTORDR1.Item("ORDR_DATE_CLOSED") = rowSOTINVH1.Item("INV_DATE_SHIPPED")
            rowSOTORDR1.Item("ORDR_YYYYPP_CLOSED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
        Next


        If CURR_CODE = "" Or Val(CURR_EXCH_RATE) = 0 Then
            MessageBox.Show("*****************************************" & Environment.NewLine & _
                            "Please contact ABS about this shipment." & Environment.NewLine & _
                            "Let them know the Currency Code is blank!" & Environment.NewLine & _
                            "*****************************************", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If


    End Sub


End Class