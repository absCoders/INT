Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing.Printing

Public Class SOFPICK0

#Region "Declarations"

    Dim REPRINT_CONFIRMED As Boolean
    Dim SOTPICKX As String

    Dim SHIP_BOL_NOs As String = ""
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim SOTORDR0 As String = ""
    Dim SOTPICK1 As String = ""

    Dim options As New Dictionary(Of String, Boolean)
    Dim CUST_CODEs_856 As New List(Of String)
    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim sqlSOTCART1 As String
    Dim printingPickTickets As Boolean = False

    Dim ORDR_NO_MT As String
    Private sqlDerelease As String = String.Empty
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Refresh_SOTPICKX("")

        Get_PARM("SOTPARM1")
        SO_PARM_UPC_VENDOR_ID = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID")

        AUDIT.Add("ARTCUST2", "E")

        With dst

            ASCMAIN1.sql = "Select * from " & SOTPICKX
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False)
            .Tables("SOTPICKX").Columns("SHIP_BOL_NO").AllowDBNull = True
            'WHY IN THE WORLD WOULD THIS BE NULL?
            With .Tables("SOTPICKX").Columns
                .Add("SELECTED", GetType(System.String))
                .Add("OPT_PICK_TICKET", GetType(System.String))
                .Add("OPT_UCC128", GetType(System.String))
                .Add("OPT_PULL_STORE", GetType(System.String))
                .Add("OPT_PULL_STYLE", GetType(System.String))
                .Add("OPT_MANIFEST", GetType(System.String))
                .Add("CUST_856", GetType(System.String))
            End With
            .Tables("SOTPICKX").Columns("SELECTED").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO" & vbCrLf _
                & ", SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTSHIP1" & vbCrLf _
                & " where SOTORDR1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = :PARM2" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_BATCH_NO = :PARM3" & vbCrLf _
                & "   and SOTSHIP1.SHIP_ADDR_TYPE = :PARM4" & vbCrLf _
                & "   and (SOTSHIP1.SHIP_ADDR_TYPE = 'MK' or SOTSHIP1.SHIP_ADDR_CODE = :PARM5)"
            Create_TDA(.Tables.Add, "SOTPICKY", "**", 0, False, "VVVVV", 1)

            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & ", SOTORDR0.CUST_CODE" & vbCrLf _
                & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" & vbCrLf _
                & " from SOTSHIP1,SOTORDR0" & vbCrLf _
                & " where SOTSHIP1.ORDR_GROUP_NO = :PARM1" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "V", 1, "SHIP_PICK_PRINTED")
            ' .Tables("SOTSHIP1").Columns.Add("SHIP_BOL_NO_BC")

            ASCMAIN1.sql = "Select * from SOTORDR0 where ROWNUM <1"
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")

            ASCMAIN1.sql = "Select SOTORDR0.* from " & SOTORDR0 & " SOTORDR0 "
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO" _
                & " from SOTORDR1," & SOTORDR0 & " SOTORDR0" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*, SOTORDR1.WHSE_CODE" & vbCrLf _
               & " from SOTORDR5,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
               & " where SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
               & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, True, "", 2)

            Create_TDA(.Tables.Add, "SOTLABL1", "*", 1, False, "", 2)

            Create_TDA(.Tables.Add, "SOTPICK0", "*", 1, False)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)
            Create_TDA(.Tables.Add, "ARTCUST2", "*")
            .Tables("ARTCUST2").Columns.Add("CUST_ADDR_DC")
            .Tables("ARTCUST2").Columns.Add("DC_STATE")

            ASCMAIN1.sql = "Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.PICK_NO, 0 LABEL_COUNTER" & vbCrLf _
                & ", SOTPICK0.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DEPT, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO" & vbCrLf _
                & ", ITEM_CODE ITEM_CODE_1, ITEM_CODE ITEM_CODE_2, ITEM_CODE ITEM_CODE_3" & vbCrLf _
                & " from SOTPICK1,SOTPICK0,SOTORDR1,SOTORDR2" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK3", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select EDT850T2.*" & vbCrLf _
                & " from EDT850T2 where EDI_DOC_SEQ_NO in (Select DISTINCT EDI_DOC_SEQ_NO" & vbCrLf _
                & " from SOTORDR1 where ORDR_GROUP_NO in (Select DISTINCT ORDR_GROUP_NO from " & SOTORDR0 & " SOTORDR0))"
            Create_TDA(.Tables.Add, "EDT850T2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select EDT850T1.*" & vbCrLf _
                & " from EDT850T1 where EDI_DOC_SEQ_NO in (Select DISTINCT EDI_DOC_SEQ_NO" & vbCrLf _
                & " from SOTORDR1 where ORDR_GROUP_NO in (Select DISTINCT ORDR_GROUP_NO from " & SOTORDR0 & " SOTORDR0))"
            Create_TDA(.Tables.Add, "EDT850T1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from SOTPICK1 where ROWNUM <1"
            SOTPICK1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1 & " Add Primary Key (PICK_NO)")

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", TRIM(SUBSTR(LPAD(SOTORDR1.CUST_STORE_NO,6,' '),3,4)) CUST_STORE_NO4" & vbCrLf _
                & " from " & SOTPICK1 & " SOTPICK1, SOTORDR1, ARTCUST2" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE and ARTCUST2.CUST_STORE_NO (+) = SOTORDR1.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "", 1)
            ' .Tables("SOTPICK1").Columns.Add("PICK_NO_BC", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CART_SERIAL_NO", GetType(System.Int32))

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, ICTITEM1.CARTON_PACK_QTY, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " from SOTPICK2," & SOTPICK1 & " SOTPICK1, SOTORDR2, ICTITEM1" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                & "   and SOTPICK2.PICK_QTY <> 0"  'To avoid picking up records representing a cancellation or backorder generated during Pick Ticket Release
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " from SOTPICK2, SOTPICK1, SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK2.PICK_QTY <> 0" & vbCrLf _
                & "   and SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICKL", "**", 0, False, "V", 2)
            .Tables("SOTPICKL").Columns.Add("LABEL_QTY", GetType(System.Int32))

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            .Tables("SOTPICK1").Columns.Add("PICK_TOT", GetType(System.Int32), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")

            sqlSOTCART1 = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTPICK1.ORDR_NO" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,11,9) CART_NO_9" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_DIGIT" & vbCrLf _
                & ", SUBSTR(SOTCART1.CART_NO,5,6) CART_NO_PFX" & vbCrLf _
                & ", '(00) 0 0 ' || SUBSTR(SOTCART1.CART_NO,5,6) || ' ' || SUBSTR(SOTCART1.CART_NO,11,9) || SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_FMT" & vbCrLf
            ASCMAIN1.sql = sqlSOTCART1 _
                & " from SOTCART1," & SOTPICK1 & " SOTPICK1" _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO"

            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "", 1)
            .Tables("SOTCART1").Columns.Add("CART_1_OF_9", GetType(System.String))
            .Tables("SOTCART1").Columns.Add("ITEM_CODE")
            .Tables("SOTCART1").Columns.Add("ITEM_DESC")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")

            With .Tables("SOTCART1").Columns
                .Add("CART_SEQ_MAX", GetType(System.Int32))
                Create_Relation("SOTORDR1", "SOTCART1", "ORDR_NO")

                .Add("CUST_STORE_NO", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_STORE_NO")
                .Add("CUST_ZIP_CODE", GetType(System.String))
                .Add("CART_SERIAL_NO", GetType(System.Int32))
                .Add("CUST_CODE", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_CODE")
                .Add("CUST_SELECTED", GetType(System.String))
                .Add("CUST_DC_NO", GetType(System.String), "PARENT(SOTORDR1_SOTCART1).CUST_DC_NO")
            End With

            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & " from SOTCART2,SOTCART1," & SOTPICK1 & " SOTPICK1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "", 2)

            Create_Relation("SOTORDR2", "SOTCART2", "ORDR_NO,ORDR_LNO")
            dst.Tables("SOTORDR2").Columns.Add("CART_COUNT", GetType(System.Int32), "COUNT(CHILD(SOTORDR2_SOTCART2).CART_LNO)")


            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_BIN from ICTITEM1" & vbCrLf _
                & " where ITEM_CODE in" & vbCrLf _
                & "(Select Distinct SOTORDR2.ITEM_CODE from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO)"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
            .Tables("ICTITEM1").Columns.Add("LOCATION_CODES")

            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1" & vbCrLf _
                & " where ITEM_CODE in" & vbCrLf _
                & "(Select Distinct SOTORDR2.ITEM_CODE from SOTORDR2,SOTORDR1," & SOTORDR0 & " SOTORDR0" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO)" & vbCrLf _
                & "  and WHTLOCB1.LOCATION_QTY > 0"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "", 4)

            For Each TABLE_NAME As String In New String() {"SOTSDIV1", "ICTWHSE1", "SOTSVIA1", "TATTERM1", "SOTUCCL1"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            ASCMAIN1.sql = "Select X.WHSE_CODE, ICTWHSE1.WHSE_DESC" & vbCrLf _
                & ", X.ORDR_CNT_PICK, X.ORDR_AMT_PICK" & vbCrLf _
                & ", Y.ORDR_CNT_PICK_UNPRINTED, Y.ORDR_AMT_PICK_UNPRINTED" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTORDR0.WHSE_CODE, Sum(ORDR_CNT_PICK) ORDR_CNT_PICK, Sum (ORDR_AMT_PICK) ORDR_AMT_PICK " & vbCrLf _
                & "  from SOTORDR0" & vbCrLf _
                & "where SOTORDR0.ORDR_CNT_PICK <> 0" & vbCrLf _
                & " group by SOTORDR0.WHSE_CODE) X" & vbCrLf _
                & ", (Select WHSE_CODE, Count (Distinct SOTPICK1.PICK_NO) ORDR_CNT_PICK_UNPRINTED" & vbCrLf _
                & ", Sum (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) ORDR_AMT_PICK_UNPRINTED" & vbCrLf _
                & "from SOTSHIP1,SOTPICK1,SOTPICK2 where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & " and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                & " and SOTSHIP1.SHIP_PICK_PRINTED is Null" & vbCrLf _
                & "group by SOTSHIP1.WHSE_CODE" & vbCrLf _
                & ") Y" & vbCrLf _
                & "  where ICTWHSE1.WHSE_CODE = X.WHSE_CODE " & vbCrLf _
                & "    and Y.WHSE_CODE (+) = X.WHSE_CODE"

            ' 12/18/2015 - Ship from non 3pl warehouse
            If ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA" Then
                ASCMAIN1.sql &= " AND  X.WHSE_CODE IN (Select WHSE_CODE from ICTWHSE1 WHERE LP_WHSE_ID IS NULL)"
            End If
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("ORDR_CNT_PICK").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")

            sqlDerelease = "Select SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO from SOTPICK1, SOTORDR1 where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTDREL1", sqlDerelease, 0, False, String.Empty, 0)

        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKL.DataSource = dst.Tables("SOTPICKL")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")

        With grdSOTPICKX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each COLUMN_NAME As String In New String() {"SELECTED", "SHIP_BOL_NO", "ORDR_GROUP_NO", "PICK_BATCH_NO", "CUST_CODE", "CUST_CODE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"SELECTED", "OPT_PICK_TICKET", "OPT_UCC128", "OPT_PULL_STORE", "OPT_PULL_STYLE", _
                                  "OPT_MANIFEST"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSOTPICKL.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"LABEL_QTY"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Color.Beige
                End If
            Next
        End With

        Create_Summary(grdSOTPICKX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTPICKX, New String() {"SELECTED", "ORDR_CNT_PICK", "ORDR_QTY_PICK", "ORDR_AMT_PICK", "ORDR_CNT_CART"})
        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKL, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICKL, New String() {"PICK_QTY", "LABEL_QTY"})

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, New String() {"ORDR_CNT_PICK", "ORDR_AMT_PICK", "ORDR_CNT_PICK_UNPRINTED", "ORDR_AMT_PICK_UNPRINTED"})

        Show_Filter(grdSOTPICKX, True)
        grdSOTPICKX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.sql = "Select EDTTRPM1.CUST_CODE from EDTTRPM1 " & vbCrLf _
                & " where (EDI_STATUS = 'P' or EDI_STATUS = 'T')" & vbCrLf _
                & "   and EDI_DOC_NO = '856'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            CUST_CODEs_856.Add(CUST_CODE)
        Next

        lblDefaultPrinter.Text = Default_Printer()

        Bind_Controls(grpSHIPTO, "SOTORDR5")
        'Bind_Controls(grpSHIPTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows))

        ASCMAIN1.Add_Value_List(grdSOTPICKX, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyboard", "W:Web", "E:EDI", "S:SRep"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTOREL1", Absx1.txtFor("WHSE_CODE").Text) Then Exit Sub
                End If

            Case "Refresh"

            Case "Print"
                Dim ORDR_GROUP_NOs As String = Get_Selected_ORDR_GROUP_NOs()

                If dst.Tables("SOTPICKX").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "You Must First Select Something to Print"
                End If

                ORDR_GROUP_NOs = ""

            Case "Update"

            Case "Print UCC128 Labels"

            Case "Print Address Labels"
                If numLabels.Value = 0 Then
                    EMsg &= vbCr & "Number of Address Labels is 0"
                End If

            Case "Print Address/Item Labels", "Print Content Labels"
                If dst.Tables("SOTPICKL").Select("LABEL_QTY > 0").Length = 0 Then
                    EMsg &= vbCr & "You must enter a label qty for at least one item"
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                If ScreenMode Then
                    Refresh_SOTPICKX("")
                Else
                    Refresh_ICTWHSEX()
                End If

            Case "Print"
                EntryMode = "E"
                printingPickTickets = True

                Print_Pick_Tickets("")
                Print_Documents()
                If optBOL.Value = "0" Then
                    Update_Record()
                    Refresh_SOTPICKX("")
                End If

                printingPickTickets = False

                Setup_SOTSHIP1()

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Re-Print Confirmed"
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                        Dim CUST_CODE As String = ASCMAIN1.CodeSelector.SelectedCode
                        Refresh_SOTPICKX(CUST_CODE)
                        REPRINT_CONFIRMED = True
                    End If
                End If

            Case "Print UCC128 Labels"
                If grdSOTCART1.Selected.Rows.Count = 0 Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
                        grow.Selected = True
                    Next
                End If
                If grdSOTCART1.Selected.Rows.Count <> 0 Then
                    Print_UCC128_Labels()
                End If

            Case "Print Address Labels"
                Print_Address_Labels(Val(numLabels.Value & ""), txtComment.Text, "ADDRESS")

            Case "Print Address/Item Labels"
                Print_Address_Labels_from_Details("ADDRESS")

            Case "Print Content Labels"
                Fill_Records("SOTUCCL1")
                Print_Address_Labels_from_Details("CONTENT")
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                '.Items("Refresh").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                .Items("Re-Print Confirmed").Settings.Enabled = iScreenMode
                .Items("Re-Print Confirmed").Visible = False
            End With
            .Groups("Selection Options").Visible = ScreenMode
            .Groups("Label Printing").Visible = ScreenMode
        End With

        grdICTWHSEX.Visible = Not ScreenMode
        grdSOTPICKX.Visible = ScreenMode

        spl.Panel1Collapsed = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Setup_Print_Option()

        chkDirect2Printer.Visible = Not ScreenMode
        lblDefaultPrinter.Visible = Not ScreenMode
        Set_Read_Only_for_ctl(chkDirect2Printer, False)

        If ScreenMode Then
            Setup_tabLabels()
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTPICK1", "SOTSHIP1", "SOTPICKX", _
             "SOTCART2", "SOTPICK2", "SOTCART1", "SOTORDR5", _
             "SOTORDR0", "EDT850T2", "EDT850T1", "ICTITEM1", "WHTLOCB1", _
             "SOTLABL1", "SOTPICK0", "ARTCUST1", "ARTCUST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        WHSE_CODE = ""
        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
        End If
        Refresh_ICTWHSEX()

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Refresh_SOTPICKX("")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Pick_Tickets(Optional PICK_NOs As String = "")

        ASCMAIN1.Progress("Now Preparing to Print Documents", "Please Wait ...")

        Set_Options()

        EnforceConstraints(False)
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTPICK1)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR0", "SOTORDR1", "SOTORDR2", "SOTORDR5", "ICTITEM1", "WHTLOCB1", _
             "SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2", _
             "EDT850T1", "EDT850T2", "SOTPICK0", "ARTCUST1", "ARTCUST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        SHIP_BOL_NOs = ""

        If PICK_NOs <> "" Then
            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where PICK_NO in (" & PICK_NOs & ")"
            ASCDATA1.ExecuteSQL("Insert into " & SOTPICK1 & " " & ASCMAIN1.sql)
            ASCMAIN1.AnalyzeTable(SOTPICK1)

            Dim ORDR_GROUP_NO As String = grdSOTPICKX.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim ORDR_GROUP_NOs As String = "'" & ORDR_GROUP_NO & "'"

            Fill_Records("SOTSHIP1", ORDR_GROUP_NO)

            Customer_Data(ORDR_GROUP_NOs)
            Get_SOTORDRx(ORDR_GROUP_NOs)
            Get_SOTPICKx()
            Get_EDI_Data()
        Else
            Dim g As Integer = 0
            Dim ORDR_GROUP_NOs As String = ""

            Dim PICK_BATCH_NOs As String = ""

            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                Dim ORDR_GROUP_NO As String = rowSOTPICKX.Item("ORDR_GROUP_NO")
                Dim SHIP_BOL_NO As String = rowSOTPICKX.Item("SHIP_BOL_NO") & ""
                Dim SHIP_ADDR_TYPE As String = rowSOTPICKX.Item("SHIP_ADDR_TYPE") & ""
                Dim SHIP_ADDR_CODE As String = rowSOTPICKX.Item("SHIP_ADDR_CODE") & ""
                Dim PICK_BATCH_NO As String = rowSOTPICKX.Item("PICK_BATCH_NO") & ""
                Dim CUST_CODE As String = rowSOTPICKX.Item("CUST_CODE")

                g = g + 1
                ASCMAIN1.Progress("-", CStr(g) & ":" & CUST_CODE & "/" & ORDR_GROUP_NO)

                If InStr(ORDR_GROUP_NOs, ORDR_GROUP_NO) = 0 Then
                    ORDR_GROUP_NOs = ORDR_GROUP_NOs & ", '" & ORDR_GROUP_NO & "'"
                End If
                If InStr(PICK_BATCH_NOs, PICK_BATCH_NO) = 0 Then
                    PICK_BATCH_NOs = PICK_BATCH_NOs & ", '" & PICK_BATCH_NO & "'"
                    Fill_Records("SOTPICK0", PICK_BATCH_NO, False)

                    ASCMAIN1.sql = "Select SOTSHIP1.*, DECODE (SHIP_ADDR_TYPE,'DC',SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" _
                        & " from SOTSHIP1 where PICK_BATCH_NO = '" & PICK_BATCH_NO & "'"
                    Fill_Records("SOTSHIP1", PICK_BATCH_NO, False, ASCMAIN1.sql)
                End If

                If SHIP_BOL_NO <> "" Then
                    Get_SHIP_BOL_NO(SHIP_BOL_NO, SHIP_BOL_NOs, REPRINT_CONFIRMED)
                Else
                    ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" _
                        & " from SOTSHIP1 " _
                            & " where SOTSHIP1.PICK_BATCH_NO = '" & PICK_BATCH_NO & "'" _
                            & "   and SOTSHIP1.SHIP_ADDR_TYPE = 'MK'" _
                            & "   and SOTSHIP1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    If REPRINT_CONFIRMED Then
                        ' WE WOULD ALSO WANT STATUS = 'F' IF RESURRECTING FROM THE DEAD
                        ASCMAIN1.sql &= "   and (SOTSHIP1.SHIP_STATUS = 'P' OR SOTSHIP1.SHIP_STATUS = 'F')"
                    Else
                        ASCMAIN1.sql &= "   and SOTSHIP1.SHIP_STATUS = 'P'"
                    End If
                    For Each rowSOTSHIP1 As DataRow In ASCDATA1.GetDataTable.Rows
                        SHIP_BOL_NO = rowSOTSHIP1.Item("SHIP_BOL_NO")
                        Get_SHIP_BOL_NO(SHIP_BOL_NO, SHIP_BOL_NOs, REPRINT_CONFIRMED)
                    Next
                End If
            Next

            PICK_BATCH_NOs = Mid$(PICK_BATCH_NOs, 2)
            ORDR_GROUP_NOs = Mid$(ORDR_GROUP_NOs, 2)

            Customer_Data(ORDR_GROUP_NOs)
            Get_SOTORDRx(ORDR_GROUP_NOs)
            Get_SOTPICKx()
            Get_EDI_Data()

            ASCMAIN1.Progress("")
        End If

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("CART_SERIAL_NO") = 0
        Next

    End Sub

    Sub Get_SHIP_BOL_NO(SHIP_BOL_NO As String, ByRef SHIP_BOL_NOs As String, REPRINT_CONFIRMED As Boolean)
        If InStr(SHIP_BOL_NOs, SHIP_BOL_NO) = 0 Then
            SHIP_BOL_NOs = SHIP_BOL_NOs & ",'" & SHIP_BOL_NO & "'"
        End If

        ASCMAIN1.sql = " Select * from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        If REPRINT_CONFIRMED Then
            ' WE WOULD ALSO WANT STATUS = 'F' IF RESURRECTING FROM THE DEAD
            ASCMAIN1.sql &= "   and (SOTPICK1.PICK_STATUS = 'P' OR SOTPICK1.PICK_STATUS = 'F')"
        Else
            ASCMAIN1.sql &= "   and SOTPICK1.PICK_STATUS = 'P'"
        End If

        ASCMAIN1.sql = "Insert into " & SOTPICK1 & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Customer_Data(ORDR_GROUP_NOs As String)
        ASCMAIN1.sql = "Select ARTCUST1.* from ARTCUST1" _
            & " where CUST_CODE in (" _
            & " Select DISTINCT CUST_CODE from SOTORDR0" _
            & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"
        Fill_Records("ARTCUST1", "", False, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select ARTCUST2.*" _
            & " from ARTCUST2" _
            & " where CUST_CODE in (" _
            & " Select DISTINCT CUST_CODE from SOTORDR0" _
            & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"
        Fill_Records("ARTCUST2", "", False, ASCMAIN1.sql)
    End Sub

    Sub Get_SOTORDRx(ORDR_GROUP_NOs As String)
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " Select * from SOTORDR0 where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & ")")
        ASCMAIN1.AnalyzeTable(SOTORDR0)

        Fill_Records("SOTORDR0", "", False)
        Fill_Records("SOTORDR1", "", False)
        Fill_Records("SOTORDR2", "", False)
        Fill_Records("SOTORDR5", "", False)

        Fill_Records("ICTITEM1", "", False)

        Fill_Records("WHTLOCB1", "", False)

        For Each rowICTITEM1 As DataRow In dst.Tables("ICTITEM1").Select("")
            Dim LOCATION_CODES As String = ""
            Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE")
            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("ITEM_CODE = '" & ITEM_CODE & "'", "LOCATION_QTY DESC")
                LOCATION_CODES &= "," & rowWHTLOCB1.Item("LOCATION_CODE")
            Next
            rowICTITEM1.Item("LOCATION_CODES") = Mid(LOCATION_CODES, 2)
        Next
    End Sub

    Sub Get_SOTPICKx()
        Fill_Records("SOTPICK1", "", False)
        'For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
        '    rowSOTPICK1.Item("PICK_NO_BC") = BC_OCode128(Trim$(rowSOTPICK1.Item("PICK_NO")), 1, 0, 0)
        'Next

        Fill_Records("SOTPICK2", "", False)
        Fill_Records("SOTCART1", "", False)
        Carton_Serialization()
        Fill_Records("SOTCART2", "", False)
        SetCartItemCode()
    End Sub

    Sub Get_EDI_Data()
        Fill_Records("EDT850T2", "", False)
        Fill_Records("EDT850T1", "", False)
    End Sub

    Sub Carton_Serialization()
        ASCMAIN1.Progress("-", "Carton Serialization")

        For Each row As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTCART1"), New String() {"PICK_NO"}).Rows
            Dim PICK_NO As String = row.Item("PICK_NO")
            Dim sqlw As String = "PICK_NO = '{0}'"
            sqlw = String.Format(sqlw, PICK_NO)
            Dim CART_SEQ_MAX As Int32 = Val(dst.Tables("SOTCART1").Select(sqlw).Length)
            Dim CART_SERIAL_NO As Integer = 0
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(sqlw, "CART_NO")
                CART_SERIAL_NO += 1
                rowSOTCART1.Item("CART_SERIAL_NO") = CART_SERIAL_NO
                rowSOTCART1.Item("CART_SEQ_MAX") = CART_SEQ_MAX
                rowSOTCART1.Item("CART_1_OF_9") = CStr(CART_SERIAL_NO) & " of " & CStr(CART_SEQ_MAX)
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTCART1"), New String() {"PICK_NO"}).Select("")
            Dim PICK_NO As String = row.Item("PICK_NO")
            Dim sqlw As String = "PICK_NO = '" & PICK_NO & "'"
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = Val(dst.Tables("SOTCART1").Compute("Count(CART_NO)", sqlw) & "")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = Val(dst.Tables("SOTCART1").Compute("Sum(CART_TOTAL_WGT_CALC)", sqlw) & "")
        Next
    End Sub

    Sub Update_Record()

        BeginTrans()

        If SHIP_BOL_NOs <> "" Then
            ASCMAIN1.sql = "Update SOTPICK1 " _
                & " Set PICK_PRINTED = SYSDATE, PICK_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'" _
                & " where SHIP_BOL_NO in (" & Mid(SHIP_BOL_NOs, 2) & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                     & " Select 'SOTORDR1', ORDR_NO, PICK_PRINTED, PICK_PRINTED_OPER, 'PICKTP','Pick Ticket Print', NULL from SOTPICK1" _
                     & " where SHIP_BOL_NO in (" & Mid(SHIP_BOL_NOs, 2) & ")"
            ASCDATA1.ExecuteSQL()

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO in (" & Mid(SHIP_BOL_NOs, 2) & ")")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                If rowSOTSHIP1.Item("SHIP_PICK_PRINTED") & "" = "" Then
                    rowSOTSHIP1.Item("SHIP_PICK_PRINTED") = DATETIME_STAMP
                End If
            Next
            Update_Record_TDA("SOTSHIP1")
        End If

        CommitTrans("")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKX, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select All X", "Show Pick Tickets", "Carton Summary", "De-Release Shipment", "De-Release Selected Shipments")
        Load_Popup_Menu(grdSOTPICK1, "BBBB", "Select All", "De-Select All", "Sales Order Inquiry", "De-Release Pick Ticket")
        Load_Popup_Menu(grdSOTCART1, "BBB", "Select All", "De-Select All", "Print UCC128 Labels")
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

                Case "grdSOTPICKX"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All X"), UltraWinToolbars.ButtonTool)
                    If grdSOTPICKX.ActiveCell Is Nothing OrElse _
                            (grdSOTPICKX.ActiveCell.Value & "" = "" _
                             Or Not New String() {"ORDR_GROUP_NO", "CUST_CODE", "PICK_BATCH_NO"}.Contains(grdSOTPICKX.ActiveCell.Column.Key)) Then
                        tlb_btn.SharedProps.Visible = False
                        tlb_btn.Tag = ""
                    Else
                        tlb_btn.Tag = grdSOTPICKX.ActiveCell.Column.Key & " = '" & grdSOTPICKX.ActiveCell.Value & "'"
                        tlb_btn.SharedProps.Caption = "Select All " & grdSOTPICKX.ActiveCell.Column.Header.Caption & " = " & grdSOTPICKX.ActiveCell.Value
                        ' tlb_btn.SharedProps.Caption = "Select All " & grdSOTPICKX.ActiveCell.Value
                        tlb_btn.SharedProps.Visible = True
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("De-Release Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = chkDeRelease.Checked
                    tlb_btn = DirectCast(tlb_pop.Tools("De-Release Selected Shipments"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = chkDeRelease.Checked

                Case "grdSOTPICK1"
                    tlb_btn = DirectCast(tlb_pop.Tools("De-Release Pick Ticket"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = chkDeRelease.Checked


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All", "Select All X"

                If grd.Name = "grdSOTPICK1" Or grd.Name = "grdSOTCART1" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Selected = (e.Tool.Key = "Select All")
                    Next
                Else
                    If e.Tool.Key = "Select All X" Then
                        Dim sqlw As String = IIf(e.Tool.Key = "Select All X", e.Tool.Tag, "")
                        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select(sqlw)
                            rowSOTPICKX.Item("SELECTED") = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                        Next
                    Else
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                            grow.Cells("SELECTED").Value = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                            grow.Update()
                        Next
                    End If
                    Display_Totals()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Show Pick Tickets"
                Show_Pick_Tickets()

            Case "Print UCC128 Labels"
                If grdSOTCART1.Selected.Rows.Count = 0 Then
                    If grdSOTCART1.ActiveRow IsNot Nothing Then grdSOTCART1.ActiveRow.Selected = True
                End If
                If grdSOTCART1.Selected.Rows.Count <> 0 Then
                    Print_UCC128_Labels()
                End If

            Case "Carton Summary"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Load_SOTPICK1(SHIP_BOL_NO, True)
                Dim RPT As String = "SORPICKS"
                Print_Report_Begin()
                Generate_Report(RPT, "Shipment Carton Summary")
                Print_Report_End()

            Case "De-Release Pick Ticket"
                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Text
                Dim pickList As List(Of String) = New List(Of String)
                pickList.Add(PICK_NO)
                DeRelease_PickTickets(pickList)

            Case "De-Release Selected Shipments"
                Dim pickList As List(Of String) = New List(Of String)
                'For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                '    pickList.Add(grow.Cells("PICK_NO").Value)
                'Next
                For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
                    Dim SHIP_BOL_NO As String = rowSOTPICKX.Item("SHIP_BOL_NO")
                    ASCMAIN1.sql = "Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = :PARM1"
                    For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {SHIP_BOL_NO}).Select("")
                        pickList.Add(row.Item("PICK_NO"))
                    Next
                Next
                If pickList.Count > 0 Then DeRelease_PickTickets(pickList)

            Case "De-Release Shipment"
                Dim pickList As List(Of String) = New List(Of String)
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & grd.ActiveRow.Cells("SHIP_BOL_NO").Text & "'")
                    pickList.Add(rowSOTPICK1.Item("PICK_NO"))
                Next
                DeRelease_PickTickets(pickList)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub

#End Region

    Sub Print_Record()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM   
        'Fill_Records("SOTSVIA1", SHIP_CODE)
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORSHIP1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "Shipper Invoice Report", , , , , False)
        'Print_Report_End()
    End Sub

#Region "grdSOTWHSEX"

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

#End Region

#Region "grdSOTPICKX"

    Private Sub grdSOTPICKX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICKX.AfterRowActivate
        Setup_SOTSHIP1()
    End Sub

    Private Sub grdSOTPICKX_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICKX.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTPICKX_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdSOTPICKX.ClickCell

        If optBOL.Value = "1" Then
            If e.Cell.Column.Key = "QTY_ADDR_LABEL" Then
                If Val(e.Cell.Value & "") > 3 Then
                    e.Cell.Value = 0
                Else
                    e.Cell.Value = Val(e.Cell.Value) + 1
                End If
            End If
            'e.Cell.Row.Cells("SELECTED").Value = "1"
        End If
        ' If e.Cell.Row.DataChanged Then
        e.Cell.Row.Update()
        ' End If
    End Sub

    Private Sub grdSOTPICKX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTPICKX.DoubleClickRow
        'Absx1.txtFor("ORDR_NO").Text = e.Row.Cells("ORDR_NO").Value
        'Click_Command("Load")
    End Sub

#End Region

    Sub Print_Documents()

        Dim RecordSelectionFormula As String = String.Empty
        Print_Report_Begin()

        If options("OPT_PICK_TICKET") Then
            If dst.Tables("SOTPICKX").Select("SELECTED = '1' and OPT_PICK_TICKET = '1' and ORDR_SOURCE <> 'W'").Length > 0 Then Print_Report("SORPICK1", "Pick Tickets")
            If dst.Tables("SOTPICKX").Select("SELECTED = '1' and OPT_PICK_TICKET = '1' and ORDR_SOURCE = 'W'").Length > 0 Then Print_Report("SORPICKW", "Pick Tickets - Web")
        End If

        If options("OPT_MANIFEST") Then
            RecordSelectionFormula = CreateRecordSelectionFormula("OPT_MANIFEST")
            Print_Report("SORPICKS", "Carton Summary", RecordSelectionFormula)
        End If

        If options("OPT_PULL_STYLE") Then
            RecordSelectionFormula = CreateRecordSelectionFormula("OPT_PULL_STYLE")
            Print_Report("SORPICK5", "Pull Sheet by Item", RecordSelectionFormula)
        End If

        '  If options("OPT_PULL_STORE") Then Print_Report("SORPICK4", "Distribution")
        If options("OPT_UCC128") Then
            RecordSelectionFormula = CreateRecordSelectionFormula("OPT_UCC128")

            ' Split is not working properly
            ' RecordSelectionFormula = RecordSelectionFormula.Split(" IN ")(2).Trim.Replace(Chr(34), "'").Replace("[", "(").Replace("]", ")")
            ' RecordSelectionFormula = "SHIP_BOL_NO IN " & RecordSelectionFormula
            Dim inLoc As Int16 = InStr(RecordSelectionFormula, " IN ")
            Dim part1 As String = RecordSelectionFormula.Substring(0, inLoc)
            Dim part2 As String = RecordSelectionFormula.Substring(inLoc + 3)
            part2 = part2.Trim.Replace(Chr(34), "'").Replace("[", "(").Replace("]", ")")
            RecordSelectionFormula = "SHIP_BOL_NO IN " & part2

            Print_UCC128_Labels_for_Selected_Shipments(RecordSelectionFormula)
        End If

        'Dim PRINTER_ID As String = rowPRINTER_ID.Item(0)
        'Dim rowSOTPRNT1 As DataRow = dst.Tables("SOTPRNT1").Rows.Find(PRINTER_ID)
        'Dim PRINTER_PORT As String = rowSOTPRNT1.Item("PRINTER_PORT") & ""
        ''If PRINTER_PORT = "" Then PRINTER_PORT = "\\192.168.130.201\" & PRINTER_ID
        If chkDirect2Printer.Checked Then
            Dim PRINTER_PORT As String = lblDefaultPrinter.Text
            Print_Report_End(True, , PRINTER_PORT)
        Else
            Print_Report_End()
        End If

    End Sub

    Private Function CreateRecordSelectionFormula(ByVal columnName As String) As String
        Dim sqlw As String = "SELECTED = '1' "
        Dim criteria As String = String.Empty
        Dim RecordSelectionFormula As String = String.Empty

        For Each row As DataRow In dst.Tables("SOTPICKX").Select(sqlw & " and " & columnName & " = '1'")
            criteria &= ", " & Chr(34) & row.Item("SHIP_BOL_NO") & Chr(34)
        Next

        If criteria.Length = 0 Then
            Return String.Empty
        End If

        criteria = criteria.Substring(1).Trim
        RecordSelectionFormula = "{SOTSHIP1.SHIP_BOL_NO} IN [" & criteria & "]"

        Return RecordSelectionFormula
    End Function

    Sub Print_Report(xRPT As String, Optional title As String = "", Optional RecordSelectionFormula As String = "")
        Dim RPT As String = xRPT
        'If RPT = "SORPICK3" Then
        '    CR_params.Add("DETAIL", "")
        'End If
        Generate_Report(RPT, title, "", RecordSelectionFormula)
    End Sub

    Sub Set_Options()
        Dim sqlw As String = "SELECTED = '1' and "
        options.Clear()
        With dst.Tables("SOTPICKX")
            For Each OPT As String In New String() _
                {"OPT_PICK_TICKET", "OPT_UCC128", "OPT_PULL_STORE", "OPT_PULL_STYLE", "OPT_MANIFEST"}
                options.Add(OPT, (.Select(sqlw & OPT & " = '1'").Length > 0))
            Next
        End With
    End Sub

    Sub Refresh_SOTPICKX(CUST_CODE_x As String)

        If SOTPICKX <> "" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTPICKX)
        End If

        For i As Integer = 0 To 1 ' 0 : DC, 1 : MK
            If i = 0 Then
                ASCMAIN1.sql = "Select SOTORDR0.ORDR_GROUP_NO, SOTSHIP1.PICK_BATCH_NO" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                    & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                    & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_VIA_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO SHIP_BOL_NO_X" & vbCrLf _
                    & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_ROUTING_INST" & vbCrLf _
                    & ", SOTSHIP1.SHIP_PICK_PRINTED, SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE" & vbCrLf
            Else
                ASCMAIN1.sql = "Select DISTINCT SOTORDR0.ORDR_GROUP_NO, SOTSHIP1.PICK_BATCH_NO" & vbCrLf _
                    & ", SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                    & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                    & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_ADDR_TYPE, NULL SHIP_ADDR_CODE" & vbCrLf _
                    & ", SOTSHIP1.SHIP_VIA_CODE" & vbCrLf _
                    & ", 'MK' SHIP_BOL_NO_X" & vbCrLf _
                    & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_ROUTING_INST" & vbCrLf _
                    & ", SOTSHIP1.SHIP_PICK_PRINTED, SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE" & vbCrLf
            End If

            If SOTPICKX <> "" Then
                ASCMAIN1.sql &= ", 0 ORDR_QTY_PICK, 0 ORDR_AMT_PICK, 0 ORDR_CNT_PICK, 0 ORDR_CNT_CART, NULL ORDR_HIGH_PRIORITY, NULL ORDR_HIGH_PRIORITY_NOTE"
            End If

            ASCMAIN1.sql &= " from SOTSHIP1,SOTORDR0,ARTCUST1" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            If SOTPICKX = "" Then
                ASCMAIN1.sql &= " and ROWNUM < 1"
                SOTPICKX = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add Primary Key (SHIP_BOL_NO)") ' TEST
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_QTY_PICK NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_AMT_PICK NUMBER (13,2)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_CNT_PICK NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_CNT_CART NUMBER (8,0)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_HIGH_PRIORITY VARCHAR2(1)")
                ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add ORDR_HIGH_PRIORITY_NOTE VARCHAR2(30)")
                Exit Sub
            End If

            If CUST_CODE_x = "" Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
            Else
                ' FOR REVIVE-FROM-THE-DEAD SPECIALS
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'F' and SOTORDR0.CUST_CODE = '" & CUST_CODE_x & "' "
            End If

            If i = 0 Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_ADDR_TYPE = 'DC'"
            Else
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_ADDR_TYPE = 'MK'"
            End If
            If optBOL.Value = "0" Then
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is NULL "
            Else
                ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is NOT NULL "
            End If

            ASCDATA1.ExecuteSQL("Insert into " & SOTPICKX & " " & ASCMAIN1.sql)
        Next i

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & " Select SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY) ORDR_QTY_PICK" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
            & ", Max (SOTORDR1.ORDR_HIGH_PRIORITY) ORDR_HIGH_PRIORITY" & vbCrLf _
            & ", Max (SOTORDR1.ORDR_HIGH_PRIORITY_NOTE) ORDR_HIGH_PRIORITY_NOTE " & vbCrLf _
            & " from SOTPICK2,SOTPICK1,SOTORDR1," & SOTPICKX & " SOTPICKX" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTPICKX.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & " group by SOTPICK1.SHIP_BOL_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_QTY_PICK = R1.ORDR_QTY_PICK, ORDR_AMT_PICK = R1.ORDR_AMT_PICK, ORDR_HIGH_PRIORITY = R1.ORDR_HIGH_PRIORITY, ORDR_HIGH_PRIORITY_NOTE = R1.ORDR_HIGH_PRIORITY_NOTE" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_CNT_PICK = " & vbCrLf _
            & "    (Select Count (*) from SOTPICK1 where SHIP_BOL_NO = R1.SHIP_BOL_NO and SOTPICK1.PICK_STATUS = 'P')" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "   Update " & SOTPICKX & " Set ORDR_CNT_CART = " & vbCrLf _
            & "    (Select Count (*) from SOTCART1,SOTPICK1 where SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTPICK1.SHIP_BOL_NO = R1.SHIP_BOL_NO and SOTPICK1.PICK_STATUS = 'P')" & vbCrLf _
            & "    where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Me.Cursor = Cursors.WaitCursor
        REPRINT_CONFIRMED = False

        EnforceConstraints(False)

        Fill_Records("SOTPICKX")

        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("")
            Dim CUST_CODE As String = rowSOTPICKX.Item("CUST_CODE")
            rowSOTPICKX.Item("OPT_PICK_TICKET") = "1"

            If Val(rowSOTPICKX.Item("ORDR_CNT_PICK") & "") > 1 Then
                '  rowSOTPICKX.Item("OPT_PULL_STORE") = "1"
                rowSOTPICKX.Item("OPT_PULL_STYLE") = "1"
            Else
                '  rowSOTPICKX.Item("OPT_PULL_STORE") = "0"
                rowSOTPICKX.Item("OPT_PULL_STYLE") = "0"
            End If

            If Val(rowSOTPICKX.Item("ORDR_CNT_CART") & "") > 5 Then ' MAYBE SHOULD BE PARAMETERIZED
                rowSOTPICKX.Item("OPT_MANIFEST") = "1"
            Else
                rowSOTPICKX.Item("OPT_MANIFEST") = "0"
            End If

            Dim rowSOTUCCL1 As DataRow = dst.Tables("SOTUCCL1").Rows.Find(CUST_CODE)
            If rowSOTUCCL1 IsNot Nothing AndAlso rowSOTUCCL1.Item("UCC128_PREPRINT") & "" = "1" Then
                rowSOTPICKX.Item("OPT_UCC128") = "1"
            Else
                rowSOTPICKX.Item("OPT_UCC128") = "0"
            End If

            If CUST_CODEs_856.Contains(CUST_CODE) Then
                rowSOTPICKX.Item("CUST_856") = "1"
            Else
                rowSOTPICKX.Item("CUST_856") = "0"
            End If
        Next

        grdSOTPICKX.Text = "Shipments for " & WHSE_CODE

        Display_Totals()
        Sort_grdColumns(grdSOTPICKX, "")

        EnforceConstraints(True)

        grdSOTPICKX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Display_Totals()

    End Sub

    Sub Setup_Print_Option()
        UltraExplorerBar1.Groups("Screen Control").Items("Re-Print Confirmed").Settings.Enabled = (optBOL.Value = "1") And ScreenMode
    End Sub

    Sub Show_Pick_Tickets()
        If grdSOTPICKX.ActiveRow Is Nothing OrElse Not grdSOTPICKX.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        splSOTPICKX.Panel2Collapsed = False

        Dim CUST_CODE As String = grdSOTPICKX.ActiveRow.Cells("CUST_CODE").Value
        Dim ORDR_GROUP_NO As String = grdSOTPICKX.ActiveRow.Cells("ORDR_GROUP_NO").Value
        Dim PICK_BATCH_NO As String = grdSOTPICKX.ActiveRow.Cells("PICK_BATCH_NO").Value
        Dim SHIP_ADDR_TYPE As String = grdSOTPICKX.ActiveRow.Cells("SHIP_ADDR_TYPE").Value
        Dim SHIP_ADDR_CODE As String = grdSOTPICKX.ActiveRow.Cells("SHIP_ADDR_CODE").Value & ""

        ASCMAIN1.sql = "Select SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            & " from SOTPICK1,SOTORDR1,SOTSHIP1" & vbCrLf _
            & " where SOTORDR1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_BATCH_NO = '" & PICK_BATCH_NO & "'" & vbCrLf _
            & "   and SOTSHIP1.SHIP_ADDR_TYPE = '" & SHIP_ADDR_TYPE & "'" & vbCrLf
        If SHIP_ADDR_TYPE = "DC" Then
            ASCMAIN1.sql &= "   and SOTSHIP1.SHIP_ADDR_CODE = '" & SHIP_ADDR_CODE & "'"
        End If

        Dim dt As DataTable = ASCDATA1.GetDataTable
        Stop ' need to dynamially generate a view
        ASCMAIN1.CodeSelector.Get_SQL("")
        ASCMAIN1.CodeSelector.UseDataFromTable = dt
        If ASCMAIN1.CodeSelector.Selections <> 0 Then

            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("")
                With rowSOTPICKX
                    If .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO And _
                       .Item("PICK_BATCH_NO") = PICK_BATCH_NO And _
                       .Item("SHIP_ADDR_TYPE") = SHIP_ADDR_TYPE Then
                        .Item("SELECTED") = "1"
                    Else
                        .Item("SELECTED") = "0"
                    End If
                End With
            Next

            Dim PICK_NOs As String = ""
            For Each PICK_NO As String In ASCMAIN1.CodeSelector.SelectedCodes
                PICK_NOs &= ",'" & PICK_NO & "'"
            Next
            Print_Pick_Tickets(Mid(PICK_NOs, 2))
            Print_Report("SORPICK1")
        End If
    End Sub

    Function Get_Selected_ORDR_GROUP_NOs() As String
        Dim ORDR_GROUP_NO As String = ""
        Dim ORDR_GROUP_NOs As String = ""

        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SELECTED = '1'")
            ORDR_GROUP_NO = rowSOTPICKX.Item("ORDR_GROUP_NO")
            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then
                ORDR_GROUP_NOs &= ",'" & ORDR_GROUP_NO & "'"
            End If
        Next
        ORDR_GROUP_NOs = Mid$(ORDR_GROUP_NOs, 2)
        Return ORDR_GROUP_NOs
    End Function

    Private Sub optBOL_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optBOL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Print_Option()
        Refresh_SOTPICKX("")
    End Sub

    Private Sub grdSOTPICKX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICKX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("SELECTED").Value & "" = "1" Then
                e.Row.Cells("SELECTED").Appearance.BackColor = Drawing.Color.Green
            Else
                e.Row.Cells("SELECTED").Appearance.BackColor = Drawing.Color.Empty
            End If
            If e.Row.Cells("ORDR_HIGH_PRIORITY").Value & "" = "1" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Sub Refresh_ICTWHSEX()
        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
    End Sub

    Sub Setup_SOTSHIP1()
        If grdSOTPICKX.ActiveRow Is Nothing OrElse Not grdSOTPICKX.ActiveRow.IsDataRow Then
            splSOTPICK1.Visible = False
        Else
            splSOTPICK1.Visible = True
            Dim SHIP_BOL_NO As String = grdSOTPICKX.ActiveRow.Cells("SHIP_BOL_NO").Value

            Load_SOTPICK1(SHIP_BOL_NO, True)

            Sort_grdColumns(grdSOTPICK1, "PICK_NO")

            chkEdit.Checked = False
            Setup_SHIPTO_Edit()

        End If
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Setup_SOTPICK1()
    End Sub

    Sub Setup_SOTPICK1()
        If printingPickTickets Then Exit Sub

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTCART1.Visible = False
        Else
            grdSOTCART1.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value

            Load_SOTCART1("", PICK_NO)

            Sort_grdColumns(grdSOTCART1, "CART_NO")
        End If

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTPICKL.Visible = False
        Else
            grdSOTPICKL.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            Fill_Records("SOTPICKL", PICK_NO)
            Sort_grdColumns(grdSOTPICKL, "PICK_LNO")
        End If

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            tabLabels.Visible = False
        Else
            tabLabels.Visible = True
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            ASCMAIN1.sql = "Select SOTORDR5.*, SOTORDR1.WHSE_CODE" & vbCrLf _
                & " from SOTORDR5,SOTORDR1" & vbCrLf _
                & " where SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE = 'ST'"
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

            chkEdit.Checked = False
        End If
    End Sub

    Sub Load_SOTPICK1(SHIP_BOL_NO As String, load_cartons As Boolean)

        EnforceConstraints(False)

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & ", SOTORDR0.CUST_CODE" & vbCrLf _
            & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" & vbCrLf _
            & " from SOTSHIP1,SOTORDR0" & vbCrLf _
            & " where SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
            & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
        Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CUST_STORE_NO from SOTPICK1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & " from SOTPICK1,SOTPICK2,SOTORDR2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

        dst.Tables("SOTCART1").Rows.Clear()
        dst.Tables("SOTCART2").Rows.Clear()

        ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO" & vbCrLf _
           & " from SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
           & " from SOTORDR2,SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

        EnforceConstraints(True)

        If load_cartons Then
            Load_SOTCART1(SHIP_BOL_NO, "")
        End If

    End Sub

    Sub Load_SOTCART1(SHIP_BOL_NO As String, PICK_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipment Data")

        EnforceConstraints(False)

        ASCMAIN1.sql = sqlSOTCART1 & " from SOTCART1,SOTPICK1" _
            & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
            & IIf(SHIP_BOL_NO <> "", _
                  "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", _
                  "   and SOTCART1.PICK_NO = '" & PICK_NO & "'")
        Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
            & " from SOTCART2,SOTCART1,SOTPICK1" & vbCrLf _
            & " where SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
            & IIf(SHIP_BOL_NO <> "", _
                  "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", _
                  "   and SOTCART1.PICK_NO = '" & PICK_NO & "'")
        Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)
        SetCartItemCode()

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_UCC128_Labels()
        Try
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTCART1.Selected.Rows
                Dim CART_NO As String = grow.Cells("CART_NO").Value
                Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                cartonLabel.PrintLabel()
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Print_UCC128_Labels_for_Selected_Shipments(ByVal RecordSelection As String)
        Try
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select(RecordSelection, "SHIP_BOL_NO,PICK_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                ASCMAIN1.sql = "Select CART_NO from SOTCART1 where PICK_NO = '" & PICK_NO & "'"
                For Each rowCART_NO As DataRow In ASCDATA1.GetDataTable.Select("", "CART_NO")
                    Dim CART_NO As String = rowCART_NO.Item("CART_NO")
                    Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                    cartonLabel.PrintLabel()
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Print_Address_Labels_from_Details(LABEL_TEMPLATE As String)
        Dim TOTAL_LABELS As Integer = Val(dst.Tables("SOTPICKL").Compute("SUM(LABEL_QTY)", "LABEL_QTY > 0") & "")
        Dim STARTING_LABEL As Integer = 1
        For Each rowSOTPICKL As DataRow In dst.Tables("SOTPICKL").Select("LABEL_QTY > 0")
            Dim labelQty As Int32 = Val(rowSOTPICKL.Item("LABEL_QTY") & "")
            Dim labelComment As String = rowSOTPICKL.Item("ITEM_CODE") & " - " & rowSOTPICKL.Item("ITEM_DESC")
            Print_Address_Labels(labelQty, labelComment, LABEL_TEMPLATE, TOTAL_LABELS, STARTING_LABEL)
            STARTING_LABEL += labelQty
        Next
    End Sub

    Sub Print_Address_Labels(labelQty As Integer, labelComment As String, LABEL_TEMPLATE As String, Optional TOTAL_LABELS As Integer = 0, Optional STARTING_LABEL As Integer = 0)
        Dim pickNo As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)

        Try
            Dim addressLabel As New AddressLabel(pickNo, labelComment, LABEL_TEMPLATE, rowSOTORDR5)
            addressLabel.Set1of9(STARTING_LABEL, TOTAL_LABELS)
            addressLabel.PrintLabel(labelQty)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private Sub SetUpPortsAndPrinters()

        Dim tooltip As New System.Windows.Forms.ToolTip()

        ' Label Printer Port
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If

            'txtLabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.BackColor = Drawing.Color.Green
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
        End Try

    End Sub

    Function Default_Printer()
        Dim settings As New PrinterSettings
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If settings.IsDefaultPrinter Then
                Return printer
            End If
        Next
        Return String.Empty
    End Function

    Private Sub txtPICK_NO_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtPICK_NO.ValueChanged
        If txtPICK_NO.TextLength = 10 Then
            Dim PICK_NO As String = txtPICK_NO.Text
            Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
            If rowSOTPICK1 IsNot Nothing Then
                Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO")
                Dim row As DataRow = dst.Tables("SOTPICKX").Rows.Find(SHIP_BOL_NO)
                If row IsNot Nothing Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICKX.Rows
                        If grow.IsDataRow AndAlso grow.Cells("SHIP_BOL_NO").Value & "" = SHIP_BOL_NO Then
                            grdSOTPICKX.ActiveRow = grow
                            For Each grow2 As UltraWinGrid.UltraGridRow In grdSOTPICK1.Rows
                                If grow2.IsDataRow AndAlso grow2.Cells("PICK_NO").Value & "" = PICK_NO Then
                                    grdSOTPICK1.ActiveRow = grow2
                                    Exit Sub
                                End If
                            Next
                            Exit Sub
                        End If
                    Next
                End If
            Else
                ' DISPLAY MESSAGE?
            End If
            txtPICK_NO.Text = ""
        End If
    End Sub

    Private Sub tabLabels_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabLabels.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabLabels()
    End Sub

    Sub Setup_tabLabels()
        With UltraExplorerBar1.Groups("Label Printing")
            .Items("Print UCC128 Labels").Visible = (tabLabels.SelectedTab.Key = "UCC128")
            .Items("Print Address Labels").Visible = (tabLabels.SelectedTab.Key = "Address")
            .Items("Print Address/Item Labels").Visible = (tabLabels.SelectedTab.Key = "Content")
            .Items("Print Content Labels").Visible = (tabLabels.SelectedTab.Key = "Content")
        End With
    End Sub

    Private Sub chkEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkEdit.CheckedChanged
        If chkEdit.Checked Then
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO, , , , 1) Then
                chkEdit.Checked = False
                Exit Sub
            Else
                dst.Tables("SOTORDR5").AcceptChanges()
                ORDR_NO_MT = ORDR_NO
            End If
        Else
            If ORDR_NO_MT <> "" Then
                ASCMAIN1.MultiTask_Release(, , 1)
                ORDR_NO_MT = ""
            End If
        End If
        Setup_SHIPTO_Edit()
    End Sub

    Sub Setup_SHIPTO_Edit()
        Set_Read_Only(grpSHIPTO, Not chkEdit.Checked)
        Set_Read_Only_for_ctl(chkEdit, False)
        cmdUpdate.Visible = chkEdit.Checked
        cmdCancel.Visible = chkEdit.Checked
        chkSaveToMasterFile.Checked = False
        chkSaveToMasterFile.Visible = chkEdit.Checked
    End Sub

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click

        BeginTrans()

        dst.Tables("SOTORDXR").Rows.Clear()

        Synch_TABLE_NAME("SOTORDR5")

        ASCMAIN1.sql = "Select Max (REV_NO) from SOTORDXR where ORDR_NO = '" & ORDR_NO_MT & "'"
        Dim REV_NO As Integer = Val(ASCDATA1.GetDataValue & "") + 1

        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO_MT & "' and CUST_ADDR_TYPE = 'ST'")(0)
        TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO_MT, DATETIME_STAMP, ASCMAIN1.USER_ID, "ADRCHG", "Ship-To Address Changed")
        TAC.SOCMAIN1.Log_Changes(Me, ORDR_NO_MT, rowSOTORDR5, "SOTORDR5", False, REV_NO, 0, Now)

        Update_Record_TDA("SOTORDR5")
        Update_Record_TDA("SOTORDXR")

        If chkSaveToMasterFile.Checked Then

            Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO_MT)
            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")
            dst.Tables("ARTCUST2").Rows.Clear()
            Dim rowARTCUST2 As DataRow = Fill_Record("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

            Dim COLUMN_NAMEs() As String = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_CONTACT"}

            For Each COLUMN_NAME As String In COLUMN_NAMEs
                rowARTCUST2.Item(Replace(COLUMN_NAME, "CUST_", "CUST_STORE_")) = rowSOTORDR5.Item(COLUMN_NAME)
            Next

            Update_Record_TDA("ARTCUST2")
        End If

        CommitTrans()

        chkEdit.Checked = False
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        chkEdit.Checked = False
    End Sub

    Private Sub SetCartItemCode()

        Dim CART_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim rowSOTCART2 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing

        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            CART_NO = rowSOTCART1.Item("CART_NO")

            Dim numItems As Int16 = ASCDATA1.SelectDistinct(dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'"), "ITEM_CODE").Rows.Count
            If numItems = 1 Then
                rowSOTCART2 = dst.Tables("SOTCART2").Select("CART_NO = '" & CART_NO & "'")(0)
                ORDR_NO = rowSOTCART2.Item("ORDR_NO") & String.Empty
                ORDR_LNO = Val(rowSOTCART2.Item("ORDR_LNO") & String.Empty)

                rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                If rowSOTORDR2 IsNot Nothing Then
                    rowSOTCART1.Item("ITEM_CODE") = rowSOTORDR2.Item("ITEM_CODE")
                    rowSOTCART1.Item("ITEM_DESC") = rowSOTORDR2.Item("ITEM_DESC")
                Else
                    rowSOTCART1.Item("ITEM_CODE") = "Could not resolve item"
                End If
            Else
                rowSOTCART1.Item("ITEM_CODE") = "Mixed Items"
            End If
        Next

    End Sub

    ''' <summary>
    ''' De-Releases one or more Pick Tickets.
    ''' </summary>
    ''' <param name="PickTicketList"></param>
    ''' <remarks></remarks>
    Sub DeRelease_PickTickets(ByVal PickTicketList As List(Of String))

        Try

            Static tblSOTPICK1 As String = String.Empty

            If PickTicketList Is Nothing OrElse PickTicketList.Count = 0 Then
                MessageBox.Show("No Pick Tickets selected to De-Release", "De-Release", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim pick_nos As String = String.Empty
            For Each pickNo As String In PickTicketList
                pickNo = pickNo.TrimStart
                If pickNo.Length > 0 Then
                    pick_nos &= ", '" & pickNo & "'"
                End If
            Next

            If pick_nos.Length = 0 Then
                MessageBox.Show("No Pick Tickets selected to De-Release", "De-Release", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Else
                pick_nos = pick_nos.Substring(1).Trim
            End If

            Dim sql_pick As String = "SELECT PICK_NO, ORDR_NO, SHIP_BOL_NO FROM SOTPICK1 WHERE PICK_NO IN (" & pick_nos & ") and (SOTPICK1.PICK_STATUS = 'P' OR SOTPICK1.PICK_STATUS = 'C')"
            If tblSOTPICK1.Length > 0 Then
                ASCDATA1.ExecuteSQL("Truncate Table " & tblSOTPICK1)
                ASCDATA1.ExecuteSQL("Insert Into " & tblSOTPICK1 & " " & sql_pick)
            Else
                tblSOTPICK1 = ASCMAIN1.Temp_Table(sql_pick)
            End If

            Dim numTickets As Int16 = ASCDATA1.GetDataValue("select count(*) from " & tblSOTPICK1)
            If numTickets = 0 Then
                MessageBox.Show("No Pick Tickets selected to De-Release", "De-Release", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If MessageBox.Show("Do you want to De-Release " & numTickets & " Pick Ticket(s)?", "De-Release", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            chkDeRelease.Checked = False

            BeginTrans()
            ASCMAIN1.Progress("Now De-Releasing Pick Tickets", "")

            ASCMAIN1.Progress("-", "Items")
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select SOTORDR1.WHSE_CODE" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", SUM (NVL(SOTPICK2.PICK_QTY,0)) QTY" & vbCrLf _
                & ", SUM (NVL(SOTPICK2.PICK_QTY_CANC_REL,0)) QTY_CANC" & vbCrLf _
                & ", SUM (NVL(SOTPICK2.PICK_QTY_BACK_REL,0)) QTY_BACK" & vbCrLf _
                & " from SOTORDR2,SOTPICK2,SOTPICK1,SOTORDR1, " & tblSOTPICK1 & " SOTPICKX" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICKX.PICK_NO" & vbCrLf _
                & " group by SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE;" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & " Update ICTSTAT2 " & vbCrLf _
                & " Set WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - R1.QTY, " & vbCrLf _
                & "     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) + R1.QTY + R1.QTY_CANC" & vbCrLf _
                & " where ITEM_CODE = R1.ITEM_CODE" & vbCrLf _
                & "   and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & " If SQL%NOTFOUND Then" & vbCrLf _
                & "   Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_OPEN)" & vbCrLf _
                & "   Values (R1.ITEM_CODE, R1.WHSE_CODE, -1 * R1.QTY, R1.QTY + R1.QTY_CANC);" & vbCrLf _
                & " End If;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            'ASCMAIN1.Progress("-", "Cartons")
            'ASCMAIN1.sql = "Delete FROM SOTCART2 where CART_NO in (Select CART_NO from SOTCART1 where PICK_NO in (SELECT PICK_NO from " & tblSOTPICK1 & "))"
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Delete from SOTCART1 where PICK_NO in (SELECT PICK_NO from " & tblSOTPICK1 & ")"
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.Progress("-", "Status")
            ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'O', ORDR_HOLD = '1', ORDR_HOLD_REASON = 'DE-RELEASED' where ORDR_NO in (Select ORDR_NO from " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.Progress("-", "Orders")
            ASCMAIN1.sql = "Update SOTORDR1 Set " _
                & "  ORDR_DATE_CLOSED = Null, ORDR_YYYYPP_CLOSED = Null, ORDR_YYYYPP_UPDATED = Null" & vbCrLf _
                & ", REORD_MEMO_IND = Null, ORDR_DATE_REL = Null, ORDR_REL_BATCH_NO = Null, ORDR_BATCHED = Null" & vbCrLf _
                & ", ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'Q'" & vbCrLf _
                & " where ORDR_NO in (Select ORDR_NO from " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC,EVENT_KEY) " & vbCrLf _
                & " Select 'SOTORDR1', ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'DREL', 'Pick Ticket De-Released', PICK_NO" & vbCrLf _
                & " from SOTPICK1 Where PICK_NO IN (Select PICK_NO from " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'O' where ORDR_NO in (Select ORDR_NO from " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_RELEASE = Null where ORDR_NO in (Select ORDR_NO from " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.Progress("-", "Tickets")
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select SOTPICK2.* from SOTPICK1,SOTPICK2, " & tblSOTPICK1 & " SOTPICKX" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & " and SOTPICK1.PICK_NO = SOTPICKX.PICK_NO;" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & " Update SOTORDR2 " & vbCrLf _
                & " Set ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - NVL(R1.PICK_QTY,0)," & vbCrLf _
                & "     ORDR_QTY_OPEN = NVL(ORDR_QTY_OPEN,0) + NVL(R1.PICK_QTY,0) + NVL(R1.PICK_QTY_CANC_REL,0)," & vbCrLf _
                & "     ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = R1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is" & vbCrLf _
                & " Select DISTINCT ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
                & "     where SHIP_BOL_NO in (Select DISTINCT SHIP_BOL_NO from " & tblSOTPICK1 & ");" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & "     SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTPICK1 set " & vbCrLf _
                & " PICK_STATUS = 'D', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                & " Where PICK_NO IN (SELECT PICK_NO FROM " & tblSOTPICK1 & ")"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.Progress("-", "Shipments")
            ASCMAIN1.sql = "Select SHIP_BOL_NO" & vbCrLf _
                & ", Sum (Decode (PICK_STATUS,'P',1,0)) PICK" & vbCrLf _
                & ", Sum (Decode (PICK_STATUS,'F',1,0)) SHIP" & vbCrLf _
                & ", Count (*) TOTAL" & vbCrLf _
                & " from SOTPICK1 " & vbCrLf _
                & " where SHIP_BOL_NO in " & vbCrLf _
                & " (Select SHIP_BOL_NO from " & tblSOTPICK1 & " WHERE SHIP_BOL_NO IS NOT NULL)" & vbCrLf _
                & " group by SHIP_BOL_NO"

            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)

                If Val(row.Item("PICK") & "") = 0 Then
                    Dim SHIP_STATUS As String = ""
                    If Val(row.Item("SHIP") & "") = 0 Then
                        SHIP_STATUS = "D"
                    Else
                        SHIP_STATUS = "F" ' SHOULDNT SET F WITHOUT OTHER FIELDS WHICH GET THEIR VALUE VIA DATA ENTRY IN SHIPMENTS CONF
                        Stop ' MUST RESEARCH HOW THIS IS POSSIBLE, IF IT EVER HAPPENS
                        SHIP_STATUS = ""
                    End If
                    If SHIP_STATUS <> "" Then
                        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_STATUS = '" & SHIP_STATUS & "'" _
                            & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                            & ", LP_STATUS = NULL" _
                            & " where SHIP_BOL_NO = '" & row.Item("SHIP_BOL_NO") & "'"
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            Next

            CommitTrans("De-Release successful")
            Click_Command("Refresh")

            Try
                Fill_Records("SOTDREL1", String.Empty, True, sqlDerelease & " and SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & tblSOTPICK1 & ")")
                Dim RPT_TITLE As String = "De-Released Orders"
                Dim reportFile As String = "SORDREL2"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Preparing " & RPT_TITLE)

                Print_Report_Begin()
                CR_params.Add("SUBT", "")

                Generate_Report(reportFile, RPT_TITLE)
                Print_Report_End()
            Catch ex As Exception
                MessageBox.Show("The following error occured when displaying the report: " & ex.Message, "Print Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try


        Catch ex As Exception
            Rollback("De-Release Aborted for the following reason: " & ex.Message)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub


End Class

#Region "Tests Printing Raw"

Public Class Tests
    Private Sub UltraButton2_Click(sender As System.Object, e As System.EventArgs)

        'Dim fn As String = "c:\vs\aha\temp\zpl.txt"
        Dim fn As String = "C:\Users\wjz\Desktop\Zebra\dist\zpl_sample.txt"
        Dim zpl As String = My.Computer.FileSystem.ReadAllText(fn)
        Stop
        Dim p As New myPrinter
        p.prt(zpl)

        'zpl = "^XA" & _
        '      "^FO50,50" & _
        '      "^A0N,50,50" & _
        '      "^FDHello, World!^FS" & _
        '      "^XZ"

        'Dim p2 As New myPrinter
        'p2.prt(zpl)
    End Sub

    Private Sub UltraButton3_Click(sender As System.Object, e As System.EventArgs)

        '"Zebra  ZP Series-200 dpi"

        Try
            Dim s As New StringBuilder
            s.AppendLine("^XA")
            s.AppendLine("^PR30")
            s.AppendLine("^FO0,50")
            s.AppendLine("^BCN,60,N,N,N,N")
            s.AppendLine("^FDtext")
            s.AppendLine("^FS")
            s.AppendLine("^FO0,150")
            s.AppendLine("^ADN,36,20")
            s.AppendLine("^FDtext")
            s.AppendLine("^FS")
            s.AppendLine("^XZ")

            Dim fn As String = "c:\vs\aha\temp\zpl.txt"
            'Dim fn As String = "C:\Users\wjz\Desktop\Zebra\dist\zpl_sample.txt"
            Dim zpl As String = My.Computer.FileSystem.ReadAllText(fn)
            s = New StringBuilder
            s.AppendLine(zpl)

            For Each F As String In My.Computer.FileSystem.GetFiles("C:\VS\AHA\UCC128")

                If F.Contains("VONMAUR") Then
                    Dim x As String = My.Computer.FileSystem.ReadAllText(F)
                    s = New StringBuilder
                    s.AppendLine(x)

                    Dim p As New PrintRaw
                    '  p.Print(s, txtPrinter.Text)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

End Class

Public Class myPrinter
    Friend TextToBePrinted As String
    Public Sub prt(ByVal text As String)
        TextToBePrinted = text
        Dim prn As New Printing.PrintDocument
        Using (prn)
            prn.PrinterSettings.PrinterName _
               = "Zebra  ZP Series-200 dpi"
            AddHandler prn.PrintPage, _
               AddressOf Me.PrintPageHandler
            prn.Print()
            RemoveHandler prn.PrintPage, _
               AddressOf Me.PrintPageHandler
        End Using
    End Sub

    Private Sub PrintPageHandler(ByVal sender As Object, _
       ByVal args As Printing.PrintPageEventArgs)
        Dim myFont As New Font("Microsoft San Serif", 10)
        args.Graphics.DrawString(TextToBePrinted, _
           New Font(myFont, FontStyle.Regular), _
           Brushes.Black, 50, 50)
    End Sub
End Class

Public Class PrintRaw
    Public Sub Print(ByVal codes As StringBuilder, ByVal printer As String)
        SendToPrinter("Printing zebras", codes.ToString, printer)
    End Sub

    Private Structure Docinfo
        <MarshalAs(UnmanagedType.LPWStr)> Public DocumentName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public OutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public DataType As String
    End Structure

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=False, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function OpenPrinter(ByVal pPrinterName As String, ByRef phPrinter As IntPtr, ByVal pDefault As Integer) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=False, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Integer, ByRef pDocInfo As Docinfo) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Ansi, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal data As String, ByVal buf As Integer, ByRef pcWritten As Integer) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Long
    End Function

    Private Sub SendToPrinter(ByVal printerJobName As String, ByVal rawStringToSendToThePrinter As String, ByVal printerNameAsDescribedByPrintManager As String)
        Dim handleForTheOpenPrinter = New IntPtr()
        Dim documentInformation = New Docinfo()
        Dim printerBytesWritten = 0
        documentInformation.DocumentName = printerJobName
        documentInformation.DataType = vbNullString
        documentInformation.OutputFile = vbNullString
        OpenPrinter(printerNameAsDescribedByPrintManager, handleForTheOpenPrinter, 0)
        StartDocPrinter(handleForTheOpenPrinter, 1, documentInformation)
        StartPagePrinter(handleForTheOpenPrinter)
        WritePrinter(handleForTheOpenPrinter, rawStringToSendToThePrinter, rawStringToSendToThePrinter.Length, printerBytesWritten)
        EndPagePrinter(handleForTheOpenPrinter)
        EndDocPrinter(handleForTheOpenPrinter)
        ClosePrinter(handleForTheOpenPrinter)
    End Sub

End Class
#End Region