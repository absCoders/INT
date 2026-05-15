Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.Forms
Imports System.IO

Public Class RSFLGIDL

#Region "Class Variables"

    Private SO_PARM_MDB_PROVIDER As String = String.Empty
    Private SO_PARM_FTP_IP_ADDRESS As String = String.Empty
    Private SO_PARM_FTP_USER As String = String.Empty
    Private SO_PARM_FTP_PASSWORD As String = String.Empty
    Private SO_PARM_DIR_INBOUND As String = String.Empty
    Private SO_PARM_DIR_OUTBOUND As String = String.Empty
    Private SO_PARM_DIR_ARCHIVE As String = String.Empty
    Private SO_PARM_DOWNLOAD_FILE_NAME As String = String.Empty
    Private SO_PARM_DOWNLOAD_FILE_EXT As String = String.Empty
    Private SO_PARM_LOCAL_DIR_INBOUND As String = String.Empty
    Private SO_PARM_LOCAL_DIR_OUTBOUND As String = String.Empty
    Private SO_PARM_LOCAL_DIR_ARCHIVE As String = String.Empty
    Private SO_PARM_FTP_HOST_NAME As String = String.Empty

    Private remoteDirectoryFileList As List(Of String)
    Private Const oracleLgiSchemaName As String = "JHX"
    Private LgiTableNames() = New String() {"DOORFILE", "INVFILE", "SALESFILE"}

    Private viewLostRetails As DataView = New DataView
    'Private viewItems As DataView = New DataView
    Private viewRetails As DataView = New DataView
    Private viewInventory As DataView = New DataView
    Private rowICTITEM1 As DataRow = Nothing

    Private Const CustCodeNotMapped As String = "** Unk **"
    Private Const ItemCodeNotMapped As String = "<Blank>"
    Private Const ediTransTypeQtySold As String = "QS"
    Private Const ediTransTypeQtyReturned As String = "QU"
    Private Const ediTransTypeQtyInvAvaliable As String = "QA"

    Private OriginalSKU As String = String.Empty
    Private OriginalUPC As String = String.Empty
    Private OriginalItemCode As String = String.Empty
    Private NewItemCode As String = String.Empty
    Private NewItemDesc As String = String.Empty
    Private NewItemUpcCode As String = String.Empty
    Private IGNORE_ITEM_CODE As String = String.Empty

    Dim CNTL_NO As String

    Dim RSTRETL1 As String
    Dim YW(0) As String
    Dim YP(0) As String

    Private Enum FormStates
        LoadDatabases
        LoadTableData
        UpdateToOracle
    End Enum

    Private currentState As FormStates = FormStates.LoadDatabases

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sql As String = String.Empty

        With dst
            Create_TDA(.Tables.Add, "SOTPARMF", "*")

            With .Tables.Add("RSTRETLW")
                .Columns.Add("CUST_CODE", GetType(System.String))
                .Columns.Add("OPS_YYYYWW", GetType(System.String))
                .Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("OPS_YYYYWW")}
            End With

            Create_TDA(.Tables.Add, "RSTLGIED", "*")
            .Tables("RSTLGIED").Columns.Add("FILE_DATE", GetType(System.DateTime))
            .Tables("RSTLGIED").Columns.Add("IMPORT", GetType(System.String))

            Create_TDA(.Tables.Add, "RSTCUST1", "*")
            .Tables("RSTCUST1").Columns.Add("CUST_NAME", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_ADDR1", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_ADDR2", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_CITY", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_STATE", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_ZIP_CODE", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("CUST_PHONE", GetType(System.String))
            .Tables("RSTCUST1").Columns.Add("NEEDS_MAPPING", GetType(System.String))
            .Tables("RSTCUST1").Columns("NEEDS_MAPPING").DefaultValue = "0"

            Create_TDA(.Tables.Add, "RSTITEM1", "*")

            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            MyBase.Fill_Records("ARTCUST1", "", True, "Select * From ARTCUST1 Order by Cust_Code")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            Create_TDA(.Tables.Add, "LGI_DOORFILE", "*")

            Create_TDA(.Tables.Add, "LGI_INVFILE", "*")
            .Tables("LGI_INVFILE").PrimaryKey = New DataColumn() {.Tables("LGI_INVFILE").Columns("INVENTORYID")}

            Create_TDA(.Tables.Add, "LGI_SALESFILE", "*")
            .Tables("LGI_SALESFILE").PrimaryKey = New DataColumn() {.Tables("LGI_SALESFILE").Columns("SALEID")}

            Create_TDA(.Tables.Add, "EDT852T1", "*")
            .Tables("EDT852T1").Columns.Add("RETAILERID", GetType(System.String))
            .Tables("EDT852T1").Columns.Add("DOORID", GetType(System.String))

            ASCMAIN1.sql = "Select * from RSTRETL1 where ROWNUM < 1"
            RSTRETL1 = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & RSTRETL1 & " Add Primary Key (EDI_DOC_SEQ_NO,CUST_CODE,CUST_STORE_NO,ITEM_CODE)")
            ASCMAIN1.sql = "Select * from " & RSTRETL1
            Create_TDA(.Tables.Add, RSTRETL1, "*")

            'Create_TDA(.Tables.Add, "EDT852T2", "*")
            'Create_TDA(.Tables.Add, "EDT852T3", "*")

            Create_Lookup("ICTITEM1", "*", "ITEM_CODE = :PARM1 OR ITEM_UPC_CODE = :PARM1", "V", False)

            sql = "Select * From LGI_INVFILE"
            Create_TDA(.Tables.Add, "EDTRETLS", sql, 0, False, String.Empty, 0, String.Empty)
            .Tables("EDTRETLS").Columns.Add("CUST_CODE", GetType(System.String))
            .Tables("EDTRETLS").Columns("CUST_CODE").MaxLength = .Tables("RSTCUST1").Columns("CUST_CODE").MaxLength
            .Tables("EDTRETLS").Columns.Add("CUST_STORE_NO", GetType(System.String))
            '.Tables("EDTRETLS").Columns.Add("TRANS_TYPE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("ITEM_CODE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("ITEM_DESC", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("OPS_YYYYWW", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("OPS_YYYYPP", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("EDI_FROM_DATE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("EDI_TO_DATE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("EDI_TRAN_TYPE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("EDI_ZA_TRAN_TYPE", GetType(System.String))
            .Tables("EDTRETLS").Columns.Add("EXTENSION", GetType(System.Double), "UNITS * MSRP")
            .Tables("EDTRETLS").Columns.Add("IMPORT", GetType(System.String))
            .Tables("EDTRETLS").Columns("IMPORT").DefaultValue = "0"
            .Tables("EDTRETLS").Columns.Add("IGNORE", GetType(System.String))
            .Tables("EDTRETLS").Columns("IGNORE").DefaultValue = "0"

            .Tables("EDTRETLS").Columns.Add("IGNORE_ITEM_CODE", GetType(System.String))
            .Tables("EDTRETLS").Columns("IGNORE_ITEM_CODE").DefaultValue = "0"

            Dim rstCols As DataColumn() = Nothing
            Dim edtCols As DataColumn() = Nothing

            rstCols = New DataColumn() {.Tables("RSTLGIED").Columns("RS_CNTL_NO"), .Tables("RSTLGIED").Columns("IMPORT")}
            edtCols = New DataColumn() {.Tables("EDTRETLS").Columns("CNTL_NO"), .Tables("EDTRETLS").Columns("IMPORT")}
            .Relations.Add("RSTLGIED_EDTRETLS", rstCols, edtCols)

            rstCols = New DataColumn() {.Tables("RSTCUST1").Columns("RETAILERID"), .Tables("RSTCUST1").Columns("CUST_CODE"), .Tables("RSTCUST1").Columns("IGNORE")}
            edtCols = New DataColumn() {.Tables("EDTRETLS").Columns("RETAILERID"), .Tables("EDTRETLS").Columns("CUST_CODE"), .Tables("EDTRETLS").Columns("IGNORE")}
            .Relations.Add("RSTCUST1_EDTRETLS", rstCols, edtCols)

            With .Tables.Add("RSTLGIDX")
                .Columns.Add("DATA_DATE", GetType(System.DateTime))
                .Columns.Add("WEEK_END_DATE", GetType(System.DateTime))
                .Columns.Add("YYYYWW")
                .Columns.Add("YYYYPP")
                .PrimaryKey = New DataColumn() {.Columns("DATA_DATE")}
            End With

            'ASCMAIN1.sql = "Select * from GLTPARM3"
            'Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False)
        End With

        'Fill_Records("GLTPARM3")

        viewLostRetails.Table = dst.Tables("EDTRETLS")
        viewLostRetails.RowFilter = "(CUST_CODE = '" & CustCodeNotMapped & "' AND IMPORT = '1') OR ITEM_CODE = '' OR IGNORE_ITEM_CODE = '1'"
        viewLostRetails.Sort = "RETAILERID"
        grdLostRetails.DataSource = viewLostRetails

        'viewItems.Table = dst.Tables("EDTRETLS")
        'viewItems.RowFilter = "ITEM_CODE = ''"
        'viewItems.Sort = "UPC"
        'grdItems.DataSource = viewItems

        viewRetails.Table = dst.Tables("EDTRETLS")
        sql = "ITEM_CODE <> '' AND CUST_CODE <> '" & CustCodeNotMapped & "' AND IMPORT = '1' AND IGNORE = '0' AND IGNORE_ITEM_CODE = '0'"
        sql &= " And (EDI_ZA_TRAN_TYPE = '" & ediTransTypeQtyReturned & "' OR EDI_ZA_TRAN_TYPE = '" & ediTransTypeQtySold & "')"
        viewRetails.RowFilter = sql
        viewRetails.Sort = "CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        grdRetails.DataSource = viewRetails

        viewInventory.Table = dst.Tables("EDTRETLS")
        sql = "ITEM_CODE <> '' AND CUST_CODE <> '" & CustCodeNotMapped & "' AND IMPORT = '1' AND IGNORE = '0' AND IGNORE_ITEM_CODE = '0'"
        sql &= " And EDI_ZA_TRAN_TYPE = '" & ediTransTypeQtyInvAvaliable & "'"
        viewInventory.RowFilter = sql
        viewInventory.Sort = "CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        grdInventory.DataSource = viewInventory

        grdRSTCUST1.DataSource = dst.Tables("RSTCUST1")
        grdRSTITEM1.DataSource = dst.Tables("RSTITEM1")
        grdRSTLGIED.DataSource = dst.Tables("RSTLGIED")

        MyBase.Create_Summary(grdRSTCUST1, "RETAILERID", "Count")
        MyBase.Create_Summary(grdRSTITEM1, "SKU", "Count")

        MyBase.Create_Summary(grdLostRetails, "RETAILERID", "Count")
        MyBase.Create_Summary(grdLostRetails, "UNITS")
        MyBase.Create_Summary(grdLostRetails, "EXTENSION")

        'MyBase.Create_Summary(grdItems, "UPC", "Count")
        'MyBase.Create_Summary(grdItems, "UNITS")
        'MyBase.Create_Summary(grdItems, "EXTENSION")

        MyBase.Create_Summary(grdRetails, "CUST_CODE", "Count")
        MyBase.Create_Summary(grdRetails, "UNITS")
        MyBase.Create_Summary(grdRetails, "EXTENSION")

        MyBase.Create_Summary(grdInventory, "CUST_CODE", "Count")
        MyBase.Create_Summary(grdInventory, "UNITS")
        MyBase.Create_Summary(grdInventory, "EXTENSION")

        grdRSTITEM1.DisplayLayout.Bands(0).Columns("SKU").CellActivation = UltraWinGrid.Activation.ActivateOnly
        grdRSTITEM1.DisplayLayout.Bands(0).Columns("SKU").CellClickAction = UltraWinGrid.CellClickAction.CellSelect

        Me.Clear_Record()

        currentState = FormStates.LoadDatabases
        MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").Text = "LGI"
        Me.Proceed_PreReq("Load Databases")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Databases"
                SO_PARM_LOCAL_DIR_INBOUND = String.Empty
                SO_PARM_MDB_PROVIDER = MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").Text
                SO_PARM_MDB_PROVIDER = SO_PARM_MDB_PROVIDER.Trim
                MyBase.Fill_Records("SOTPARMF", SO_PARM_MDB_PROVIDER)
                If dst.Tables("SOTPARMF").Rows.Count = 0 Then
                    SO_PARM_MDB_PROVIDER = String.Empty
                    EMsg = "Missing or Invalid MDB Provider"
                Else
                    SO_PARM_MDB_PROVIDER = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_MDB_PROVIDER") & String.Empty).ToString.Trim
                    SO_PARM_FTP_IP_ADDRESS = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_IP_ADDRESS") & String.Empty).ToString.Trim
                    SO_PARM_FTP_USER = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_USER") & String.Empty).ToString.Trim
                    SO_PARM_FTP_PASSWORD = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_PASSWORD") & String.Empty).ToString.Trim
                    SO_PARM_DIR_INBOUND = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_DIR_INBOUND") & String.Empty).ToString.Trim
                    SO_PARM_DIR_OUTBOUND = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_DIR_OUTBOUND") & String.Empty).ToString.Trim
                    SO_PARM_DIR_ARCHIVE = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_DIR_ARCHIVE") & String.Empty).ToString.Trim
                    SO_PARM_DOWNLOAD_FILE_NAME = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_DOWNLOAD_FILE_NAME") & String.Empty).ToString.Trim
                    SO_PARM_DOWNLOAD_FILE_EXT = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_DOWNLOAD_FILE_EXT") & String.Empty).ToString.Trim
                    SO_PARM_LOCAL_DIR_INBOUND = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_LOCAL_DIR_INBOUND") & String.Empty).ToString.Trim
                    SO_PARM_LOCAL_DIR_OUTBOUND = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_LOCAL_DIR_OUTBOUND") & String.Empty).ToString.Trim
                    SO_PARM_LOCAL_DIR_ARCHIVE = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_LOCAL_DIR_ARCHIVE") & String.Empty).ToString.Trim
                    SO_PARM_FTP_HOST_NAME = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_HOST_NAME") & String.Empty).ToString.Trim

                    If ASCMAIN1.DBS_SERVER = "" And ASCMAIN1.Running_in_VS Then
                        Dim local_dir As String = "C:\Documents and Settings\wjz\Desktop\LGI"
                        MsgBox("Changing Local Inbound Directory from " & SO_PARM_LOCAL_DIR_INBOUND & vbCrLf & " to " & local_dir & "\INBOUND" & vbCrLf & vbCrLf & "and similarly changing Archive and Outbound folder pointers too", MsgBoxStyle.OkOnly, "Developer Notification")
                        SO_PARM_LOCAL_DIR_INBOUND = local_dir & "\INBOUND"
                        SO_PARM_LOCAL_DIR_OUTBOUND = local_dir & "\OUTBOUND"
                        SO_PARM_LOCAL_DIR_ARCHIVE = local_dir & "\ARCHIVE"
                    End If

                    MyBase.Absx1.txtFor("SO_PARM_FTP_HOST_NAME").Text = dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_HOST_NAME") & String.Empty
                    If SO_PARM_LOCAL_DIR_INBOUND.Length = 0 Then
                        EMsg = "MDB Provider missing inbound directory."
                    ElseIf SO_PARM_FTP_USER.Trim.Length = 0 Then
                        EMsg = "MDB Provider missing FTP User Id."
                    ElseIf SO_PARM_FTP_PASSWORD.Trim.Length = 0 Then
                        EMsg = "MDB Provider missing FTP Password."
                    ElseIf Not My.Computer.FileSystem.DirectoryExists(SO_PARM_LOCAL_DIR_INBOUND) Then
                        EMsg = "MDB Provider inbound directory (" & SO_PARM_LOCAL_DIR_INBOUND & ") cannot be found."
                    ElseIf Not My.Computer.FileSystem.DirectoryExists(SO_PARM_LOCAL_DIR_ARCHIVE) Then
                        EMsg = "MDB Provider archive directory (" & SO_PARM_LOCAL_DIR_ARCHIVE & ") cannot be found."
                    End If

                    If Not Me.SO_PARM_LOCAL_DIR_ARCHIVE.EndsWith("\") Then
                        Me.SO_PARM_LOCAL_DIR_ARCHIVE &= "\"
                    End If

                    If Not Me.SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
                        Me.SO_PARM_LOCAL_DIR_INBOUND &= "\"
                    End If

                    If Not Me.SO_PARM_LOCAL_DIR_OUTBOUND.EndsWith("\") Then
                        Me.SO_PARM_LOCAL_DIR_OUTBOUND &= "\"
                    End If

                End If

            Case "Import Data"

            Case "Update"

                If dst.Tables("RSTLGIED").Select("IMPORT = '1'").Length = 0 Then
                    EMsg = "You must select at least one Control No to import."
                ElseIf dst.Tables("EDTRETLS").Select("ITEM_CODE <> '' AND CUST_CODE <> '" & CustCodeNotMapped & "' AND IMPORT = '1'").Length = 0 Then
                    EMsg = "There are no retails to import."
                End If

                Dim rr() As DataRow = dst.Tables("RSTITEM1").Select("IGNORE = '1' AND ISNULL(ITEM_CODE,'') <> ''")
                If rr.Length <> 0 Then
                    EMsg = "There are Items set to Ignore, with a non-blank Item Code indicated" & vbCr & "Ex: " & rr(0).Item("SKU")
                End If

            Case "Cancel"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Databases"
                Me.Load_Record()
                currentState = FormStates.LoadTableData
                Me.Mode_Settings(True)

            Case "Import Data"
                Me.Import_Tables()
                currentState = FormStates.UpdateToOracle
                Me.Mode_Settings(True)

            Case "SKU/Item Matching"
                Item_Matching()

            Case "Update"
                Me.Update_Record()
                currentState = FormStates.LoadDatabases
                Me.Mode_Settings(False)

            Case "Cancel"
                currentState = FormStates.LoadDatabases
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1

                Select Case currentState
                    Case FormStates.LoadDatabases
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Import Data").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("SKU/Item Matching").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.False
                        Me.grdRSTLGIED.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Case FormStates.LoadTableData
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Import Data").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("SKU/Item Matching").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        Me.grdRSTLGIED.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Case FormStates.UpdateToOracle
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Import Data").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("SKU/Item Matching").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        Me.grdRSTLGIED.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                End Select

            End With
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        SplitContainer2.Visible = (currentState = FormStates.UpdateToOracle)

    End Sub

    Private Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("RSTLGIED").Clear()
        dst.Tables("LGI_DOORFILE").Clear()
        dst.Tables("LGI_INVFILE").Clear()
        dst.Tables("LGI_SALESFILE").Clear()
        dst.Tables("RSTCUST1").Clear()
        dst.Tables("RSTITEM1").Clear()
        dst.Tables("ARTCUST2").Clear()
        dst.Tables("EDTRETLS").Clear()
        dst.Tables("RSTRETLW").Clear()

        'dst.Tables("EDT852T0").Clear()
        dst.Tables("EDT852T1").Clear()
        dst.Tables(RSTRETL1).Clear()
        'dst.Tables("EDT852T2").Clear()
        'dst.Tables("EDT852T3").Clear()

        dst.EnforceConstraints = True

        'MyBase.Absx1.txtFor("SO_PARM_LOCAL_DIR_INBOUND").Clear()

    End Sub

    Private Sub Load_Record()
        currentState = FormStates.LoadTableData
    End Sub

    Private Sub Update_Record()

        ASCMAIN1.Progress("Now Updating")

        ASCMAIN1.sql = "Truncate Table " & RSTRETL1
        ASCDATA1.ExecuteSQL()

        MyBase.BeginTrans()

        Try

            Me.UpdateCustomerDoors()
            Me.CreateEdt852Records()

            ASCMAIN1.Progress("Updating Oracle", String.Empty)

            ASCMAIN1.Progress("-", "RSTLGIED")
            Call Update_Record_TDA("RSTLGIED")

            ASCMAIN1.Progress("-", "LGI_DOORFILE")
            Call Update_Record_TDA("LGI_DOORFILE")

            ASCMAIN1.Progress("-", "LGI_INVFILE")
            Call Update_Record_TDA("LGI_INVFILE")

            ASCMAIN1.Progress("-", "LGI_SALESFILE")
            Call Update_Record_TDA("LGI_SALESFILE")

            'ASCDATA1.ExecuteSQL("Update LGI_DOORFILE Set CNTL_NO = '" & CNTL_NO & "' where CNTL_NO = '0000000000'")
            'ASCDATA1.ExecuteSQL("Update LGI_INVFILE Set CNTL_NO = '" & CNTL_NO & "' where CNTL_NO = '0000000000'")
            'ASCDATA1.ExecuteSQL("Update LGI_SALESFILE Set CNTL_NO = '" & CNTL_NO & "' where CNTL_NO = '0000000000'")

            ASCMAIN1.Progress("-", "RSTCUST1")
            Call Update_Record_TDA("RSTCUST1")

            ASCMAIN1.Progress("-", "RSTITEM1")
            Call Update_Record_TDA("RSTITEM1")

            ASCMAIN1.Progress("-", "ARTCUST2")
            Call Update_Record_TDA("ARTCUST2")


            ASCMAIN1.Progress("-", "EDT852T1")
            Call Update_Record_TDA("EDT852T1")

            'Dim T As DataTable = dst.Tables(RSTRETL1).Clone
            'For Each ROW As DataRow In dst.Tables(RSTRETL1).Select("ITEM_CODE = ''")
            '    Dim R As DataRow = T.NewRow
            '    R.ItemArray = ROW.ItemArray
            '    T.Rows.Add(R)
            'Next
            'Stop
            ASCMAIN1.Progress("-", "RSTRETL1")
            Call Update_Record_TDA(RSTRETL1)


            ASCMAIN1.sql = "Insert into RSTRETL1" _
            & " SELECT EDI_DOC_SEQ_NO, CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
            & ", SUM (QTY_SOLD), SUM (AMT_SOLD), OPS_YYYYPP, OPS_YYYYWW" _
            & ", SUM (QTY_EOW) FROM " & RSTRETL1 _
            & " GROUP BY EDI_DOC_SEQ_NO, CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
            & ", OPS_YYYYPP, OPS_YYYYWW"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.Progress("-", "RSTRETLx")
            For Each rowEDT852T1 In dst.Tables("EDT852T1").Rows
                Dim EDI_DOC_SEQ_NO As String = rowEDT852T1.ITEM("EDI_DOC_SEQ_NO")
                TAC.RSCMAIN1.Update_RSTRETLx(EDI_DOC_SEQ_NO, "+")
            Next

            'Dim YWmax As String = YW(UBound(YW))
            'For i As Int16 = UBound(YW) - 1 To 1 Step -1

            '    ASCMAIN1.sql = "Select RSTRETL1.CUST_CODE, SUM (RSTRETL1.QTY_EOW) EOW " _
            '    & " from RSTRETL1,EDT852T1 where EDT852T1.EDI_DOC_SEQ_NO = RSTRETL1.EDI_DOC_SEQ_NO" _
            '    & " and EDT852T1.EDI_CUST_BATCH_NO = 'LGI' and EDT852T1.OPS_YYYYWW = '" & YW(i) & "'" _
            '    & " GROUP BY RSTRETL1.CUST_CODE"
            '    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            '        Dim CUST_CODE As String = row.Item("CUST_CODE")
            '        Dim EOW As Int32 = Val(row.Item("EOW") & "")
            '        If EOW = 0 Then
            '            ASCDATA1.ExecuteSP("RSPLGIO1_CUST_CODE", "VVV", New Object() {YWmax, YW(i), CUST_CODE}, New String() {"YW_FROM", "YW_TO", "CUST_CODE_IN"})
            '        Else
            '            YWmax = YW(i)
            '        End If
            '    Next
            'Next


            Dim YWmax As String = YW(UBound(YW))
            For i As Int16 = UBound(YW) - 1 To 1 Step -1

                ASCMAIN1.sql = "Select SUM (QTY_EOW) EOW " _
                & " from RSTRETL1,EDT852T1 where EDT852T1.EDI_DOC_SEQ_NO = RSTRETL1.EDI_DOC_SEQ_NO" _
                & " and EDT852T1.EDI_CUST_BATCH_NO = 'LGI' and EDT852T1.OPS_YYYYWW = '" & YW(i) & "'"
                Dim EOW As Int32 = Val(ASCDATA1.GetDataValue)
                If EOW = 0 Then
                    ASCDATA1.ExecuteSP("RSPLGIO1", "VV", New Object() {YWmax, YW(i)}, New String() {"YW_FROM", "YW_TO"})
                Else
                    YWmax = YW(i)
                End If
            Next


        Catch ex As Exception
            MyBase.Rollback(ex.Message)
            Exit Sub
        End Try

        MyBase.CommitTrans("Retail Sales Updated")

        Try
            ASCMAIN1.Progress("Archiving Databases", String.Empty)
            For Each rowRSTLGIED As DataRow In dst.Tables("RSTLGIED").Select("IMPORT = '1'")
                ASCMAIN1.Progress("-", String.Empty)
                Dim fileName As String = rowRSTLGIED.Item("FILE_NAME") & String.Empty
                System.IO.File.Move(Me.SO_PARM_LOCAL_DIR_INBOUND & fileName, Me.SO_PARM_LOCAL_DIR_ARCHIVE & fileName)
            Next
        Catch ex As Exception
            MessageBox.Show("Error archiving LGI Retail Sales file: " & ex.Message, "Archive Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Private Procedures"

    ''' <summary>
    ''' Imports LGI Retail Sales Dats
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Import_Tables()

        MyBase.EnforceConstraints(False)

        If Not chkSkipftp.Checked Then
            ' FTP files from LGI to local directory
            Me.DownloadLGIdata()
        End If

        ' Get any downloaded Unprocessed MDBs
        Me.GetUnprocessedLGIDatabases()

        ' Load Retail Mapping and look for any new Retailers
        Me.MapRetailers()

        ' Load SKUs with mappings to other Item Codes
        Me.LoadUnMappedItems()

        ' Map the UPCs to the ABSolution Item Codes
        Me.CreateRetailSalesRecords()

        ' Grab any Retailers in Sales or Inv tables but not in Doorfile
        Me.GetOrphanedCustomers()

        'For Each TABLE_NAME As String In New String() {"LGI_DOORFILE", "LGI_INVFILE", "LGI_SALESFILE", "EDTRETLS"}
        '    For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
        '        row.Item("CNTL_NO") = CNTL_NO ' "0000000000" ' CNTL_NO
        '    Next
        'Next


        MyBase.EnforceConstraints(True)

    End Sub

    ''' <summary>
    ''' Creat EDT852 records
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateEdt852Records()

        ASCMAIN1.Progress("Create 852 Documents", String.Empty)

        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim EDI_LINE_NO As Integer = 1
        Dim DATA_DATE As Date = Nothing ' String = String.Empty
        Dim sql As String = String.Empty

        Dim rowEDT852T1 As DataRow = Nothing
        Dim rowRSTRETL1 As DataRow = Nothing
        'Dim rowEDT852T2 As DataRow = Nothing
        'Dim rowEDT852T3 As DataRow = Nothing

        Dim CUST_CODE As String = String.Empty
        Dim EDI_STORE_NO As String = String.Empty
        Dim EDI_FROM_DATE As Date = Nothing
        Dim EDI_TO_DATE As Date = Nothing
        Dim OPS_YYYYWW As String = ""
        Dim OPS_YYYYPP As String = ""
        Dim CUST_STORE_NO As String = ""
        Dim ITEM_CODE As String = ""

        ReDim YW(0)
        ReDim YP(0)
        For Each row As DataRow In ASCDATA1.SelectDistinct("EDTRETLS", "OPS_YYYYWW", "OPS_YYYYPP").Select("", "OPS_YYYYWW")
            Dim YWi As Int32 = UBound(YW) + 1
            ReDim Preserve YW(YWi)
            YW(YWi) = row.Item("OPS_YYYYWW")
            ReDim Preserve YP(YWi)
            YP(YWi) = row.Item("OPS_YYYYPP")
        Next

        For Each rowEDTRETLS As DataRow In dst.Tables("EDTRETLS") _
        .Select("IMPORT = '1' AND IGNORE <> '1' AND IGNORE_ITEM_CODE <> '1' AND ITEM_CODE <> ''", "CUST_CODE, OPS_YYYYWW")

            If CUST_CODE <> rowEDTRETLS.Item("CUST_CODE") & String.Empty _
            Or OPS_YYYYWW <> rowEDTRETLS.Item("OPS_YYYYWW") & String.Empty Then
                'Or EDI_STORE_NO <> rowEDTRETLS.Item("CUST_STORE_NO") & String.Empty _
                'Or EDI_FROM_DATE <> CDate(rowEDTRETLS.Item("EDI_FROM_DATE") & String.Empty) _
                'Or EDI_TO_DATE <> CDate(rowEDTRETLS.Item("EDI_TO_DATE") & String.Empty) Then

                CUST_CODE = rowEDTRETLS.Item("CUST_CODE") & String.Empty
                OPS_YYYYPP = rowEDTRETLS.Item("OPS_YYYYPP") & String.Empty
                OPS_YYYYWW = rowEDTRETLS.Item("OPS_YYYYWW") & String.Empty
                EDI_STORE_NO = rowEDTRETLS.Item("CUST_STORE_NO") & String.Empty
                EDI_FROM_DATE = rowEDTRETLS.Item("EDI_FROM_DATE") & String.Empty
                EDI_TO_DATE = rowEDTRETLS.Item("EDI_TO_DATE") & String.Empty

                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO")
                EDI_LINE_NO = 1

                Dim rowRSTRETLW As DataRow = dst.Tables("RSTRETLW").NewRow
                rowRSTRETLW.Item("CUST_CODE") = CUST_CODE
                rowRSTRETLW.Item("OPS_YYYYWW") = OPS_YYYYWW
                rowRSTRETLW.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                dst.Tables("RSTRETLW").Rows.Add(rowRSTRETLW)

                ' ***** Fill EDT852T1 *****
                rowEDT852T1 = dst.Tables("EDT852T1").NewRow
                rowEDT852T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                rowEDT852T1.Item("EDI_FROM_DATE") = EDI_FROM_DATE
                rowEDT852T1.Item("EDI_TO_DATE") = EDI_TO_DATE
                rowEDT852T1.Item("EDI_CUST_BATCH_NO") = "LGI"
                rowEDT852T1.Item("EDI_DEPT_NO") = SO_PARM_MDB_PROVIDER
                rowEDT852T1.Item("EDI_STORE_NO") = EDI_STORE_NO
                rowEDT852T1.Item("EDI_STATUS") = "M"
                rowEDT852T1.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowEDT852T1.Item("OPS_YYYYWW") = OPS_YYYYWW
                rowEDT852T1.Item("CUST_CODE") = CUST_CODE
                rowEDT852T1.Item("INIT_DATE") = DATETIME_STAMP
                rowEDT852T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowEDT852T1.Item("LAST_DATE") = DATETIME_STAMP
                rowEDT852T1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowEDT852T1.Item("DATA_LEVEL") = "I"
                dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)

            End If


            ASCMAIN1.Progress("-", Val(EDI_DOC_SEQ_NO) & " / " & Val(EDI_LINE_NO))
            Application.DoEvents()

            ' ***** Fill RSTRETL1 *****
            Dim EDI_ZA_TRAN_TYPE = rowEDTRETLS.Item("EDI_ZA_TRAN_TYPE")
            rowRSTRETL1 = dst.Tables(RSTRETL1).NewRow
            rowRSTRETL1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
            rowRSTRETL1.Item("CUST_CODE") = CUST_CODE
            rowRSTRETL1.Item("CUST_STORE_NO") = EDI_STORE_NO
            rowRSTRETL1.Item("ITEM_CODE") = rowEDTRETLS.Item("ITEM_CODE") & String.Empty
            'If rowRSTRETL1.Item("ITEM_CODE") & "" = "" Then
            '    rowRSTRETL1.Item("ITEM_CODE") = "MDDV0008"
            'End If
            If EDI_ZA_TRAN_TYPE = "QA" Then
                rowRSTRETL1.Item("QTY_SOLD") = 0
                rowRSTRETL1.Item("AMT_SOLD") = 0
                rowRSTRETL1.Item("QTY_EOW") = Val(rowEDTRETLS.Item("UNITS") & String.Empty)
            ElseIf EDI_ZA_TRAN_TYPE = "QS" Or EDI_ZA_TRAN_TYPE = "QU" Then
                rowRSTRETL1.Item("QTY_SOLD") = Val(rowEDTRETLS.Item("UNITS") & String.Empty)
                rowRSTRETL1.Item("AMT_SOLD") = Val(rowEDTRETLS.Item("EXTENSION") & String.Empty)
                rowRSTRETL1.Item("QTY_EOW") = 0
            Else
                MsgBox("Unrecognized Transaction Type: " & EDI_ZA_TRAN_TYPE)
                Stop
            End If


            If rowRSTRETL1.Item("ITEM_CODE") & "" <> "" Then
                ' for some reason, IGNORE_ITEM_CODE is not getting communicated from RSTITEM1 to EDTRETLS

                'If rowRSTRETL1.Item("ITEM_CODE") & "" = "" Then Stop


                rowRSTRETL1.Item("OPS_YYYYPP") = rowEDTRETLS.Item("OPS_YYYYPP")
                rowRSTRETL1.Item("OPS_YYYYWW") = rowEDTRETLS.Item("OPS_YYYYWW")
                dst.Tables(RSTRETL1).Rows.Add(rowRSTRETL1)

                '' ***** Fill EDT852T2 *****
                'rowEDT852T2 = dst.Tables("EDT852T2").NewRow
                'rowEDT852T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                'rowEDT852T2.Item("EDI_LINE_NO") = EDI_LINE_NO
                'rowEDT852T2.Item("EDI_ITEM_UP") = rowEDTRETLS.Item("UPC") & String.Empty
                'rowEDT852T2.Item("EDI_ITEM_EN") = rowEDTRETLS.Item("ITEM_CODE") & String.Empty
                'dst.Tables("EDT852T2").Rows.Add(rowEDT852T2)

                '' ***** Fill EDT852T3 *****
                'rowEDT852T3 = dst.Tables("EDT852T3").NewRow
                'rowEDT852T3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                'rowEDT852T3.Item("EDI_LINE_NO") = EDI_LINE_NO
                'rowEDT852T3.Item("EDI_ZA_NO") = "1"
                'rowEDT852T3.Item("EDI_ZA_TRAN_TYPE") = rowEDTRETLS.Item("EDI_ZA_TRAN_TYPE") & String.Empty
                'rowEDT852T3.Item("EDI_SDQ_STORE_01") = rowEDTRETLS.Item("DOORID") & String.Empty
                'rowEDT852T3.Item("EDI_SDQ_QTY_AMT_01") = Val(rowEDTRETLS.Item("UNITS") & String.Empty)
                'dst.Tables("EDT852T3").Rows.Add(rowEDT852T3)

                EDI_LINE_NO += 1

            End If

        Next

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    ''' <summary>
    ''' Download LGI Retail sales MDB fiules
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DownloadLGIdata() As Boolean

        Dim fileName As String = String.Empty
        Dim iFiles As Integer = 0

        Try
            ASCMAIN1.Progress("Downloading from " & MyBase.Absx1.txtFor("SO_PARM_FTP_HOST_NAME").Text, String.Empty)
            DownloadLGIdata = False

            Dim ftpEDCFTPC1 As New RSCFTPC1("LGI")
            Dim extList As List(Of String) = New List(Of String)
            extList.Add(SO_PARM_DOWNLOAD_FILE_EXT)
            ftpEDCFTPC1.DownLoadFilesFileExtension = extList

            Dim numFiles As Integer = ftpEDCFTPC1.DownloadFiles(False, True)

            Dim errorMsg As String = String.Empty
            If numFiles = 0 Then
                errorMsg &= "No files downloaded." & Environment.NewLine
            End If

            For Each errMsg As String In ftpEDCFTPC1.ErrorList
                errorMsg &= errMsg & Environment.NewLine
            Next

            If errorMsg <> String.Empty Then
                MessageBox.Show(errorMsg, "Ftp Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            DownloadLGIdata = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "FTP Files", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DownloadLGIdata = False
        End Try

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Function

    ''' <summary>
    ''' Extract the data from the LGI MDB Tables
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillOracleTables(ByVal RS_CNTL_NO As String, ByVal mdbFilename As String)

        Dim databaseName As String = String.Empty
        Dim dbLocation As String = String.Empty
        Dim oracleTableName As String = String.Empty

        If Not SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
            SO_PARM_LOCAL_DIR_INBOUND = SO_PARM_LOCAL_DIR_INBOUND & "\"
        End If

        dbLocation = SO_PARM_LOCAL_DIR_INBOUND & mdbFilename
        Dim clsEDCPRMDB As RSCPRMDB = New RSCPRMDB
        clsEDCPRMDB.DisplayMessageControl = Me.txtmsg
        clsEDCPRMDB.CreateConnectionSting(4.0, dbLocation, String.Empty, String.Empty)
        clsEDCPRMDB.OracleTableNamePrefix = Me.SO_PARM_MDB_PROVIDER & "_"
        clsEDCPRMDB.OracleSchema = oracleLgiSchemaName

        CNTL_NO = RS_CNTL_NO ' this will be used in Update

        For Each mdbTableName As String In LgiTableNames

            ASCMAIN1.Progress("Importing Table " & mdbTableName & " from " & mdbFilename)

            Dim tbl As DataTable = New DataTable
            Dim CNTL_LNO As Integer = 1
            'tbl = clsEDCPRMDB.CreateOracleTableFromMdbSchema(mdbTABLES, mdbTABLES)
            oracleTableName = SO_PARM_MDB_PROVIDER & "_" & mdbTableName
            tbl = clsEDCPRMDB.CopyMdbIntoOracle(mdbTableName, oracleTableName, False)
            If tbl Is Nothing Then
                Stop
            Else
                ASCMAIN1.Progress("Assigning Control Keys for " & mdbTableName)
                Dim rowCount As Long = tbl.Rows.Count
                For Each row As DataRow In tbl.Rows
                    row.Item("CNTL_NO") = RS_CNTL_NO ' "0000000000" ' RS_CNTL_NO
                    row.Item("CNTL_LNO") = CNTL_LNO
                    CNTL_LNO += 1

                    If CNTL_LNO Mod 10 = 0 Then
                        ASCMAIN1.Progress("-", Format((CNTL_LNO / rowCount) * 100, "0.00") & "%")
                    End If
                    dst.Tables(oracleTableName).ImportRow(row)
                    Application.DoEvents()
                Next
                ' Cannot use this since tbl is stripped of it keys, you get an error
                'dst.Tables(mdbTABLES).Merge(tbl)
            End If

            If mdbTableName = "SALESFILE" Or mdbTableName = "INVFILE" Then
                ASCMAIN1.Progress("Removing Previously Loaded Data from " & mdbTableName)
                Dim TABLE_NAME As String = SO_PARM_MDB_PROVIDER & "_" & mdbTableName
                Dim ID As String = "SALEID"
                If mdbTableName = "INVFILE" Then
                    ID = "INVENTORYID"
                End If
                ASCMAIN1.sql = "Select Distinct " & ID & " from " & TABLE_NAME
                For Each row2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim row1 As DataRow = dst.Tables(oracleTableName).Rows.Find(row2.Item(0))
                    If row1 IsNot Nothing Then
                        row1.Delete()
                    End If
                Next
            End If

            'ASCMAIN1.Progress("Updating Oracle with Table " & mdbTableName)
            'Dim TABLE_NAME As String = SO_PARM_MDB_PROVIDER & "_" & mdbTableName
            'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where CNTL_NO = '0000000000'")
            'Call Update_Record_TDA(TABLE_NAME)
            'If mdbTableName = "SALESFILE" Or mdbTableName = "INVFILE" Then
            '    Dim ID As String = "SALESID"
            '    If mdbTableName = "INVFILE" Then
            '        ID = "INVENTORYID"
            '    End If
            '    ASCMAIN1.sql = "Select Max (CNTL_NO) from " & TABLE_NAME & " where CNTL_NO > '0000000000'"
            '    Dim CNTL_NO_last As String = ASCDATA1.GetDataValue
            '    Dim SQL As String = "Select * from " & TABLE_NAME
            '    If CNTL_NO_last <> "" Then
            '        SQL &= "" _
            '        & " where " & ID & " in (Select " & ID & " from " & TABLE_NAME & " where CNTL_NO = '0000000000'" _
            '        & " minus Select " & ID & " from " & TABLE_NAME & " where CNTL_NO = '" & CNTL_NO_last & "')" _
            '        & " where CNTL_NO = '0000000000'"
            '        Fill_Records(TABLE_NAME, , , SQL)
            '    End If
            'End If
        Next


        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    ''' <summary>
    ''' Sets up shells for orphaned customers
    ''' These customers report Sales or Inventory but do not appear in the Door master file
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetOrphanedCustomers()

        Dim tblOrphans As DataTable = Me.SelectDistinct(dst.Tables("EDTRETLS"), New String() {"RETAILERID"})

        For Each rowOrphans As DataRow In tblOrphans.Rows
            If dst.Tables("RSTCUST1").Select("RETAILERID = '" & rowOrphans.Item("RETAILERID") & "'").Length = 0 Then
                Dim rowRSTCUST1 As DataRow = dst.Tables("RSTCUST1").NewRow
                rowRSTCUST1.Item("RetailerID") = rowOrphans.Item("RETAILERID")
                rowRSTCUST1.Item("NEEDS_MAPPING") = "1"
                rowRSTCUST1.Item("Ignore") = "1"
                rowRSTCUST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowRSTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowRSTCUST1.Item("INIT_DATE") = DateTime.Now
                rowRSTCUST1.Item("LAST_DATE") = DateTime.Now
                rowRSTCUST1.Item("CUST_STORE_NO") = "000000"
                rowRSTCUST1.Item("CUST_STORE_NAME") = "Unknown Customer"
                rowRSTCUST1.Item("CUST_STORE_ADDR1") = String.Empty
                rowRSTCUST1.Item("CUST_STORE_ADDR2") = String.Empty
                rowRSTCUST1.Item("CUST_STORE_CITY") = String.Empty
                rowRSTCUST1.Item("CUST_STORE_STATE") = String.Empty
                rowRSTCUST1.Item("CUST_STORE_ZIP_CODE") = String.Empty
                rowRSTCUST1.Item("CUST_STORE_PHONE") = String.Empty
                dst.Tables("RSTCUST1").Rows.Add(rowRSTCUST1)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Get LGI databases and inmport the data
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetUnprocessedLGIDatabases()
        Call ASCMAIN1.Progress("Loading Unprocessed Database Files")

        Dim rowRSTLGIED As DataRow = Nothing
        Dim rowRSTLGIEDx As DataRow = Nothing
        Dim fileName As String = String.Empty

        Dim RS_CNTL_NO As String = String.Empty

        Dim fld As New System.IO.DirectoryInfo(SO_PARM_LOCAL_DIR_INBOUND)

        For Each mdbFile As System.IO.FileInfo In fld.GetFiles("*.mdb")
            fileName = mdbFile.Name
            rowRSTLGIEDx = ASCDATA1.GetDataRow("Select * from RSTLGIED Where Import_Date Is Null And File_Name = '" & fileName & "'")

            rowRSTLGIED = dst.Tables("RSTLGIED").NewRow
            rowRSTLGIED.Item("IMPORT") = "1"

            If rowRSTLGIEDx IsNot Nothing Then
                rowRSTLGIED.Item("RS_CNTL_NO") = rowRSTLGIEDx.Item("RS_CNTL_NO") & String.Empty
                rowRSTLGIED.Item("USER_ID") = rowRSTLGIEDx.Item("USER_ID") & String.Empty
                rowRSTLGIED.Item("DOWNLOAD_DATE") = rowRSTLGIEDx.Item("DOWNLOAD_DATE") & String.Empty
            Else
                rowRSTLGIED.Item("RS_CNTL_NO") = ASCMAIN1.Next_Control_No("RSTLGIED.RS_CNTL_NO")
                rowRSTLGIED.Item("USER_ID") = ASCMAIN1.USER_ID
                rowRSTLGIED.Item("DOWNLOAD_DATE") = DateTime.Now
            End If

            ' Commom Value Fields
            rowRSTLGIED.Item("FILE_NAME") = fileName
            rowRSTLGIED.Item("FILE_DATE") = mdbFile.LastWriteTime
            rowRSTLGIED.Item("IMPORT_USER") = ASCMAIN1.USER_ID
            rowRSTLGIED.Item("IMPORT_DATE") = DateTime.Now
            dst.Tables("RSTLGIED").Rows.Add(rowRSTLGIED)

            RS_CNTL_NO = rowRSTLGIED.Item("RS_CNTL_NO")
            Me.FillOracleTables(RS_CNTL_NO, fileName)

            Exit For ' can only do 1 at a time now that we are checking 2 year spans - wjz 08/17/09

        Next

        If dst.Tables("RSTLGIED").Rows.Count > 0 Then
            Me.grdRSTLGIED.Text = CStr(dst.Tables("RSTLGIED").Rows.Count) & " MDBs located in " & SO_PARM_LOCAL_DIR_INBOUND
        Else
            Me.grdRSTLGIED.Text = SO_PARM_LOCAL_DIR_INBOUND & " contains no MDBs."
        End If

        Call ASCMAIN1.Progress("")
    End Sub

    ''' <summary>
    ''' Maps the UPCs to Item Codes
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateRetailSalesRecords()

        Dim rowRSTCUST1 As DataRow = Nothing
        Dim rowEDTRETLS As DataRow = Nothing
        Dim rowGLTPARM3 As DataRow = Nothing
        Dim DATA_DATE As Date = Nothing ' String.Empty

        Dim ITEM_CODE As String = String.Empty
        Dim ITEM_DESC As String = String.Empty
        Dim ITEM_UPC_CODE As String = String.Empty
        Dim ITEM_SKU_CODE As String = String.Empty
        Dim lastUpcCode As String = String.Empty
        Dim lastSku As String = String.Empty

        Dim sql As String = String.Empty

        Dim RetailerID As String = String.Empty
        Dim lastRetailerID As String = String.Empty
        Dim CUST_CODE As String = String.Empty

        Dim DoorID As String = String.Empty
        Dim numRecs As Long = 100
        Dim ignoreRetails As String = String.Empty
        Dim ignoreItem As String = String.Empty

        Dim salesTables() As String = New String() {"LGI_INVFILE", "LGI_SALESFILE"}

        For Each tableName As String In salesTables
            ASCMAIN1.Progress("Processing " & tableName, "")
            numRecs = 0
            For Each rowTable As DataRow In dst.Tables(tableName).Select(String.Empty, "UPC")
                ITEM_UPC_CODE = (rowTable.Item("UPC") & String.Empty).ToString.Trim
                ITEM_SKU_CODE = (rowTable.Item("SKU") & String.Empty).ToString.Trim

                If ITEM_SKU_CODE.Length = 0 Then
                    ITEM_SKU_CODE = ItemCodeNotMapped
                End If

                If ITEM_UPC_CODE.Length = 0 Then
                    ITEM_UPC_CODE = ItemCodeNotMapped
                End If

                DoorID = ASCMAIN1.Format_Field((rowTable.Item("DoorID") & String.Empty).ToString.Trim, "CUST_STORE_NO")

                If numRecs Mod 100 = 0 Then
                    ASCMAIN1.Progress("-", Format((numRecs / dst.Tables(tableName).Rows.Count) * 100, "0.00") & "%")
                    Application.DoEvents()
                End If

                numRecs += 1

                'If Val(rowTable.Item("RETAILERID") & "") = 23103 Then Stop

                ' Update the Item Code and Description
                If (ITEM_UPC_CODE.Length > 0 And ITEM_UPC_CODE <> "<Blank>") Or (ITEM_SKU_CODE.Length > 0 And ITEM_SKU_CODE <> "<Blank>") Then
                    If (ITEM_UPC_CODE.Length > 0 And ITEM_UPC_CODE <> "<Blank>") Then
                        If ITEM_UPC_CODE <> lastUpcCode Then
                            ITEM_CODE = String.Empty
                            ITEM_DESC = String.Empty
                            lastUpcCode = ITEM_UPC_CODE
                            lastSku = ITEM_SKU_CODE
                            rowICTITEM1 = MyBase.LookUp("ICTITEM1", ITEM_UPC_CODE)

                            If rowICTITEM1 Is Nothing Then
                                rowICTITEM1 = MyBase.LookUp("ICTITEM1", ITEM_SKU_CODE)
                            End If

                            If rowICTITEM1 IsNot Nothing Then
                                ITEM_CODE = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                                ITEM_DESC = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                            End If
                        End If
                    Else
                        If ITEM_SKU_CODE <> lastSku Then
                            ITEM_CODE = String.Empty
                            ITEM_DESC = String.Empty
                            lastUpcCode = ITEM_UPC_CODE
                            lastSku = ITEM_SKU_CODE
                            rowICTITEM1 = MyBase.LookUp("ICTITEM1", ITEM_SKU_CODE)

                            If rowICTITEM1 IsNot Nothing Then
                                ITEM_CODE = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                                ITEM_DESC = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                            End If
                        End If
                    End If
                Else
                    ITEM_CODE = String.Empty
                    ITEM_DESC = String.Empty
                End If

                ignoreItem = "0"
                If ITEM_CODE.Length = 0 Then
                    Try
                        Dim rowRSTITEM1 As DataRow = dst.Tables("RSTITEM1").Select("SKU = '" & ITEM_SKU_CODE & "' AND UPC = '" & ITEM_UPC_CODE & "'")(0)
                        If rowRSTITEM1 IsNot Nothing Then
                            ignoreItem = rowRSTITEM1.Item("IGNORE") & String.Empty
                            If (rowRSTITEM1.Item("ITEM_CODE") & String.Empty).ToString.Trim.Length > 0 And ignoreItem <> "1" Then
                                ITEM_CODE = (rowRSTITEM1.Item("ITEM_CODE") & String.Empty).ToString.Trim
                                If ITEM_CODE.Length > 0 Then
                                    rowICTITEM1 = MyBase.LookUp("ICTITEM1", ITEM_CODE)
                                    If rowICTITEM1 IsNot Nothing Then
                                        ITEM_CODE = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                                        ITEM_DESC = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                                    Else
                                        ITEM_CODE = String.Empty
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        ' Nothing
                    End Try
                End If

                If ITEM_CODE.Length = 0 Then
                    If dst.Tables("RSTITEM1").Select("SKU = '" & ITEM_SKU_CODE & "'").Length = 0 Then
                        Dim rowRSTITEM1 As DataRow = dst.Tables("RSTITEM1").NewRow
                        rowRSTITEM1.Item("SKU") = ITEM_SKU_CODE
                        rowRSTITEM1.Item("UPC") = ITEM_UPC_CODE
                        rowRSTITEM1.Item("ITEM_CODE") = ""
                        rowRSTITEM1.Item("IGNORE") = "1"
                        ignoreItem = "1"
                        rowRSTITEM1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowRSTITEM1.Item("INIT_DATE") = DateTime.Now
                        rowRSTITEM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowRSTITEM1.Item("LAST_DATE") = DateTime.Now
                        dst.Tables("RSTITEM1").Rows.Add(rowRSTITEM1)
                    End If
                End If

                ' Update the Customer Code
                If lastRetailerID <> rowTable.Item("RetailerID") & String.Empty Then
                    lastRetailerID = (rowTable.Item("RetailerID") & String.Empty).ToString.Trim
                    RetailerID = rowTable.Item("RetailerID") & String.Empty
                    CUST_CODE = CustCodeNotMapped
                    sql = "RetailerID = '" & RetailerID & "'"
                    Try
                        rowRSTCUST1 = dst.Tables("RSTCUST1").Select(sql)(0)
                        CUST_CODE = rowRSTCUST1.Item("CUST_CODE") & String.Empty
                        ignoreRetails = rowRSTCUST1.Item("IGNORE") & String.Empty

                    Catch ex As Exception
                        Stop
                        ' Nothing
                    End Try
                End If

                rowEDTRETLS = dst.Tables("EDTRETLS").NewRow
                rowEDTRETLS.Item("CNTL_NO") = rowTable.Item("CNTL_NO") & String.Empty
                rowEDTRETLS.Item("CNTL_LNO") = rowTable.Item("CNTL_LNO") & String.Empty
                rowEDTRETLS.Item("CUST_CODE") = CUST_CODE
                rowEDTRETLS.Item("CUST_STORE_NO") = DoorID
                'rowEDTRETLS.Item("TRANS_TYPE") = tableName.Substring(0, 1)
                rowEDTRETLS.Item("ITEM_CODE") = ITEM_CODE
                'If ASCMAIN1.Running_in_VS Then
                '    If ITEM_CODE = "" Then Stop
                'End If
                rowEDTRETLS.Item("ITEM_DESC") = ITEM_DESC
                rowEDTRETLS.Item("DATA_DATE") = rowTable.Item("DATA_DATE") ' & String.Empty
                rowEDTRETLS.Item("UNITS") = rowTable.Item("UNITS") & String.Empty

                If Val(rowEDTRETLS.Item("UNITS")) < 0 Then
                    rowEDTRETLS.Item("EDI_TRAN_TYPE") = ediTransTypeQtyReturned
                Else
                    rowEDTRETLS.Item("EDI_TRAN_TYPE") = ediTransTypeQtySold
                End If

                If tableName = "LGI_INVFILE" Then
                    rowEDTRETLS.Item("EDI_ZA_TRAN_TYPE") = ediTransTypeQtyInvAvaliable
                ElseIf tableName = "LGI_SALESFILE" Then
                    rowEDTRETLS.Item("EDI_ZA_TRAN_TYPE") = rowEDTRETLS.Item("EDI_TRAN_TYPE")
                Else
                    Stop
                End If

                rowEDTRETLS.Item("MSRP") = Val(rowTable.Item("MSRP") & String.Empty)

                rowEDTRETLS.Item("RETAILERID") = rowTable.Item("RETAILERID") & String.Empty
                rowEDTRETLS.Item("DOORID") = rowTable.Item("DOORID") & String.Empty
                rowEDTRETLS.Item("UPC") = rowTable.Item("UPC") & String.Empty
                rowEDTRETLS.Item("SKU") = rowTable.Item("SKU") & String.Empty
                rowEDTRETLS.Item("IMPORT") = "1"
                rowEDTRETLS.Item("IGNORE_ITEM_CODE") = ignoreItem

                rowEDTRETLS.Item("IGNORE") = ignoreRetails

                Try
                    DATA_DATE = CDate(rowTable.Item("DATA_DATE") & String.Empty) ' .ToString("dd-MMM-yyyy")
                Catch ex As Exception
                    DATA_DATE = CDate(DateTime.Now).Date ' .ToString("dd-MMM-yyyy")
                End Try

                Dim rowRSTLGIDX As DataRow = dst.Tables("RSTLGIDX").Rows.Find(DATA_DATE)
                If rowRSTLGIDX Is Nothing Then
                    rowRSTLGIDX = dst.Tables("RSTLGIDX").NewRow

                    'sql = "SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE = (SELECT MAX(WEEK_END_DATE) WEEK_END_DATE FROM GLTPARM3 WHERE WEEK_END_DATE <= '" & DATA_DATE & "')"
                    'rowGLTPARM3 = ASCDATA1.GetDataRow(sql)

                    sql = "Select * From GLTPARM3 Where WEEK_END_DATE = " _
                    & " (SELECT MIN(WEEK_END_DATE) WEEK_END_DATE from GLTPARM3" _
                    & " where WEEK_END_DATE >= '" & Format(DATA_DATE, "dd-MMM-yyyy") & "')"
                    rowGLTPARM3 = ASCDATA1.GetDataRow(sql)

                    rowRSTLGIDX.Item("DATA_DATE") = DATA_DATE
                    rowRSTLGIDX.Item("WEEK_END_DATE") = rowGLTPARM3.Item("WEEK_END_DATE")
                    rowRSTLGIDX.Item("YYYYWW") = rowGLTPARM3.Item("YYYYWW")
                    rowRSTLGIDX.Item("YYYYPP") = rowGLTPARM3.Item("YYYYPP")
                    dst.Tables("RSTLGIDX").Rows.Add(rowRSTLGIDX)
                End If

                rowEDTRETLS.Item("EDI_FROM_DATE") = CDate(rowRSTLGIDX.Item("WEEK_END_DATE")).AddDays(-6)
                rowEDTRETLS.Item("EDI_TO_DATE") = rowRSTLGIDX.Item("WEEK_END_DATE")
                rowEDTRETLS.Item("OPS_YYYYPP") = rowRSTLGIDX.Item("YYYYPP")
                rowEDTRETLS.Item("OPS_YYYYWW") = rowRSTLGIDX.Item("YYYYWW")

                dst.Tables("EDTRETLS").Rows.Add(rowEDTRETLS)
            Next
        Next

        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub LoadUnMappedItems()

        Call ASCMAIN1.Progress("Loading Items")

        Dim Sql As String = "Select * From RSTITEM1"
        Fill_Records("RSTITEM1", String.Empty, True, Sql)

    End Sub

    ''' <summary>
    ''' Maps the LGI Retailer IDs to the proper ABSolution Customer
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub MapRetailers()

        Call ASCMAIN1.Progress("Loading Customer Doors")

        Dim RetailerID As String = String.Empty
        Dim DoorID As String = String.Empty
        Dim wkString As String = String.Empty

        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowRSTCUST1 As DataRow = Nothing

        Dim Sql As String = "Select * From RSTCUST1"
        Fill_Records("RSTCUST1", String.Empty, True, Sql)

        For Each rowRSTCUST1 In dst.Tables("RSTCUST1").Rows
            If rowRSTCUST1.Item("IGNORE").ToString.Trim.Length = 0 Then
                rowRSTCUST1.Item("IGNORE") = "0"
            End If

            If rowRSTCUST1.Item("CUST_CODE").ToString.Trim.Length = 0 Then
                rowRSTCUST1.Item("CUST_CODE") = CustCodeNotMapped
            Else
                Try
                    rowARTCUST1 = Nothing
                    rowARTCUST1 = dst.Tables("ARTCUST1").Select("CUST_CODE = '" & rowRSTCUST1.Item("CUST_CODE") & "'")(0)
                    rowRSTCUST1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    rowRSTCUST1.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                    rowRSTCUST1.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                    rowRSTCUST1.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                    rowRSTCUST1.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    rowRSTCUST1.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                    rowRSTCUST1.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                Catch ex As Exception
                    ' NOTHING
                End Try
            End If
        Next

        Dim tblDistinctRetailers As DataTable = Me.SelectDistinct(dst.Tables("LGI_DOORFILE"), New String() {"RetailerID"})

        For Each rowRetailer As DataRow In tblDistinctRetailers.Rows
            RetailerID = rowRetailer.Item("RetailerID") & String.Empty
            DoorID = ASCMAIN1.Format_Field((rowRetailer.Item("DOORID") & String.Empty).ToString.Trim, "CUST_STORE_NO")

            ' See if the Retailer Id is known, if not then auto fill the tables
            Sql = "RetailerID = '" & RetailerID & "'"
            If dst.Tables("RSTCUST1").Select(Sql).Length = 0 Then
                rowRSTCUST1 = dst.Tables("RSTCUST1").NewRow
                rowRSTCUST1.Item("RetailerID") = RetailerID
                rowRSTCUST1.Item("CUST_CODE") = CustCodeNotMapped
                rowRSTCUST1.Item("IGNORE") = "1"
                rowRSTCUST1.Item("NEEDS_MAPPING") = "1"
                rowRSTCUST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowRSTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowRSTCUST1.Item("INIT_DATE") = DateTime.Now
                rowRSTCUST1.Item("LAST_DATE") = DateTime.Now
                'rowRSTCUST1.Item("CUST_STORE_NO") = DoorID
                rowRSTCUST1.Item("CUST_NAME") = (rowRetailer.Item("CHAIN_NAME") & String.Empty).ToString.PadRight(60, " ").Trim
                rowRSTCUST1.Item("CUST_ADDR1") = (rowRetailer.Item("ADDRESS1") & String.Empty).ToString.PadRight(60, " ").Trim
                rowRSTCUST1.Item("CUST_ADDR2") = (rowRetailer.Item("ADDRESS2") & String.Empty).ToString.PadRight(60, " ").Trim
                rowRSTCUST1.Item("CUST_CITY") = (rowRetailer.Item("CITY") & String.Empty).ToString.PadRight(30, " ").Trim
                rowRSTCUST1.Item("CUST_STATE") = (rowRetailer.Item("STATE") & String.Empty).ToString.PadRight(2, " ").Trim
                rowRSTCUST1.Item("CUST_ZIP_CODE") = (rowRetailer.Item("ZIP") & String.Empty).ToString.PadRight(10, " ").Trim

                wkString = String.Empty
                For Each chrStr As Char In rowRetailer.Item("PHONE") & String.Empty
                    If Char.IsDigit(chrStr) Then
                        wkString &= chrStr
                    End If
                Next
                If wkString.Length > 10 Then wkString = wkString.Substring(0, 10)
                rowRSTCUST1.Item("CUST_PHONE") = wkString
                dst.Tables("RSTCUST1").Rows.Add(rowRSTCUST1)
            Else
                ' if Unmapped Customer then reload the Address and Name Information
                Try
                    rowRSTCUST1 = dst.Tables("RSTCUST1").Select(Sql)(0)
                    If rowRSTCUST1.Item("CUST_CODE") = CustCodeNotMapped Then
                        rowRSTCUST1.Item("CUST_NAME") = (rowRetailer.Item("CHAIN_NAME") & String.Empty).ToString.PadRight(60, " ").Trim
                        rowRSTCUST1.Item("CUST_ADDR1") = (rowRetailer.Item("ADDRESS1") & String.Empty).ToString.PadRight(60, " ").Trim
                        rowRSTCUST1.Item("CUST_ADDR2") = (rowRetailer.Item("ADDRESS2") & String.Empty).ToString.PadRight(60, " ").Trim
                        rowRSTCUST1.Item("CUST_CITY") = (rowRetailer.Item("CITY") & String.Empty).ToString.PadRight(30, " ").Trim
                        rowRSTCUST1.Item("CUST_STATE") = (rowRetailer.Item("STATE") & String.Empty).ToString.PadRight(2, " ").Trim
                        rowRSTCUST1.Item("CUST_ZIP_CODE") = (rowRetailer.Item("ZIP") & String.Empty).ToString.PadRight(10, " ").Trim
                    End If
                Catch ex As Exception
                    ' Nothing
                End Try
            End If
        Next
    End Sub

    Private Function SelectDistinct(ByRef dtSource As DataTable, ByRef columnNames() As String) As DataTable
        Return SelectDistinct(dtSource, columnNames, String.Empty)
    End Function

    ''' <summary>
    ''' Selects Distinct Column values from a datatable
    ''' </summary>
    ''' <param name="dtSource"></param>
    ''' <param name="columnNames"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SelectDistinct(ByRef dtSource As DataTable, ByRef columnNames() As String, ByVal WhereClause As String) As DataTable

        Dim dtOut As DataTable = New DataTable

        dtOut = dtSource.Clone

        Dim distinctValue As String = String.Empty
        Dim distinctValueLast As String = String.Empty

        Try
            For Each dr As DataRow In dtSource.Select(WhereClause, String.Join(", ", columnNames)) 'sort the table
                distinctValue = String.Empty
                For count As Integer = columnNames.GetLowerBound(0) To columnNames.GetUpperBound(0)
                    If dr(columnNames(count)) IsNot Nothing Then
                        distinctValue &= CStr(dr(columnNames(count)))
                    End If
                    distinctValue &= ","
                Next

                If Not distinctValueLast.Equals(distinctValue) Then
                    Try
                        dtOut.ImportRow(dr)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Select Distinct", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    distinctValueLast = distinctValue
                End If

            Next dr
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Select Distinct", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dtOut

    End Function

    ''' <summary>
    ''' Creates any New Customer Doors
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateCustomerDoors()

        ASCMAIN1.Progress("Create Customer Doors", String.Empty)

        Dim fields() As String = New String() {"CUST_CODE", "CUST_STORE_NO", "RETAILERID", "DOORID"}
        Dim sql As String = String.Empty

        Dim rowARTCUST2 As DataRow = Nothing
        Dim tblDoors As DataTable = Me.SelectDistinct(viewRetails.ToTable, fields)

        For Each rowDoors As DataRow In tblDoors.Rows
            Dim CUST_CODE As String = rowDoors.Item("CUST_CODE") & String.Empty
            Dim CUST_STORE_NO As String = rowDoors.Item("CUST_STORE_NO") & String.Empty

            ASCMAIN1.Progress("-", CUST_CODE)

            sql = "Select * from ARTCUST2 Where CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "'"
            Dim rowARTCUST2x As DataRow = ASCDATA1.GetDataRow(sql)

            If rowARTCUST2x IsNot Nothing Then
                Continue For
            End If

            sql = "RETAILERID = '" & rowDoors.Item("RETAILERID") & "' AND DOORID = '" & rowDoors.Item("DOORID") & "'"
            If dst.Tables("LGI_DOORFILE").Select(sql).Length > 0 Then
                rowARTCUST2x = dst.Tables("LGI_DOORFILE").Select(sql)(0)
            End If

            If rowARTCUST2x Is Nothing Then
                sql = "RETAILERID = '" & rowDoors.Item("RETAILERID") & "'"
                If dst.Tables("RSTCUST1").Select(sql).Length > 0 Then
                    rowARTCUST2x = dst.Tables("RSTCUST1").Select(sql)(0)
                End If
            End If

            If rowARTCUST2x Is Nothing Then
                Stop
            End If

            rowARTCUST2 = dst.Tables("ARTCUST2").NewRow

            Select Case rowARTCUST2x.Table.TableName
                Case "LGI_DOORFILE"
                    rowARTCUST2.Item("CUST_CODE") = rowDoors.Item("CUST_CODE") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_NO") = rowDoors.Item("CUST_STORE_NO") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_NAME") = rowARTCUST2x.Item("CHAIN_NAME") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_ADDR1") = rowARTCUST2x.Item("ADDRESS1") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_ADDR2") = rowARTCUST2x.Item("ADDRESS2") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_CITY") = rowARTCUST2x.Item("CITY") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_STATE") = rowARTCUST2x.Item("STATE") & String.Empty

                    Dim wkString As String = rowARTCUST2x.Item("ZIP") & String.Empty
                    If wkString.Trim.ToUpper = "UNKNOWN" Then
                        wkString = String.Empty
                    End If
                    If wkString.Length > 10 Then wkString = wkString.Substring(0, 10)
                    rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = wkString

                    wkString = String.Empty
                    For Each chrStr As Char In rowARTCUST2x.Item("PHONE") & String.Empty
                        If Char.IsDigit(chrStr) Then
                            wkString &= chrStr
                        End If
                    Next
                    If wkString.Length > 10 Then wkString = wkString.Substring(0, 10)
                    rowARTCUST2.Item("CUST_STORE_PHONE") = wkString

                Case "RSTCUST1"
                    rowARTCUST2.Item("CUST_CODE") = rowDoors.Item("CUST_CODE") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_NO") = rowDoors.Item("CUST_STORE_NO") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_NAME") = rowARTCUST2x.Item("CUST_NAME") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_ADDR1") = rowARTCUST2x.Item("CUST_ADDR1") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_ADDR2") = rowARTCUST2x.Item("CUST_ADDR2") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_CITY") = rowARTCUST2x.Item("CUST_CITY") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_STATE") = rowARTCUST2x.Item("CUST_STATE") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = rowARTCUST2x.Item("CUST_ZIP_CODE") & String.Empty
                    rowARTCUST2.Item("CUST_STORE_PHONE") = rowARTCUST2x.Item("CUST_PHONE") & String.Empty

                Case Else
                    Stop
            End Select

            dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)
        Next

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

#End Region

#Region "From Controls Events"

    Private Sub grdRSTCUST1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTCUST1.AfterRowActivate
        If grdRSTCUST1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdRSTCUST1.ActiveRow.Cells("CUST_NAME").Text.Trim.Length = 0 Then
            grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
        End If

    End Sub

    ''' <summary>
    ''' Prevent a Customer Codeless retailer to get loaded into Retail Sales
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub grdRSTCUST1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTCUST1.BeforeRowUpdate

        If grdRSTCUST1.ActiveRow.Cells("CUST_CODE").Text.Trim <> CustCodeNotMapped Then
            Dim retailerid As String = grdRSTCUST1.ActiveRow.Cells("RETAILERID").Text.Trim
            Dim CUST_CODE As String = grdRSTCUST1.ActiveRow.Cells("CUST_CODE").Text.Trim
            Dim row As DataRow = Nothing

            Try
                row = dst.Tables("RSTCUST1").Select("RETAILERID <> '" & retailerid & "' AND CUST_CODE = '" & CUST_CODE & "'")(0)
            Catch ex As Exception
                ' Nothing
            End Try

            If row IsNot Nothing Then
                MessageBox.Show("Retailer " & row.Item("RetailerID") & " is already assigned to customer " & CUST_CODE, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If

            If grdRSTCUST1.ActiveRow.Cells("CUST_NAME").Text.Trim = String.Empty Then
                MessageBox.Show("Retailer Name is required.", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                e.Cancel = True
                Exit Sub
            End If

        End If

        Try
            If grdRSTCUST1.ActiveRow.Cells("CUST_CODE").Text.Trim = String.Empty Or grdRSTCUST1.ActiveRow.Cells("CUST_CODE").Value = CustCodeNotMapped Then
                grdRSTCUST1.ActiveRow.Cells("IGNORE").Value = "1"
                grdRSTCUST1.ActiveRow.Cells("CUST_CODE").Value = CustCodeNotMapped
            Else
                grdRSTCUST1.ActiveRow.Cells("IGNORE").Value = "0"
            End If
        Catch ex As Exception
            ' Nothing
        End Try
    End Sub

    Private Sub grdRSTCUST1_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTCUST1.CellChange
        If grdRSTCUST1.ActiveRow Is Nothing Then Exit Sub
        If grdRSTCUST1.ActiveCell.Column.Key <> "IGNORE" Then Exit Sub

        Try
            grdRSTCUST1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cell Change Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub grdRSTCUST1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTCUST1.ClickCellButton
        Select Case grdRSTCUST1.ActiveCell.Column.Key
            Case "CUST_CODE"

                Try
                    Dim CUST_CODE_PREV As String = e.Cell.Text
                    'grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    MyBase.grdClickCellButton(grdRSTCUST1, String.Empty, False)
                    If e.Cell.Text <> "" And CUST_CODE_PREV = "" Then
                        e.Cell.Row.Cells("IGNORE").Value = "0"
                    End If
                    'grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    'Stop
                Catch ex As Exception
                    'grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    MessageBox.Show(ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Finally
                    Try
                        'grdRSTCUST1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                End Try

            Case Else
                Exit Sub

        End Select

        grdRSTCUST1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
    End Sub

    ''' <summary>
    ''' Initialize Grid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub grdRSTCUST1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTCUST1.InitializeLayout
        'e.Layout.Bands(0).Columns("CUST_CODE").ValueList = Me.uddCUST_CODE
    End Sub

    ''' <summary>
    ''' Used to show prompt on the screen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtmsg_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtmsg.ValueChanged
        'If Val(txtmsg.Text) Mod 100 = 0 Then
        ASCMAIN1.Progress("-", txtmsg.Text)
        'End If
    End Sub

    Private Sub grdRSTLGIED_CellChange(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTLGIED.CellChange
        If grdRSTLGIED.ActiveRow Is Nothing Then Exit Sub
        If grdRSTLGIED.ActiveCell.Column.Key <> "IMPORT" Then Exit Sub

        Try
            grdRSTLGIED.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cell Change Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub grdRSTITEM1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTITEM1.AfterRowActivate
        If grdRSTITEM1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        'If grdRSTITEM1.ActiveRow.Cells("ITEM_CODE").Text.Trim.Length = 0 Then
        '    grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        'Else
        '    grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        'End If

        ' Initialize the variables so the After Update can change the values
        Me.OriginalItemCode = (grdRSTITEM1.ActiveRow.Cells("ITEM_CODE").Value & String.Empty).ToString.Trim
        Me.NewItemCode = Me.OriginalItemCode
        Me.NewItemDesc = String.Empty
        Me.NewItemUpcCode = String.Empty

        Me.OriginalSKU = (grdRSTITEM1.ActiveRow.Cells("SKU").Value & String.Empty).ToString.Trim
        Me.OriginalUPC = (grdRSTITEM1.ActiveRow.Cells("UPC").Value & String.Empty).ToString.Trim

    End Sub

    Private Sub grdRSTITEM1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTITEM1.AfterRowUpdate

            Dim sql As String = "(UPC = '" & Me.OriginalUPC & "'  OR UPC = '') AND (SKU = '" & Me.OriginalSKU & "' OR SKU = '')"
        If Me.OriginalItemCode <> Me.NewItemCode Then
            For Each rowEDTRETLS In dst.Tables("EDTRETLS").Select(sql)
                rowEDTRETLS.Item("ITEM_CODE") = Me.NewItemCode
                rowEDTRETLS.Item("ITEM_DESC") = Me.NewItemDesc
                rowEDTRETLS.Item("UPC") = Me.NewItemUpcCode
            Next
        End If

        sql = "ITEM_CODE = '" & Me.NewItemCode & "'"
        For Each rowEDTRETLS In dst.Tables("EDTRETLS").Select(Sql)
            rowEDTRETLS.Item("IGNORE_ITEM_CODE") = Math.Abs(Val(Me.IGNORE_ITEM_CODE)).ToString.Trim
        Next

    End Sub

    Private Sub grdRSTITEM1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTITEM1.BeforeRowUpdate

        Me.NewItemCode = (grdRSTITEM1.ActiveRow.Cells("ITEM_CODE").Value & String.Empty).ToString.Trim
        Me.IGNORE_ITEM_CODE = (grdRSTITEM1.ActiveRow.Cells("IGNORE").Value & String.Empty).ToString.Trim
    End Sub

    Private Sub grdRSTITEM1_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTITEM1.CellChange
        If grdRSTITEM1.ActiveRow Is Nothing Then Exit Sub
        If grdRSTITEM1.ActiveCell.Column.Key <> "IGNORE" Then Exit Sub

        Try
            grdRSTITEM1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cell Change Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub grdRSTITEM1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTITEM1.ClickCellButton
        Select Case grdRSTITEM1.ActiveCell.Column.Key
            Case "ITEM_CODE"

                Try
                    'grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    MyBase.grdClickCellButton(grdRSTITEM1, String.Empty, False)

                    'grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                    Me.NewItemCode = grdRSTITEM1.ActiveRow.Cells("ITEM_CODE").Value & String.Empty

                    If Me.OriginalItemCode <> Me.NewItemCode Then
                        rowICTITEM1 = MyBase.LookUp("ICTITEM1", NewItemCode)
                        If rowICTITEM1 IsNot Nothing Then
                            grdRSTITEM1.ActiveRow.Cells("ITEM_UPC_CODE").Value = rowICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
                            Me.NewItemDesc = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                            Me.NewItemUpcCode = rowICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
                        End If
                        If Me.NewItemCode <> "" Then
                            e.Cell.Row.Cells("IGNORE").Value = "0"
                        End If
                    End If

                Catch ex As Exception
                    'grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    'MessageBox.Show(ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Finally
                    Try
                        'grdRSTITEM1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    Catch ex As Exception
                        'MessageBox.Show(ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try

                End Try


            Case Else
                Exit Sub

        End Select

        grdRSTITEM1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

    End Sub

#End Region

    Private Sub Item_Matching()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Matching SKUs to Items")

        For Each row As DataRow In dst.Tables("RSTITEM1").Select("ISNULL(ITEM_CODE,'') = ''")
            Dim ITEM_CODE As String = ""

            Dim SKU As String = Trim(row.Item("SKU") & "").ToUpper
            Dim rowICTITEM1 As DataRow = Nothing

            SKU = Replace(SKU, ".", "-")

            If SKU <> "" Then
                rowICTITEM1 = LookUp("ICTITEM1", SKU)
                If rowICTITEM1 IsNot Nothing Then
                    ITEM_CODE = SKU
                Else
                    If InStr(SKU, " ") <> 0 Then
                        rowICTITEM1 = LookUp("ICTITEM1", Split(SKU, " ")(0))
                        If rowICTITEM1 IsNot Nothing Then
                            ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                        End If
                    End If
                End If
            End If

            If ITEM_CODE = "" Then
                ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 " _
                & " where ITEM_CODE like " _
                & "'" & SKU & "%" & "'"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count = 1 Then
                    ITEM_CODE = tbl.Rows(0).Item(0)
                    rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                End If
            End If

            If ITEM_CODE = "" Then
                ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 " _
                & " where ITEM_CODE like " _
                & "'" & Mid$(SKU, 1, Len(SKU) - 1) & "%" & Mid(SKU, Len(SKU), 1) & "'"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count = 1 Then
                    ITEM_CODE = tbl.Rows(0).Item(0)
                    rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                End If
            End If


            If ITEM_CODE = "" Then
                ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 " _
                & " where ITEM_CODE like " _
                & "'" & Mid$(SKU, 1, Len(SKU) - 1) & "%" & "'"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count = 1 Then
                    ITEM_CODE = tbl.Rows(0).Item(0)
                    rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                End If
            End If

            If ITEM_CODE <> "" Then
                'If ITEM_CODE = "EBS66005WT" Then Stop
                row.Item("ITEM_CODE") = ITEM_CODE
                row.Item("IGNORE") = "0"
            End If

        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

End Class