Imports System.Windows.Forms
Imports ABSolution
Imports Infragistics.Win
Imports System.IO
Imports System.Data

Public Class EDF832O1

#Region "Class Variables"

    Private companyCode As String = String.Empty
    Private Const ediApplicationId As String = "SC"

    Private catalogueItemSql As String = String.Empty
    Private compareFields As String = String.Empty
    Private sqlEDT832O2 As String = String.Empty
    Private sqlICTITEM1 As String = String.Empty

    Private Const itemAdded As String = "02"
    Private Const itemDeleted As String = "03"
    Private Const itemChanged As String = "04"

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    ''' <summary>
    ''' ** Sets up required Data Tables and intializes form controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            companyCode = ASCMAIN1.DBS_COMPANY

            Get_PARM("ASTPARM1")
            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT832O1", "*", 0, True, String.Empty, 0, String.Empty)
            Create_TDA(.Tables.Add, "EDT832O2", "*")
            Create_TDA(.Tables.Add, "EDT832OX", "*")

            ASCMAIN1.sql = " SELECT * FROM EDTTRPM1 WHERE EDTTRPM1.EDI_DOC_NO = '832' AND EDI_STATUS = 'P'"
            Create_TDA(.Tables.Add, "EDTTRPM1", "**")
            Fill_Records("EDTTRPM1", String.Empty, True, ASCMAIN1.sql)

            .Relations.Add("EDT832O1_EDT832O2", _
                           New DataColumn() {dst.Tables("EDT832O1").Columns("COMPANY_CODE"), dst.Tables("EDT832O1").Columns("EDI_OUTBOUND_DOC_NO"), dst.Tables("EDT832O1").Columns("EDI_CATALOG_VERSION")}, _
                           New DataColumn() {dst.Tables("EDT832O2").Columns("COMPANY_CODE"), dst.Tables("EDT832O2").Columns("EDI_OUTBOUND_DOC_NO"), dst.Tables("EDT832O2").Columns("ITEM_SELECTION_CODE")})
        End With

        ASCMAIN1.sql = " Select 'COMPANY_CODE' COMPANY_CODE, 'EDI_OUR_ID' EDI_OUR_ID, 'EDI_TP_ID' EDI_TP_ID" _
            & ", ICTITEM1.ITEM_CODE, NVL(ITEM_DESC_CAT, SUBSTR(ICTITEM1.ITEM_DESC, 1, 20)) ITEM_DESC, ICTITEM1.ITEM_DESC ITEM_DESC_LONG" & vbCr _
            & ", NVL(ICTITEM1.CATALOG_SELECTION_CODE, ICTCOLL1.CATALOG_SELECTION_CODE) ITEM_SELECTION_CODE" & vbCr _
            & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCr _
            & ", NVL(ICTITEM1.NRF_COLOR_CODE, '000') ITEM_COLOR_CODE, ICTCOLRN.NRF_COLOR_DESC ITEM_COLOR_DESC" & vbCr _
            & ", NVL(ICTITEM1.NRF_SIZE_CODE, '00000') ITEM_SIZE_CODE, ICTSIZEN.NRF_SIZE_DESC ITEM_SIZE_DESC" & vbCr _
            & ", ICTITEM1.ITEM_SO_QTY_MIN, ICTITEM1.ITEM_SO_QTY_MULT" & vbCr _
            & ", ICTITEM1.ITEM_UPC_CODE ITEM_UPC, ICTITEM1.ITEM_EAN_CODE ITEM_EAN" & vbCr _
            & ", LPAD(NVL(ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE), 14, '0')  ITEM_GTIN" & vbCr _
            & ", DECODE(NVL(ITEM_STATUS, 'A'), 'I', NVL(NVL(ITEM_STATUS_DATE, LAST_DATE), SYSDATE), NULL) EDI_DISCONTINUE_DATE" & vbCr _
            & ", NULL HAZARD_CODE,  ICTSLCT1.CATALOG_SELECTION_DESC" & vbCr _
            & ", ICTITEM1.ITEM_WEIGHT, ICTITEM1.ITEM_UNIT_LENGTH, ICTITEM1.ITEM_UNIT_WIDTH, ICTITEM1.ITEM_UNIT_HEIGHT" & vbCr _
            & ", ICTBRAN1.BRAND_CODE, ICTBRAN1.BRAND_NAME, TATCNTRY.COUNTRY_CODE3 COUNTRY_NO, ICTCOLL1.COLLECTION_NAME, NVL(ICTITEM1.ITEM_AEROSOL, '0') ITEM_AEROSOL, ICTITEM1.COMMODITY_CODE, ICTITEM1.HMAT_CODE, ICTHMAT1.HMAT_DESC" & vbCr _
            & " from ICTITEM1, ICTCOLL1, ICTCOLRN, ICTSIZEN, ICTSLCT1, ICTBRAN1, TATCNTRY, ICTHMAT1" & vbCr _
            & " where ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCr _
            & " and NVL(ICTITEM1.NRF_COLOR_CODE, '000') = ICTCOLRN.NRF_COLOR_CODE (+)" & vbCr _
            & " and NVL(ICTITEM1.NRF_SIZE_CODE, '00000') = ICTSIZEN.NRF_SIZE_CODE (+)" & vbCr _
            & " and NVL(ICTITEM1.ITEM_HIDE_FROM_CAT, '0')  = '0' " & vbCr _
            & " and (ICTITEM1.ITEM_UPC_CODE IS NOT NULL or ICTITEM1.ITEM_EAN_CODE IS NOT NULL) " & vbCr _
            & " and NVL(ICTITEM1.CATALOG_SELECTION_CODE, ICTCOLL1.CATALOG_SELECTION_CODE) IS NOT NULL" & vbCr _
            & " and NVL(ICTITEM1.CATALOG_SELECTION_CODE, ICTCOLL1.CATALOG_SELECTION_CODE) = ICTSLCT1.CATALOG_SELECTION_CODE" _
            & " and ICTCOLL1.BRAND_CODE = ICTBRAN1.BRAND_CODE(+)" _
            & " and ICTITEM1.HMAT_CODE = ICTHMAT1.HMAT_CODE(+)" _
            & " and ICTITEM1.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE(+)"
        catalogueItemSql = ASCMAIN1.sql

        compareFields = " NVL(ITEM_UPC, '*'), ITEM_CODE, NVL(ITEM_COLOR_CODE, '*')" _
            & ", NVL(ITEM_SIZE_CODE, '*'), NVL(ITEM_SELECTION_CODE, '*')"

        ASCMAIN1.sql = "Select " & compareFields & vbCr _
           & "  from EDT832OX " & vbCr _
           & "  where COMPANY_CODE = 'COMPANY_CODE'" _
           & "  and EDI_OUR_ID = 'EDI_OUR_ID'" _
           & "  and EDI_TP_ID = 'EDI_TP_ID'"
        sqlEDT832O2 = ASCMAIN1.sql

        sqlICTITEM1 = "SELECT " & compareFields & " FROM (" & catalogueItemSql & ")"

        grdEDT832O1.DataSource = dst.Tables("EDT832O1")
        ASCMAIN1.Add_Value_List(grdEDT832O1, "EDI_PURPOSE_CODE", Nothing, New String() {":", "03:Delete", "02:Add", "04:Change"})

        Show_Filter(grdEDT832O1, True)
        Create_Summary(grdEDT832O1, "COMPANY_CODE", "Count")
        Create_Summary(grdEDT832O1, "EDI_DOC_LNO", "Count", grdEDT832O1.DisplayLayout.Bands(1).Key)

        If ASCMAIN1.CLIENT = "INT" Then
            With grdEDT832O1.DisplayLayout.Bands(1)
                .Columns("BRAND_CODE").Hidden = False
                .Columns("BRAND_CODE").Header.Caption = "Brand Code"

                .Columns("BRAND_NAME").Hidden = False
                .Columns("BRAND_NAME").Header.Caption = "Brand Name"

                .Columns("COUNTRY_NO").Hidden = False
                .Columns("COUNTRY_NO").Header.Caption = "Country No"

                .Columns("COLLECTION_NAME").Hidden = False
                .Columns("COLLECTION_NAME").Header.Caption = "Collection Name"

                .Columns("ITEM_AEROSOL").Hidden = False
                .Columns("ITEM_AEROSOL").Header.Caption = "Aerosol"
                ASCMAIN1.Add_Value_List(grdEDT832O1, "ITEM_AEROSOL", Nothing, New String() {":", "0:N", "1:Y"}, 1)

                .Columns("COMMODITY_CODE").Hidden = False
                .Columns("COMMODITY_CODE").Header.Caption = "Commodity Code"

                .Columns("HMAT_CODE").Hidden = False
                .Columns("HMAT_CODE").Header.Caption = "Hazmat Code"

                .Columns("HMAT_DESC").Hidden = False
                .Columns("HMAT_DESC").Header.Caption = "Hazmat Desc"

            End With
        End If


    End Sub

    ''' <summary>
    ''' Clear tables and controls based on the current state of the screen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRecord()

        EnforceConstraints(False)

        dst.Tables("EDTSYSIH").Clear()
        dst.Tables("EDT832O1").Clear()
        dst.Tables("EDT832O2").Clear()
        dst.Tables("EDT832OX").Clear()

        EnforceConstraints(True)
    End Sub

    ''' <summary>
    ''' Load up changes to go into 832 files
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadRecord()

        ASCMAIN1.Progress("Create 832 Records", String.Empty)

        Dim sql As String = String.Empty
        Dim tblData As DataTable = Nothing

        dst.Tables("EDT832OX").Rows.Clear()
        Try
            EnforceConstraints(False)

            For Each rowEDTTRPM1 As DataRow In dst.Tables("EDTTRPM1").Select()
                Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
                Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

                ASCMAIN1.sql = "Select * from EDT832OX where COMPANY_CODE = '" & companyCode & "' AND EDI_OUR_ID = '" & EDI_OUR_ID & "' AND EDI_TP_ID = '" & EDI_TP_ID & "'"
                Fill_Records("EDT832OX", String.Empty, False, ASCMAIN1.sql)

                ' Delete all differences
                ASCMAIN1.Progress("Deletions", "")
                sql = " SELECT X.*, ICTSLCT1.CATALOG_SELECTION_DESC"
                sql &= " FROM"
                sql &= " ("
                sql &= " SELECT  EDT832OX.* "
                sql &= " from EDT832OX "
                sql &= " where (" & compareFields & ")"
                sql &= " IN"
                sql &= " ("
                sql &= " SELECT " & compareFields & " from EDT832OX "
                sql &= " MINUS"
                sql &= " Select " & compareFields
                sql &= " from (" & catalogueItemSql & ")"
                sql &= " )"
                sql &= " ) X, ICTSLCT1"
                sql &= " WHERE X.ITEM_SELECTION_CODE = ICTSLCT1.CATALOG_SELECTION_CODE"
                sql &= " AND X.COMPANY_CODE = '" & companyCode & "'"
                sql &= " AND X.EDI_OUR_ID = '" & EDI_OUR_ID & "'"
                sql &= " AND X.EDI_TP_ID = '" & EDI_TP_ID & "'"

                sql &= " UNION "

                ' items in the catalog that are to be deleted from the catalog
                sql &= " SELECT X.*, ICTSLCT1.CATALOG_SELECTION_DESC"
                sql &= " FROM EDT832OX X, ICTSLCT1, ICTITEM1"
                sql &= " WHERE X.ITEM_SELECTION_CODE = ICTSLCT1.CATALOG_SELECTION_CODE"
                sql &= " AND X.COMPANY_CODE = '" & companyCode & "'"
                sql &= " AND X.EDI_OUR_ID = '" & EDI_OUR_ID & "'"
                sql &= " AND X.EDI_TP_ID = '" & EDI_TP_ID & "'"
                sql &= " AND X.ITEM_CODE = ICTITEM1.ITEM_CODE"
                sql &= " AND NVL(ICTITEM1.ITEM_HIDE_FROM_CAT, '0')  = '1'"

                tblData = ASCDATA1.GetDataTable(sql)
                For Each col As DataColumn In tblData.Columns
                    col.ReadOnly = False
                Next

                ' Set data in deleted rows to data in EDT832OX
                For Each row As DataRow In tblData.Select()
                    ASCMAIN1.sql = "COMPANY_CODE = '" & companyCode & "' AND ITEM_CODE = '" & row.Item("ITEM_CODE") & "' and EDI_OUR_ID = '" & row.Item("EDI_OUR_ID") & "' AND EDI_TP_ID = '" & row.Item("EDI_TP_ID") & "'"
                    If dst.Tables("EDT832OX").Select(ASCMAIN1.sql).Length > 0 Then
                        Dim rowEDT832OX As DataRow = dst.Tables("EDT832OX").Select(ASCMAIN1.sql)(0)
                        For Each col As DataColumn In rowEDT832OX.Table.Columns
                            If tblData.Columns.Contains(col.ColumnName) AndAlso col.ColumnName <> "ITEM_CODE" Then
                                row.Item(col.ColumnName) = rowEDT832OX.Item(col.ColumnName)
                            End If
                        Next
                    End If

                Next
                Create832Entries(tblData, itemDeleted, rowEDTTRPM1)

                ' Additions - Item Codes not in the Calalogue
                ASCMAIN1.Progress("Additions", "")
                sql = catalogueItemSql.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " and NVL(ICTITEM1.ITEM_STATUS, 'A') = 'A'"
                sql &= " and ICTITEM1.ITEM_CODE in (" & vbCr
                sql &= " Select ITEM_CODE from (" & vbCr
                sql &= sqlICTITEM1.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " Minus" & vbCr
                sql &= sqlEDT832O2.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " ))"
                tblData = ASCDATA1.GetDataTable(sql)
                Create832Entries(tblData, itemAdded, rowEDTTRPM1)

                ASCMAIN1.Progress("Changes", "")
                sql = catalogueItemSql.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " and NVL(ICTITEM1.ITEM_STATUS, 'A') = 'A'"
                sql &= " and ICTITEM1.ITEM_CODE in (" & vbCr
                sql &= " Select ITEM_CODE from (" & vbCr
                sql &= sqlICTITEM1.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " INTERSECT " & vbCr
                sql &= sqlEDT832O2.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " ))"
                tblData = ASCDATA1.GetDataTable(sql)


                Dim nothingChanged As Boolean = True
                Dim Fields As String = "ITEM_DESC,ITEM_RETAIL_PRICE,ITEM_SO_QTY_MIN,ITEM_SO_QTY_MULT,EDI_DISCONTINUE_DATE,ITEM_GTIN" _
                                       & ",BRAND_CODE,BRAND_NAME,COUNTRY_NO,COLLECTION_NAME,ITEM_AEROSOL,COMMODITY_CODE,HMAT_CODE,HMAT_DESC" _
                                       & ",ITEM_WEIGHT,ITEM_UNIT_LENGTH,ITEM_UNIT_WIDTH,ITEM_UNIT_HEIGHT"
                For Each rowChanges As DataRow In tblData.Select("")
                    Dim ITEM_CODE As String = rowChanges.Item("ITEM_CODE") & String.Empty
                    ASCMAIN1.Progress("-", ITEM_CODE)
                    For Each rowEDT832OX As DataRow In dst.Tables("EDT832OX").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                        nothingChanged = True
                        For Each field As String In Fields.Split(",")
                            field = field.Trim
                            If field.Length = 0 Then Continue For
                            If rowEDT832OX.Item(field) & String.Empty <> rowChanges.Item(field) & String.Empty Then
                                nothingChanged = False
                                Exit For
                            End If
                        Next
                    Next
                    ' If nothing changed then delete the record
                    If nothingChanged Then rowChanges.Delete()
                Next
                tblData.AcceptChanges()
                Create832Entries(tblData, itemChanged, rowEDTTRPM1)
            Next

            EnforceConstraints(True)
        Catch ex As Exception
            ClearRecord()
            MessageBox.Show(ex.Message)
        Finally
            grdEDT832O1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        End Try

    End Sub

    ''' <summary>
    ''' Sets up screen based on the form modality, state and type of processing
    ''' </summary>
    ''' <param name="tf"></param>
    ''' <remarks></remarks>
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_Description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode

            If dst.Tables("EDT832O1").Rows.Count = 0 Then
                .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
            Else
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End If

            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
        End With

        If ScreenMode Then
            grdEDT832O1.Visible = True
        Else
            ClearRecord()
            grdEDT832O1.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' Validates data when a user selects a menu option
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim sql As String = String.Empty

        Select Case eItemKey

            Case "Generate"
                If dst.Tables("EDTTRPM1").Rows.Count = 0 Then
                    EMsg &= "There are no 832 entries."
                End If

            Case "Update"
                If dst.Tables("EDT832O1").Rows.Count = 0 Then
                    EMsg = "No records to update."
                End If

            Case "Cancel"

        End Select

        If EMsg <> String.Empty Then
            MessageBox.Show(EMsg, "Cannot Proceed", MessageBoxButtons.OK)
        Else
            Call Proceed(eItemKey)
        End If

    End Sub

    ''' <summary>
    ''' When the user selects a menu option perform the action
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Generate"
                Me.LoadRecord()
                Me.Mode_Settings(True)

            Case "Update"
                Me.UpdateRecord()
                Me.Mode_Settings(False)

            Case "Cancel"
                Me.Mode_Settings(False)

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    ''' <summary>
    ''' Updates data based on the current state of the screen and the type of processing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateRecord()

        Try
            MyBase.BeginTrans()

            Dim rowEDT832OX As DataRow = Nothing

            For Each rowEDTSYSIH As DataRow In dst.Tables("EDTSYSIH").Select("", "EDI_OUTBOUND_DOC_NO")
                Dim COMPANY_CODE As String = rowEDTSYSIH.Item("COMPANY_CODE")
                Dim EDI_OUTBOUND_DOC_NO As String = rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO")
                Dim EDI_OUR_ID As String = rowEDTSYSIH.Item("EDI_OUR_ID")
                Dim EDI_TP_ID As String = rowEDTSYSIH.Item("EDI_TP_ID")

                For Each rowEDT832O1 As DataRow In dst.Tables("EDT832O1").Select("COMPANY_CODE = '" & COMPANY_CODE & "' AND EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'")

                    If rowEDT832O1.Item("EDI_PURPOSE_CODE") = itemDeleted Then
                        Continue For
                    End If

                    For Each rowEDT832O2 As DataRow In dst.Tables("EDT832O2").Select("COMPANY_CODE = '" & COMPANY_CODE & "' AND EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'")

                        Dim ITEM_CODE As String = rowEDT832O2.Item("ITEM_CODE")
                        rowEDT832OX = dst.Tables("EDT832OX").Rows.Find(New Object() {COMPANY_CODE, EDI_OUR_ID, EDI_TP_ID, ITEM_CODE})

                        If rowEDT832OX Is Nothing Then
                            rowEDT832OX = dst.Tables("EDT832OX").NewRow
                            rowEDT832OX.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT832OX.Item("EDI_OUR_ID") = EDI_OUR_ID
                            rowEDT832OX.Item("EDI_TP_ID") = EDI_TP_ID
                            rowEDT832OX.Item("ITEM_CODE") = ITEM_CODE
                            dst.Tables("EDT832OX").Rows.Add(rowEDT832OX)
                        End If

                        rowEDT832OX.Item("ITEM_DESC") = rowEDT832O2.Item("ITEM_DESC")
                        rowEDT832OX.Item("ITEM_SELECTION_CODE") = rowEDT832O2.Item("ITEM_SELECTION_CODE")
                        rowEDT832OX.Item("ITEM_RETAIL_PRICE") = rowEDT832O2.Item("ITEM_RETAIL_PRICE")
                        rowEDT832OX.Item("ITEM_COLOR_CODE") = rowEDT832O2.Item("ITEM_COLOR_CODE")
                        rowEDT832OX.Item("ITEM_COLOR_DESC") = rowEDT832O2.Item("ITEM_COLOR_DESC")
                        rowEDT832OX.Item("ITEM_SIZE_CODE") = rowEDT832O2.Item("ITEM_SIZE_CODE")
                        rowEDT832OX.Item("ITEM_SIZE_DESC") = rowEDT832O2.Item("ITEM_SIZE_DESC")
                        rowEDT832OX.Item("ITEM_SO_QTY_MIN") = rowEDT832O2.Item("ITEM_SO_QTY_MIN")
                        rowEDT832OX.Item("ITEM_SO_QTY_MULT") = rowEDT832O2.Item("ITEM_SO_QTY_MULT")
                        rowEDT832OX.Item("ITEM_UPC") = rowEDT832O2.Item("ITEM_UPC")
                        rowEDT832OX.Item("ITEM_EAN") = rowEDT832O2.Item("ITEM_EAN")
                        rowEDT832OX.Item("ITEM_GTIN") = rowEDT832O2.Item("ITEM_GTIN")
                        rowEDT832OX.Item("EDI_DISCONTINUE_DATE") = rowEDT832O2.Item("EDI_DISCONTINUE_DATE")
                        rowEDT832OX.Item("HAZARD_CODE") = rowEDT832O2.Item("HAZARD_CODE")

                        rowEDT832OX.Item("ITEM_WEIGHT") = rowEDT832O2.Item("ITEM_WEIGHT")
                        rowEDT832OX.Item("ITEM_UNIT_LENGTH") = rowEDT832O2.Item("ITEM_UNIT_LENGTH")
                        rowEDT832OX.Item("ITEM_UNIT_WIDTH") = rowEDT832O2.Item("ITEM_UNIT_WIDTH")
                        rowEDT832OX.Item("ITEM_UNIT_HEIGHT") = rowEDT832O2.Item("ITEM_UNIT_HEIGHT")
                        rowEDT832OX.Item("ITEM_DESC_LONG") = rowEDT832O2.Item("ITEM_DESC_LONG")

                        If ASCMAIN1.CLIENT = "INT" Then
                            rowEDT832OX.Item("BRAND_CODE") = rowEDT832O2.Item("BRAND_CODE")
                            rowEDT832OX.Item("BRAND_NAME") = rowEDT832O2.Item("BRAND_NAME")
                            rowEDT832OX.Item("COUNTRY_NO") = rowEDT832O2.Item("COUNTRY_NO")
                            rowEDT832OX.Item("COLLECTION_NAME") = rowEDT832O2.Item("COLLECTION_NAME")
                            rowEDT832OX.Item("ITEM_AEROSOL") = Val(rowEDT832O2.Item("ITEM_AEROSOL") & String.Empty)
                            rowEDT832OX.Item("COMMODITY_CODE") = rowEDT832O2.Item("COMMODITY_CODE")
                            rowEDT832OX.Item("HMAT_CODE") = rowEDT832O2.Item("HMAT_CODE")
                            rowEDT832OX.Item("HMAT_DESC") = rowEDT832O2.Item("HMAT_DESC")
                        End If
                    Next
                Next
            Next

            MyBase.Update_Record_TDA("EDTSYSIH")
            MyBase.Update_Record_TDA("EDT832O1")
            MyBase.Update_Record_TDA("EDT832O2")
            MyBase.Update_Record_TDA("EDT832OX")

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT ITEM_CODE FROM ICTITEM1 WHERE NVL(ICTITEM1.ITEM_STATUS, 'A') = 'I')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT ITEM_CODE FROM ICTITEM1 WHERE NVL(ICTITEM1.ITEM_HIDE_FROM_CAT, '0') = '1')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_UPC IS NULL AND ITEM_EAN IS NULL AND ITEM_GTIN IS NULL"
            ASCDATA1.ExecuteSQL()

            MyBase.CommitTrans("Update Successful")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)

        End Try

    End Sub

#End Region

#Region "Private Subs / Functions"

    ''' <summary>
    ''' Create 832O1 and 832O2 records for the specifiec Purpose Type
    ''' </summary>
    ''' <param name="tbl832Data"></param>
    ''' <param name="EDI_TRANS_PURP_CODE"></param>
    ''' <remarks></remarks>
    Private Sub Create832Entries(ByRef tbl832Data As DataTable, ByVal EDI_TRANS_PURP_CODE As String, ByRef rowEDTTRPM1 As DataRow)

        If tbl832Data Is Nothing OrElse tbl832Data.Rows.Count = 0 Then
            Exit Sub
        End If

        For Each rowSelectionCode As DataRow In ASCDATA1.SelectDistinct(tbl832Data, New String() {"ITEM_SELECTION_CODE"}).Rows
            Dim ITEM_SELECTION_CODE As String = rowSelectionCode.Item("ITEM_SELECTION_CODE") & String.Empty
            Dim rowHeader As DataRow = tbl832Data.Select("ITEM_SELECTION_CODE = '" & ITEM_SELECTION_CODE & "'")(0)

            ' Moved from up above
            Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

            Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
            rowEDTSYSIH.Item("COMPANY_CODE") = companyCode
            rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDTSYSIH.Item("EDI_APPLICATION_ID") = ediApplicationId
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "1"
            rowEDTSYSIH.Item("EDI_OUR_ID") = rowHeader.Item("EDI_OUR_ID")
            rowEDTSYSIH.Item("EDI_TP_ID") = rowHeader.Item("EDI_TP_ID")
            rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
            rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

            Dim rowEDT832O1 As DataRow = dst.Tables("EDT832O1").NewRow
            rowEDT832O1.Item("COMPANY_CODE") = companyCode
            rowEDT832O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDT832O1.Item("EDI_CATALOG_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO")
            rowEDT832O1.Item("EDI_CATALOG_VERSION") = ITEM_SELECTION_CODE

            Dim rowICTSLCT1 As DataRow = LookUp("ICTSLCT1", ITEM_SELECTION_CODE)
            If rowICTSLCT1 IsNot Nothing Then
                rowEDT832O1.Item("EDI_CATALOG_DESC") = rowICTSLCT1.Item("CATALOG_SELECTION_DESC")
            Else
                rowEDT832O1.Item("EDI_CATALOG_DESC") = rowHeader.Item("CATALOG_SELECTION_DESC")
            End If

            rowEDT832O1.Item("EDI_PURPOSE_CODE") = EDI_TRANS_PURP_CODE
            rowEDT832O1.Item("EDI_CATALOG_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            rowEDT832O1.Item("EDI_NAME") = ROWs("ASTPARM1").Item("AS_PARM_INST_NAME")
            rowEDT832O1.Item("INIT_DATE") = DateTime.Now
            rowEDT832O1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowEDT832O1.Item("LAST_DATE") = DateTime.Now
            rowEDT832O1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            dst.Tables("EDT832O1").Rows.Add(rowEDT832O1)

            Dim EDI_DOC_LNO As Int16 = 0
            Dim rowEDT832O2 As DataRow = Nothing
            For Each rowDetail As DataRow In tbl832Data.Select("ITEM_SELECTION_CODE = '" & ITEM_SELECTION_CODE & "'")
                ASCMAIN1.Progress("-", rowDetail.Item("ITEM_CODE") & "")

                rowEDT832O2 = dst.Tables("EDT832O2").NewRow
                rowEDT832O2.Item("COMPANY_CODE") = companyCode
                rowEDT832O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                EDI_DOC_LNO += 1
                rowEDT832O2.Item("EDI_DOC_LNO") = EDI_DOC_LNO
                rowEDT832O2.Item("ITEM_CODE") = rowDetail.Item("ITEM_CODE")
                rowEDT832O2.Item("ITEM_DESC") = rowDetail.Item("ITEM_DESC")
                rowEDT832O2.Item("ITEM_SELECTION_CODE") = rowDetail.Item("ITEM_SELECTION_CODE")
                rowEDT832O2.Item("ITEM_RETAIL_PRICE") = rowDetail.Item("ITEM_RETAIL_PRICE")
                rowEDT832O2.Item("ITEM_COLOR_CODE") = rowDetail.Item("ITEM_COLOR_CODE")
                rowEDT832O2.Item("ITEM_COLOR_DESC") = rowDetail.Item("ITEM_COLOR_DESC")
                rowEDT832O2.Item("ITEM_SIZE_CODE") = rowDetail.Item("ITEM_SIZE_CODE")
                rowEDT832O2.Item("ITEM_SIZE_DESC") = rowDetail.Item("ITEM_SIZE_DESC")
                rowEDT832O2.Item("ITEM_SO_QTY_MIN") = rowDetail.Item("ITEM_SO_QTY_MIN")
                rowEDT832O2.Item("ITEM_SO_QTY_MULT") = rowDetail.Item("ITEM_SO_QTY_MULT")
                rowEDT832O2.Item("ITEM_UPC") = rowDetail.Item("ITEM_UPC")
                rowEDT832O2.Item("ITEM_EAN") = rowDetail.Item("ITEM_EAN")
                rowEDT832O2.Item("ITEM_GTIN") = rowDetail.Item("ITEM_GTIN")
                rowEDT832O2.Item("EDI_DISCONTINUE_DATE") = rowDetail.Item("EDI_DISCONTINUE_DATE")
                rowEDT832O2.Item("HAZARD_CODE") = rowDetail.Item("HAZARD_CODE")

                rowEDT832O2.Item("ITEM_WEIGHT") = rowDetail.Item("ITEM_WEIGHT")
                rowEDT832O2.Item("ITEM_UNIT_LENGTH") = rowDetail.Item("ITEM_UNIT_LENGTH")
                rowEDT832O2.Item("ITEM_UNIT_WIDTH") = rowDetail.Item("ITEM_UNIT_WIDTH")
                rowEDT832O2.Item("ITEM_UNIT_HEIGHT") = rowDetail.Item("ITEM_UNIT_HEIGHT")
                rowEDT832O2.Item("ITEM_DESC_LONG") = rowDetail.Item("ITEM_DESC_LONG")

                If ASCMAIN1.CLIENT = "INT" Then
                    rowEDT832O2.Item("BRAND_CODE") = rowDetail.Item("BRAND_CODE")
                    rowEDT832O2.Item("BRAND_NAME") = rowDetail.Item("BRAND_NAME")
                    rowEDT832O2.Item("COUNTRY_NO") = rowDetail.Item("COUNTRY_NO")
                    rowEDT832O2.Item("COLLECTION_NAME") = rowDetail.Item("COLLECTION_NAME")
                    rowEDT832O2.Item("ITEM_AEROSOL") = Val(rowDetail.Item("ITEM_AEROSOL") & String.Empty)
                    rowEDT832O2.Item("COMMODITY_CODE") = rowDetail.Item("COMMODITY_CODE")
                    rowEDT832O2.Item("HMAT_CODE") = rowDetail.Item("HMAT_CODE")
                    rowEDT832O2.Item("HMAT_DESC") = rowDetail.Item("HMAT_DESC")
                End If
                dst.Tables("EDT832O2").Rows.Add(rowEDT832O2)
            Next
        Next

    End Sub

#End Region

#Region "Form Controls"

#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class