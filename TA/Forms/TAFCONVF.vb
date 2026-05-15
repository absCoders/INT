Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class TAFCONVF
    Dim grdX2Os() As UltraWinGrid.UltraGrid
    WithEvents grdE As UltraWinGrid.UltraGrid
    WithEvents grdC As UltraWinGrid.UltraGrid
    WithEvents grdTATIPIT1 As UltraWinGrid.UltraGrid
    Dim tblE As DataTable
    Dim tblC As DataTable
    Dim excelFileNum As Integer
    Dim loaded_once As Boolean = False
    Dim datafile_folder As String

    Dim dt_errors As New DataTable

    Dim LAST_CHANGE_COLUMN_NAME As String
    Dim LAST_CHANGE_CELL_VALUE As String
    Dim COPY_VALUE_clipboard As String
    Dim COLUMN_NAME_clipboard As String
    Dim COLUMN_CAPTION_clipboard As String
    Dim IPSA_IMPORT_NO As String

    Dim IPSA2INT As New Dictionary(Of String, String)
    Dim INT2IPSA As New Dictionary(Of String, String)
    Dim sqlChanges As String
    Dim TATIPITC As String

    Dim App_Red As New Infragistics.Win.Appearance
    Dim App_LtGrn As New Infragistics.Win.Appearance


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'TATIPCNT IPSA_CONTENTS
        'TATIPIC1 IPSA_ICTITEM1

        'TATIPHIM IPSA_HIERARCHY_MAP
        'TATIPMGM IPSA_MERCH_GROUP_MAP

        INT2IPSA.Add("TATIPHI1", "IPSA_HIERARCHY")
        INT2IPSA.Add("TATIPIT1", "IPSA_ITEM_MASTER")
        INT2IPSA.Add("TATIPMG1", "IPSA_MERCH_GROUP")
        INT2IPSA.Add("TATIPPK1", "IPSA_PACK")
        INT2IPSA.Add("TATIPPKG", "IPSA_PACKING")
        INT2IPSA.Add("TATIPPR1", "IPSA_PRICE_LIST")


        For Each TABLE_NAME As String In INT2IPSA.Keys
            IPSA2INT.Add(INT2IPSA(TABLE_NAME), TABLE_NAME)
        Next

        With dst
            With .Tables.Add("TATCONVX")
                .Columns.Add("FILENAME")
                .Columns.Add("FILE_DESCRIPTION")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("ROWCOUNT", GetType(System.Int64))
                .Columns.Add("INIT_DATE", GetType(System.DateTime))
                .Columns.Add("LAST_DATE", GetType(System.DateTime))
                .Columns.Add("SECS", GetType(System.Int64))

                .PrimaryKey = New DataColumn() { .Columns("FILENAME")}

                .Rows.Add(New Object() {"HIERARCHY", "Brand & Collection", IPSA2INT("IPSA_HIERARCHY")})
                .Rows.Add(New Object() {"ITEM_MASTER", "Item Master", IPSA2INT("IPSA_ITEM_MASTER")})
                .Rows.Add(New Object() {"PACKING", "Pack Qtys Item (Case, Inner, Pallet)", IPSA2INT("IPSA_PACKING")})
                .Rows.Add(New Object() {"PRICE_LIST", "Price List by Item (IPLB Costs)", IPSA2INT("IPSA_PRICE_LIST")})
            End With

            Create_TDA(.Tables.Add, "TATIPIC1", "*")
            dst.Tables("TATIPIC1").Columns.Add("IPSA_ACTION")
            dst.Tables("TATIPIC1").Columns("ITEM_LOT_CONTROL").DefaultValue = "0"


            Create_TDA(.Tables.Add, "ICTITEM1", "*")

            Create_TDA(.Tables.Add, "ICTSIZE1", "*", 0)
        End With

        grdTATCONVX.DataSource = dst.Tables("TATCONVX")

        Create_Summary(grdTATCONVX, "FILENAME", "Count")

        If ASCMAIN1.Running_in_VS Then
            datafile_folder = "C:\Users\wjz\Desktop\Interparfums\IPSA\PROD\FROM_IPSA\ITEM"
        Else
            If ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST" Then
                datafile_folder = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\IPSA\TEST\FROM_IPSA\ITEM"
            Else
                datafile_folder = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\IPSA\PROD\FROM_IPSA\ITEM"
            End If
        End If

        App_Red.ForeColor = Color.Red
        App_LtGrn.BackColor = Color.LightGreen


        Get_Changes()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Datafiles"
                If loaded_once Then
                    EMsg &= vbCr & "You may load only once.  Exit to the Menu and then re-execute."
                End If

                If EMsg = "" Then

                    Dim FILES_shortname As New List(Of String)
                    Dim FILES As New Dictionary(Of String, String)
                    For Each FILE As String In My.Computer.FileSystem.GetFiles(datafile_folder)
                        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
                        Dim FILENAME As String = Split(fi.Name, ".")(0).ToUpper
                        FILES_shortname.Add(FILENAME)
                        FILES.Add(FILENAME, FILE)
                    Next

                    ReDim grdX2Os(dst.Tables("TATCONVX").Rows.Count - 1)
                    loaded_once = True
                    Dim SHEET_NO As Integer = 0

                    ASCMAIN1.Progress("Now Loading", "")

                    If tabX2O.Tabs.Count > 1 Then
                        For T As Int16 = tabX2O.Tabs.Count - 1 To 1
                            tabX2O.Tabs.Remove(tabX2O.Tabs(T))
                        Next
                    End If

                    For Each rowTATCONVX As DataRow In dst.Tables("TATCONVX").Select("", "")

                        Dim FILENAME As String = rowTATCONVX.Item("FILENAME")
                        ASCMAIN1.Progress("-", FILENAME)

                        SHEET_NO += 1
                        Dim TABLE_NAME As String = IPSA2INT("IPSA_" & FILENAME)
                        ASCMAIN1.Progress("-", TABLE_NAME)

                        ASCMAIN1.sql = "Select * from " & TABLE_NAME

                        If dst.Tables.Contains(TABLE_NAME) Then
                            dst.Tables(TABLE_NAME).Rows.Clear()
                        Else
                            Create_TDA(dst.Tables.Add(TABLE_NAME), TABLE_NAME, "**", 0, True)
                        End If

                        Dim grd As UltraWinGrid.UltraGrid

                        If SHEET_NO = 1 Then
                            grd = grdX2O
                            grd.DataSource = Nothing
                        Else
                            tabX2O.Tabs.Add(New UltraWinTabControl.UltraTab)
                            grd = New UltraWinGrid.UltraGrid
                            grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
                            grd.Parent = tabX2O.Tabs(SHEET_NO - 1).TabPage
                            grd.Visible = True
                            grd.Dock = DockStyle.Fill
                        End If

                        tabX2O.Tabs(SHEET_NO - 1).Text = TABLE_NAME
                        tabX2O.Tabs(SHEET_NO - 1).Tag = TABLE_NAME

                        grdX2Os(SHEET_NO - 1) = grd

                        If FILES_shortname.Contains(FILENAME) Then
                            Load_Workbook(FILES(FILENAME), TABLE_NAME, SHEET_NO)
                        Else
                            Fill_Records(TABLE_NAME)
                        End If

                        grd.DataSource = dst.Tables(TABLE_NAME)
                        ASCMAIN1.grdInitializeLayout(grd)
                        grd.DisplayLayout.GroupByBox.Hidden = False
                        Show_Filter(grd, True)

                        For I As Int16 = 0 To dst.Tables(TABLE_NAME).Columns.Count - 1
                            grd.DisplayLayout.Bands(0).Columns(I).Header.Appearance.TextHAlign = HAlign.Center
                        Next
                        With grd.DisplayLayout.Override
                            .RowSelectors = DefaultableBoolean.True
                            .RowSelectorNumberStyle = UltraWinGrid.RowSelectorNumberStyle.VisibleIndex
                            .RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
                            .RowSelectorAppearance.TextHAlign = HAlign.Center
                        End With

                        grd.DisplayLayout.CaptionVisible = DefaultableBoolean.True
                        grd.DisplayLayout.CaptionAppearance.TextHAlign = HAlign.Left
                        grd.Text = FILENAME

                    Next

                    ASCMAIN1.Progress("", "")

                    Map_IPSA_to_IPLB()

                    Get_Changes()
                    Create_TDA(dst.Tables.Add("TATIPITC"), TATIPITC, "*", 0)
                    Create_TDA(dst.Tables.Add, TATIPITC, "*", 0)
                    'dst.Tables("TATIPITC").Columns.Add("ITEM_ALLOW_HALF_PACK")
                    'dst.Tables(TATIPITC).Columns.Add("ITEM_ALLOW_HALF_PACK")
                    '  Fill_Records("TATIPITC")
                    ASCMAIN1.sql = "Select * from " & TATIPITC
                    Setup_grd("Changes", "TATIPITC")
                End If

            Case "Edit Changes"
                If Not ASCMAIN1.Logical_Lock("T", "ICTITEM1") Then
                    Exit Sub
                End If
                If Not ASCMAIN1.Logical_Lock("F", "DPFPLAN1") Then
                    Exit Sub
                End If

            Case "Update"
                ' SR-6363 - New Item Master creation by Marketing 
                If EntryMode = "N" Then
                    If tblE.Columns.Contains("COMPLETED") AndAlso tblE.Columns.Contains("APPROVED") Then
                        Dim lstItems As New List(Of String)
                        For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U' AND (ISNULL(COMPLETED, '0') <> '1' OR ISNULL(APPROVED, '0') <> '1')")
                            lstItems.Add(row.Item("ITEM_CODE") & String.Empty)
                        Next

                        If lstItems.Count > 0 Then
                            EMsg = vbCr & "The following items are set to be Updated; however, they are not Completed and/or Approved." & Environment.NewLine
                            EMsg &= String.Join(Environment.NewLine, lstItems.ToArray)
                            Exit Select
                        End If
                    End If
                End If

                Dim sMsg As String = ""
                Dim setITEM_CONTAINS_ALCOHOL As Boolean = False
                Dim setITEM_LOT_CONTROL As Boolean = False
                Dim setCOMMODITY_CODE As Boolean = False
                For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U'")
                    Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(row.Item("ITEM_SHELF_LIFE_YRS") & "")
                    Dim ITEM_LOT_CONTROL As String = row.Item("ITEM_LOT_CONTROL") & ""
                    Dim ITEM_CONTAINS_ALCOHOL As String = row.Item("ITEM_CONTAINS_ALCOHOL") & ""
                    Dim COMMODITY_CODE As String = row.Item("COMMODITY_CODE") & ""
                    Dim PROD_CODE As String = row.Item("PROD_CODE") & ""
                    If ITEM_SHELF_LIFE_YRS <> 0 Then
                        If ITEM_LOT_CONTROL <> "1" Then
                            If Not setITEM_LOT_CONTROL Then
                                setITEM_LOT_CONTROL = True
                                sMsg &= vbCrLf & vbTab & "- without Lot Control"
                            End If
                        End If

                        If PROD_CODE = "SPGS" Or ITEM_SHELF_LIFE_YRS = 5 Then
                            If ITEM_CONTAINS_ALCOHOL <> "1" Then
                                If Not setITEM_CONTAINS_ALCOHOL Then
                                    setITEM_CONTAINS_ALCOHOL = True
                                    sMsg &= vbCrLf & vbTab & "- that are not marked as Hazardous"
                                End If
                            End If
                        End If


                    End If

                    If (PROD_CODE = "SPGS" Or ITEM_SHELF_LIFE_YRS = 5) Or ITEM_CONTAINS_ALCOHOL = "1" Then
                        If COMMODITY_CODE = "" Then
                            setCOMMODITY_CODE = True
                            sMsg &= vbCrLf & vbTab & "- that are Hazardous without a Commodity Code (FRAGRANCE)"
                        End If
                    End If
                Next

                If sMsg <> "" Then
                    If MsgBox("There are items with Non-0 Shelf Life" _
                              & vbCrLf & sMsg _
                              & vbCrLf & vbCrLf & "Do you want to Auto-Correct these Items?", MsgBoxStyle.YesNo, "There are Items queued for Update with Error Conditions") = MsgBoxResult.Yes Then
                        For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U'")
                            Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(row.Item("ITEM_SHELF_LIFE_YRS") & "")
                            Dim PROD_CODE As String = row.Item("PROD_CODE") & ""

                            If ITEM_SHELF_LIFE_YRS <> 0 Then
                                If row.Item("ITEM_LOT_CONTROL") & "" <> "1" Then row.Item("ITEM_LOT_CONTROL") = "1"

                                If PROD_CODE = "SPGS" Or ITEM_SHELF_LIFE_YRS = 5 Then
                                    If row.Item("ITEM_CONTAINS_ALCOHOL") & "" <> "1" Then row.Item("ITEM_CONTAINS_ALCOHOL") = "1"
                                End If

                            End If

                            Dim COMMODITY_CODE As String = row.Item("COMMODITY_CODE") & ""
                            Dim ITEM_CONTAINS_ALCOHOL As String = row.Item("ITEM_CONTAINS_ALCOHOL") & ""
                            If (PROD_CODE = "SPGS" Or ITEM_SHELF_LIFE_YRS = 5) Or ITEM_CONTAINS_ALCOHOL = "1" Then
                                If COMMODITY_CODE = "" Then
                                    row.Item("COMMODITY_CODE") = "3303009000"
                                End If
                            End If
                        Next
                    End If
                End If



                Dim dtx As New DataTable
                With dtx.Columns
                    .Add("COLUMN_NAME")
                    .Add("TABLE_NAME")
                    .Add("COLUMN_CAPTION")
                End With
                dtx.Rows.Add(New String() {"PROD_CODE", "ICTPROD1", "Product Code"})
                dtx.Rows.Add(New String() {"COLLECTION_CODE", "ICTCOLL1", "Collection Code"})
                dtx.Rows.Add(New String() {"ITEM_CLASS_CODE", "ICTCLAS1", "Class Code"})
                dtx.Rows.Add(New String() {"ITEM_CATGY_CODE", "ICTCATG1", "Item Category Code"})
                dtx.Rows.Add(New String() {"COST_CATGY_CODE", "ICTCOST1", "Cost Category Code"})
                For Each dtrow As DataRow In dtx.Select
                    Dim COLUMN_NAME As String = dtrow.Item("COLUMN_NAME")
                    Dim TABLE_NAME As String = dtrow.Item("TABLE_NAME")
                    Dim COLUMN_CAPTION As String = dtrow.Item("COLUMN_CAPTION")
                    For Each row As DataRow In ASCDATA1.SelectDistinct(tblE.Select("IPSA_ACTION = 'U'"), COLUMN_NAME).Select
                        Dim CODE_VALUE As String = row.Item(0) & ""
                        If LookUp(TABLE_NAME, CODE_VALUE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value for " & COLUMN_CAPTION & " (" & CODE_VALUE & ")"
                        End If
                    Next
                Next

                ASCMAIN1.sql = "Select * from ICTCOST1 where ITEM_SNU_CODE = 'S'"
                Dim Ss As New List(Of String)
                Dim SQL_in As String = ""
                Dim SQL_not_in As String = ""
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim COST_CATGY_CODE As String = row.Item("COST_CATGY_CODE")
                    Ss.Add(COST_CATGY_CODE)
                    SQL_in &= " or COST_CATGY_CODE = '" & COST_CATGY_CODE & "'"
                    SQL_not_in &= " and COST_CATGY_CODE <> '" & COST_CATGY_CODE & "'"
                Next
                SQL_in = " and (" & Mid(SQL_in, 5) & ")"

                For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U' and ISNULL(ITEM_RETAIL_PRICE,0) = 0" & SQL_in)
                    EMsg &= vbCr & "Cannot have $0.00 Retail in Saleable Cost Category" & " - " & row.Item("ITEM_CODE")
                    Exit For
                Next
                For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U' and ISNULL(ITEM_RETAIL_PRICE,0) <> 0" & SQL_not_in)
                    EMsg &= vbCr & "Cannot have non-$0.00 Retail in Non_Saleable Cost Category" & " - " & row.Item("ITEM_CODE")
                    Exit For
                Next

                For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U'")

                    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                    Dim COST_CATGY_CODE As String = row.Item("COST_CATGY_CODE")
                    Dim rowICTCOST1 As DataRow = LookUp("ICTCOST1", COST_CATGY_CODE)
                    Dim ITEM_SNU_CODE As String = ""
                    If rowICTCOST1 IsNot Nothing Then
                        ITEM_SNU_CODE = rowICTCOST1.Item("ITEM_SNU_CODE") & ""
                    End If
                    Dim ITEM_EAN_CODE As String = row.Item("ITEM_EAN_CODE") & ""
                    If ITEM_EAN_CODE = "" Or ITEM_EAN_CODE.Contains(" ") Then
                        If ITEM_EAN_CODE = "" And ITEM_SNU_CODE = "U" Then
                            ' it is ok for components to not have an EAN
                        Else
                            EMsg &= vbCr & "Item EAN Code is mandatory, and may not contain blanks" & " - " & ITEM_CODE
                        End If
                    Else
                        If ITEM_EAN_CODE.Length <> 13 Or Not (ITEM_EAN_CODE.StartsWith("3386460") Or ITEM_EAN_CODE.StartsWith("3386461")) Then
                            EMsg &= vbCr & $"invalid Item EAN Code ({ITEM_EAN_CODE})" & " - " & ITEM_CODE
                        End If
                    End If

                    If ITEM_EAN_CODE <> "" Then
                        Dim rowEAN As DataRow = ASCDATA1.GetDataRow("Select ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE = :PARM1 and ITEM_CODE <> :PARM2", "VV", New Object() {ITEM_EAN_CODE, ITEM_CODE})
                        If rowEAN IsNot Nothing Then
                            EMsg &= vbCr & "EAN " & ITEM_EAN_CODE & " already used for Item " & rowEAN.Item("ITEM_CODE")
                        End If
                    End If


                    Dim ITEM_ALT_SORT As String = row.Item("ITEM_ALT_SORT") & ""
                    If ITEM_ALT_SORT = "" Or ITEM_ALT_SORT.Contains(" ") Then
                        EMsg &= vbCr & "3PL Item Code is mandatory, and may not contain blanks" & " - " & ITEM_CODE
                    Else
                        '/^[a-z0-9]+$/i
                        '/^([a-zA-Z0-9 _-]+)$/
                        'Dim rx As String = "^([a-zA-Z0-9_-]+)$" ' Allow Upper/Lower case, numbers, underscore and dash
                        Dim rx As String = "[^A-Z0-9]" ' Allow Upper case, numbers

                        Dim r As New System.Text.RegularExpressions.Regex(rx)
                        If r.IsMatch(ITEM_ALT_SORT) Then
                            EMsg &= vbCr & "3PL Item Code has Special Characters which are not allowed" & " - " & ITEM_CODE
                        End If
                    End If

                    'If EMsg = "" Then
                    Dim rowITEM_ALT_SORT As DataRow = ASCDATA1.GetDataRow("Select * from ICTITEM1 where ITEM_ALT_SORT = :PARM1", "V", ITEM_ALT_SORT)
                    If rowITEM_ALT_SORT IsNot Nothing Then
                        EMsg &= vbCr & "3PL Item Code " & ITEM_ALT_SORT & " is already in use with " & rowITEM_ALT_SORT.Item("ITEM_CODE") & " - " & ITEM_CODE
                    Else
                        If tblE.Select("ITEM_ALT_SORT = '" & ITEM_ALT_SORT & "'").Length > 1 Then
                            EMsg &= vbCr & "3PL Item Code " & ITEM_ALT_SORT & " appears to be assigned to more than 1 item" & " - " & ITEM_CODE
                        End If
                    End If
                    'End If

                    Dim ITEM_DESC As String = row.Item("ITEM_DESC") & ""
                    If ITEM_DESC = "" Then
                        EMsg &= vbCr & "Cannot Update an Item with No Description" & " - " & ITEM_CODE
                    Else
                        If ITEM_DESC.Length > 30 Then
                            EMsg &= vbCr & "Item Description cannot be more than 30 Characters (3PL restriction)" & " - " & ITEM_CODE
                        End If
                        'Dim rx As String = "[^a-zA-Z0-9 .,_$-]" ' Allow Upper/Lower case, numbers, space, dot, comma, underscore and dash
                        ' Dim rx As String = "[^a-zA-Z0-9 .,$-]" ' Allow Upper/Lower case, numbers, space, dot, comma, and dash
                        Dim rx As String = "[^a-zA-Z0-9 .$]" ' Allow Upper/Lower case, numbers, space, dot

                        Dim r As New System.Text.RegularExpressions.Regex(rx)
                        If r.IsMatch(ITEM_DESC) Then
                            '  If   r.IsMatch(Replace(Replace(ITEM_DESC, " ", ""), ".", "")) Then
                            EMsg &= vbCr & "Item Description has Special Characters which are not allowed (3PL restriction)" & " - " & ITEM_CODE
                        End If
                    End If

                    For Each DCOL As DataColumn In tblE.Columns
                        If DCOL.DataType.ToString = "System.String" Or DCOL.DataType.ToString = "System.DateTime" Then
                        Else
                            If Val(row.Item(DCOL.ColumnName) & "") < 0 Then
                                EMsg &= vbCr & "Cannot Have Negative Value for " & grdE.DisplayLayout.Bands(0).Columns(DCOL.ColumnName).Header.Caption & " - " & ITEM_CODE
                            End If
                        End If
                    Next

                    Dim ITEM_COST_STD As Decimal = Val(row.Item("ITEM_COST_STD") & "")
                    'If ITEM_COST_STD <= 0 Then
                    '    EMsg &= vbCr & "Cannot Have Zero or Negative Value for Cost" & " - " & ITEM_CODE
                    'End If
                    ' Permit 0 cost per LM 11/13
                    If ITEM_COST_STD < 0 Then ' ITEM_COST_STD <= 0 Then
                        EMsg &= vbCr & "Cannot Have Negative Value for Cost" & " - " & ITEM_CODE
                    End If

                    Dim ITEM_VALUE As Decimal = Val(row.Item("ITEM_VALUE") & "")
                    If ITEM_VALUE < 0 Then
                        EMsg &= vbCr & "Cannot Have Negative Value for Value" & " - " & ITEM_CODE
                    End If

                    Dim ITEM_SO_QTY_MULT As Int64 = Val(row.Item("ITEM_SO_QTY_MULT") & "")
                    Dim ITEM_SO_QTY_MIN As Int64 = Val(row.Item("ITEM_SO_QTY_MIN") & "")
                    Dim ITEM_STD_PACK_SLS As Int64 = Val(row.Item("ITEM_STD_PACK_SLS") & "")
                    Dim ITEM_ALLOW_HALF_PACK As String = row.Item("ITEM_ALLOW_HALF_PACK") & ""

                    If ITEM_SO_QTY_MULT = 0 Then
                        EMsg &= vbCr & "SO Multiple may not be 0" & " - " & ITEM_CODE
                        ITEM_SO_QTY_MULT = 1
                    End If

                    If ITEM_SO_QTY_MIN = 0 Then
                        EMsg &= vbCr & "SO Minimum may not be 0" & " - " & ITEM_CODE
                        ITEM_SO_QTY_MIN = 1
                    End If

                    If ITEM_STD_PACK_SLS = 0 Then ITEM_STD_PACK_SLS = 1

                    ' REM UPDATE DGJ
                    'If  ITEM_STD_PACK_SLS = 0 And PROD_CODE = "SPGS" Then
                    '    ITEM_STD_PACK_SLS = rowTATIPITC.Item("CARTON_PACK_QTY")
                    'End If

                    ' LET 

                    If ITEM_SO_QTY_MIN <> 1 And ITEM_SO_QTY_MIN Mod ITEM_SO_QTY_MULT <> 0 And Not (ITEM_SO_QTY_MIN = ITEM_SO_QTY_MULT / 2 And ITEM_ALLOW_HALF_PACK = "1") Then EMsg &= vbCr & "SO Minimum is not congruous with SO Multiple" & " - " & ITEM_CODE
                    If ITEM_SO_QTY_MIN <> 1 And ITEM_SO_QTY_MIN Mod ITEM_STD_PACK_SLS <> 0 Then EMsg &= vbCr & "SO Minimum is not congruous with Standard (Inner) Pack" & " - " & ITEM_CODE
                    If ITEM_SO_QTY_MULT <> 1 And ITEM_SO_QTY_MULT Mod ITEM_STD_PACK_SLS <> 0 Then EMsg &= vbCr & "SO Multiple is not congruous with Standard (Inner) Pack" & " - " & ITEM_CODE

                    Dim ITEM_SHELF_LIFE_YRS As Int32 = Val(row.Item("ITEM_SHELF_LIFE_YRS") & "")
                    Dim ITEM_LOT_CONTROL As String = row.Item("ITEM_LOT_CONTROL") & ""
                    Dim ITEM_CONTAINS_ALCOHOL As String = row.Item("ITEM_CONTAINS_ALCOHOL") & ""
                    If ITEM_LOT_CONTROL = "1" And ITEM_SHELF_LIFE_YRS = 0 Then
                        EMsg &= vbCr & "Cannot have Lot Control if Shelf Life is 0" & " - " & ITEM_CODE
                    End If

                    If ITEM_CONTAINS_ALCOHOL = "1" And ITEM_LOT_CONTROL <> "1" Then
                        EMsg &= vbCr & "Cannot set Hazardous (Contains Alcohol) without Lot Control" & " - " & ITEM_CODE
                    End If

                    If ITEM_LOT_CONTROL <> "1" And ITEM_SHELF_LIFE_YRS <> 0 Then
                        EMsg &= vbCr & "Cannot have a non-0 Shelf Life without Lot Control" & " - " & ITEM_CODE
                    End If


                    Dim ITEM_CLASS_CODE As String = row.Item("ITEM_CLASS_CODE") & ""
                    If ITEM_CONTAINS_ALCOHOL <> "1" And ITEM_SHELF_LIFE_YRS <> 0 Then
                        If ITEM_CLASS_CODE = "511" And ITEM_SHELF_LIFE_YRS = 3 Then
                            ' lbm 01/22/22 Correct.  511 is not in my list.  I only gave you the classes which are indeed hazardous.  
                            ' i did Not give you the exclusion as there are many other classes.  
                            ' But If it will be programmed to 511 only if there Is a 3 shelf life than ok.
                        Else
                            EMsg &= vbCr & "Cannot have a non-0 Shelf Life without Marking as Hazardous" & " - " & ITEM_CODE
                        End If
                    End If


                    Dim COMMODITY_CODE As String = row.Item("COMMODITY_CODE") & ""
                    If COMMODITY_CODE <> "" Then
                        If LookUp("ICTCOMM1", COMMODITY_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Commodity Code" & " - " & ITEM_CODE
                        End If
                    End If

                    If ITEM_ALLOW_HALF_PACK = "1" And (ITEM_SO_QTY_MULT = 0 Or ITEM_SO_QTY_MULT Mod 2 <> 0) Then
                        EMsg &= vbCr & "Cannot have Half Pack unless Order Multiple is >0 and an Even Number"
                    End If
                    ' no half pack unless inner pack > 0 and even

                    If EMsg.Length > 1000 Then Exit For
                Next
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Datafiles"
                Load_Record()
                Mode_Settings(True)

            'Case "Edit"
            '    EntryMode = "E"
            '    Mode_Settings(True)

            Case "Edit New Items"
                EntryMode = "N"
                Mode_Settings(True)

            Case "Edit Changes"
                EntryMode = "C"
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done", "Cancel"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load Datafiles").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                '  .Items("Edit").Settings.Enabled = iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("Edit New Items").Settings.Enabled = iScreenMode
                .Items("Edit Changes").Settings.Enabled = iScreenMode

                .Items("Load Datafiles").Visible = Not ScreenMode
                .Items("Done").Visible = Not (ScreenMode And (EntryMode = "C" Or EntryMode = "N"))
                .Items("Update").Visible = ScreenMode And (EntryMode = "C" Or EntryMode = "N")
                .Items("Cancel").Visible = ScreenMode And (EntryMode = "C" Or EntryMode = "N")
                .Items("Edit New Items").Visible = ScreenMode And (EntryMode <> "N" And EntryMode <> "C")
                .Items("Edit Changes").Visible = ScreenMode And (EntryMode <> "N" And EntryMode <> "C")
            End With

            .Groups("Data Controls").Visible = ScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdTATCONVX.Visible = Not ScreenMode

        If ScreenMode Then
            If EntryMode = "C" Then
                grdE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                tabX2O.Tabs("New Items").Visible = False
                tabX2O.Tabs("Changes").Selected = True
                '   grdc.TabStop.
            ElseIf EntryMode = "N" Then
                grdE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                tabX2O.Tabs("Changes").Visible = False
                tabX2O.Tabs("New Items").Selected = True
            Else
                grdE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If
        Else
            Clear_Record()
        End If
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading")
        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.sql = Get_sql_IPSA_Item_Master()
        Setup_grd("IPSA Item Master")

        ASCMAIN1.sql = Get_sql_IPLB_Item_Master()
        Setup_grd("IPLB Item Master")

        'ASCMAIN1.sql = "Select * from TATIPIC1 where IPSA_STATUS_CODE = 'I'"
        'Setup_grd("New Items Ignored")

        ASCMAIN1.sql = "Select * from TATIPIC1 where IPSA_STATUS_CODE in ('P','I')"
        Setup_grd("New Items", "TATIPIC1")
        tblE.AcceptChanges()
        Sort_grdColumns(grdE, "ITEM_CODE")
        grdE.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdE.DisplayLayout.Bands(0).ColumnFilters("IPSA_STATUS_CODE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.StartsWith, "Pending")

        Sort_grdColumns(grdC, "ITEM_CODE")
        grdC.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()


        tabX2O.SelectedTab = tabX2O.Tabs(tabX2O.Tabs.Count - 1)


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        'For Each TABLE_NAME As String In New String() {"TATCONVY"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next

        'If dst.Tables("TATCONVX").Rows.Count = 0 Then
        '    ASCMAIN1.sql = "Select * from TATCONVX"
        '    Fill_Records("TATCONVX", "", , ASCMAIN1.sql)
        '    Sort_grdColumns(grdTATCONVX, "ODOBNM")
        'End If

        Application.DoEvents()
    End Sub

    Sub Update_Record()

        Fill_Records("ICTSIZE1")

        BeginTrans()

        dst.Tables("TATIPIC1").Rows.Clear()
        For Each row As DataRow In tblE.Select("", "", DataViewRowState.ModifiedCurrent)
            row.Item("IPSA_STATUS_CODE") = row.Item("IPSA_ACTION")
            row.Item("LAST_DATE") = DATETIME_STAMP
            row.Item("LAST_OPER") = ASCMAIN1.USER_ID
            Dim row2 As DataRow = dst.Tables("TATIPIC1").NewRow
            row2.ItemArray = row.ItemArray
            dst.Tables("TATIPIC1").Rows.Add(row2)
            row2.AcceptChanges()
            row2.SetModified()
        Next
        Update_Record_TDA("TATIPIC1")

        dst.Tables("ICTITEM1").Rows.Clear()
        For Each row As DataRow In tblE.Select("IPSA_ACTION = 'U'")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")

            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").NewRow
            With rowICTITEM1
                For Each DCOL As DataColumn In tblE.Columns
                    If dst.Tables("ICTITEM1").Columns.Contains(DCOL.ColumnName) Then
                        .Item(DCOL.ColumnName) = row.Item(DCOL.ColumnName)
                    End If
                    'If DCOL.ColumnName = "IPSA_ACTION" Or DCOL.ColumnName = "IPSA_STATUS_CODE" Then
                    'Else
                    '    .Item(DCOL.ColumnName) = row.Item(DCOL.ColumnName)
                    'End If
                Next

                If .Item("ITEM_CONTAINS_ALCOHOL") = "1" Then .Item("HMAT_CODE") = "1266"
                ' 12/21/2016 email Maria/Lauren ICTITEM1.HMAT_CODE to 1266 if ICTITEM1.ITEM_CONTAINS_ALCOHOL is set to 1

                Dim rowICTCOST1 As DataRow = LookUp("ICTCOST1", .Item("COST_CATGY_CODE"))
                .Item("ITEM_SNU_CODE") = rowICTCOST1.Item("ITEM_SNU_CODE")
                .Item("ITEM_UOM") = "EA"
                .Item("ITEM_STATUS") = "A"
                .Item("ITEM_STATUS_DATE") = DATETIME_STAMP.ToString("MM/dd/yyyy")
                .Item("ITEM_TYPE_CODE") = "FG"
                Dim rowICTTYPE1 As DataRow = LookUp("ICTTYPE1", .Item("ITEM_TYPE_CODE"))
                If rowICTTYPE1 IsNot Nothing Then
                    .Item("ITEM_PLAN_WASTE_PCT") = rowICTTYPE1.Item("ITEM_WASTE_PCT")
                End If

                Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", .Item("COLLECTION_CODE"))
                If rowICTCOLL1 IsNot Nothing Then
                    Dim BRAND_CODE As String = rowICTCOLL1.Item("BRAND_CODE")
                    Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", BRAND_CODE)
                    .Item("SALES_DIVISION_CODE") = rowICTBRAN1.Item("SALES_DIVISION_CODE")
                End If

                .Item("ITEM_CATGY_CODE") = "E"
                .Item("ITEM_PLAN_MAKE_BUY") = "B"

                .Item("ITEM_COST_MAKE_BUY") = "B"
                .Item("ITEM_COST_CURR_CODE") = "USD"
                .Item("ITEM_COST_FRT_CLASS") = "Z"
                .Item("ITEM_COST_STD_FUTURE") = 0
                .Item("VEND_CODE") = "IPSA"
                .Item("NRF_SIZE_CODE") = "00000"
                .Item("NRF_COLOR_CODE") = "000"
                .Item("ITEM_YYYYPP_CUR_COST") = ASCMAIN1.CYP

                Dim SIZE_DESC As String = rowICTITEM1.Item("SIZE_CODE") & ""
                If SIZE_DESC <> "" Then
                    ASCMAIN1.sql = "Select * from ICTSIZEN where REPLACE(NRF_SIZE_DESC,' ','') = :PARM1"
                    Dim rowICTSIZEN As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {SIZE_DESC})

                    If rowICTSIZEN IsNot Nothing Then
                        .Item("NRF_SIZE_CODE") = rowICTSIZEN.Item("NRF_SIZE_CODE")
                        .Item("NRF_SIZE_CODE") = "00000" ' LBM email 02/06 requesting that all NRF SIZE codes go to 00000 (MACYS COMPLAINT)
                    End If
                End If

                Dim PROD_CODE As String = .Item("PROD_CODE") & ""
                If PROD_CODE <> "" Then
                    Dim rowICTPROD1 As DataRow = LookUp("ICTPROD1", PROD_CODE)
                    If rowICTPROD1 IsNot Nothing Then
                        Dim COST_CATGY_CODE As String = rowICTPROD1.Item("COST_CATGY_CODE") & ""
                        Dim PROD_MAX_POS As Decimal = Val(rowICTPROD1.Item("PROD_MAX_POS") & "")
                        Dim PROD_MIN_POS As Decimal = Val(rowICTPROD1.Item("PROD_MIN_POS") & "")
                        Dim PROD_MIN_DAYS_SUPPLY As Decimal = Val(rowICTPROD1.Item("PROD_MIN_DAYS_SUPPLY") & "")
                        If COST_CATGY_CODE = "N" AndAlso PROD_MIN_DAYS_SUPPLY >= 0 And PROD_MAX_POS > 0 And PROD_MIN_POS > 0 And PROD_MAX_POS > PROD_MIN_POS Then
                            .Item("ITEM_POS_MAX") = PROD_MAX_POS
                            .Item("ITEM_POS_MIN") = PROD_MIN_POS
                            .Item("ITEM_MIN_DAYS_SUPPLY") = PROD_MIN_DAYS_SUPPLY
                        End If
                    End If
                End If


                .Item("ITEM_DESC_CAT") = Mid(.Item("ITEM_DESC") & "", 1, 20)

                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
            End With
            dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)

            Write_Audit_Trail(rowICTITEM1, "N")

            Dim SIZE_CODE As String = rowICTITEM1.Item("SIZE_CODE") & ""
            If SIZE_CODE <> "" Then
                Dim rowICTSIZE1 As DataRow = dst.Tables("ICTSIZE1").Rows.Find(SIZE_CODE)
                If rowICTSIZE1 Is Nothing Then
                    dst.Tables("ICTSIZE1").Rows.Add(New String() {SIZE_CODE, SIZE_CODE})
                End If
            End If
        Next

        Update_Record_TDA("ICTITEM1")
        Update_Record_TDA("ICTSIZE1")

        For Each row As DataRow In dst.Tables("ICTITEM1").Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")

            ASCMAIN1.sql = "" _
                & "Insert into ICTCOSTC" & vbCrLf _
                & "(ITEM_CODE, INIT_DATE, INIT_OPER, ITEM_COST_TOTAL, ITEM_COST_VCOST, ITEM_COST_VCURR" & vbCrLf _
                & ", ITEM_COST_FRT_CLASS, ITEM_COST_MAKE_BUY" & vbCrLf _
                & ", COLLECTION_CODE, ITEM_CLASS_CODE, ITEM_COST_CURR_CODE, COST_CATGY_CODE, ITEM_CATGY_CODE)" & vbCrLf _
                & "Select ITEM_CODE, SYSDATE, 'conv', ITEM_COST_STD, ITEM_COST_STD, ITEM_COST_STD" _
                & ", ITEM_COST_FRT_CLASS, ITEM_COST_MAKE_BUY" _
                & ", COLLECTION_CODE, ITEM_CLASS_CODE, 'USD', COST_CATGY_CODE, ITEM_CATGY_CODE" _
                & " from ICTITEM1 where ITEM_CODE = '" & ITEM_CODE & "'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, ICTCOSTC.* from ICTCOSTC where ITEM_CODE = '" & ITEM_CODE & "'"
            ASCDATA1.ExecuteSQL("Insert into ICTCOSTA " & ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Insert into ICTCOSTH " & ASCMAIN1.sql)
        Next

        dst.Tables("TATIPITC").Rows.Clear()
        ASCDATA1.ExecuteSQL("Delete from " & TATIPITC)
        For Each row As DataRow In tblC.Select("IPSA_ACTION = 'U'")
            Dim rowTATIPITC As DataRow = dst.Tables("TATIPITC").NewRow
            For Each DC As DataColumn In tblC.Columns
                If DC.ColumnName <> "IPSA_ACTION" Then
                    rowTATIPITC.Item(DC.ColumnName) = row.Item(DC.ColumnName)
                End If
            Next
            dst.Tables("TATIPITC").Rows.Add(rowTATIPITC)
        Next
        Update_Record_TDA("TATIPITC")

        dst.Tables("ICTITEM1").Rows.Clear()

        For Each rowTATIPITC As DataRow In dst.Tables("TATIPITC").Select("")
            Dim ITEM_CODE As String = rowTATIPITC.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = Fill_Record("ICTITEM1", ITEM_CODE, False, False)
            For Each C As String In New String() {
                "ITEM_UNIT_LENGTH", "ITEM_UNIT_WIDTH", "ITEM_UNIT_HEIGHT", "ITEM_WEIGHT",
                "CARTON_PACK_QTY", "ITEM_CASE_LENGTH", "ITEM_CASE_WIDTH", "ITEM_CASE_HEIGHT", "CASE_WEIGHT_GRS", "ITEM_CASE_CUBAGE",
                "ITEM_PALLET_QTY", "ITEM_PALLET_LENGTH", "ITEM_PALLET_WIDTH", "ITEM_PALLET_HEIGHT", "ITEM_PALLET_WEIGHT", "ITEM_PALLET_CUBAGE"}
                rowICTITEM1.Item(C) = rowTATIPITC.Item(C)
            Next
            rowICTITEM1.Item("ITEM_STD_PACK_PUR") = rowTATIPITC.Item("CARTON_PACK_QTY")
            'If Val(rowICTITEM1.Item("ITEM_PO_QTY_MIN") & "") = 0 Then rowICTITEM1.Item("ITEM_PO_QTY_MIN") = rowTATIPITC.Item("CARTON_PACK_QTY")
            'rowICTITEM1.Item("ITEM_PO_QTY_MULT") = rowTATIPITC.Item("CARTON_PACK_QTY")

            ' REM UPDATE DGJ
            'If rowICTITEM1.Item("ITEM_STD_PACK_SLS") = 0 And PROD_CODE = "SPGS" Then
            '    rowICTITEM1.Item("ITEM_STD_PACK_SLS") = rowTATIPITC.Item("CARTON_PACK_QTY")
            'End If

            'If dst.Tables("TATIPITC").Columns.Contains("ITEM_ALLOW_HALF_PACK") Then
            '    rowICTITEM1.Item("ITEM_ALLOW_HALF_PACK") = rowTATIPITC.Item("ITEM_ALLOW_HALF_PACK")
            'End If


            Write_Audit_Trail(rowICTITEM1, "E")

        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from " & TATIPITC & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & "  Update ICTITEM1 Set" & vbCrLf _
            & "   ITEM_UNIT_LENGTH = R1.ITEM_UNIT_LENGTH" & vbCrLf _
            & ",  ITEM_UNIT_WIDTH = R1.ITEM_UNIT_WIDTH" & vbCrLf _
            & ",  ITEM_UNIT_HEIGHT = R1.ITEM_UNIT_HEIGHT" & vbCrLf _
            & ",  ITEM_WEIGHT = R1.ITEM_WEIGHT" & vbCrLf _
            & ",  CARTON_PACK_QTY = R1.CARTON_PACK_QTY" & vbCrLf _
            & ",  ITEM_CASE_LENGTH = R1.ITEM_CASE_LENGTH" & vbCrLf _
            & ",  ITEM_CASE_WIDTH = R1.ITEM_CASE_WIDTH" & vbCrLf _
            & ",  ITEM_CASE_HEIGHT = R1.ITEM_CASE_HEIGHT" & vbCrLf _
            & ",  CASE_WEIGHT_GRS = R1.CASE_WEIGHT_GRS" & vbCrLf _
            & ",  ITEM_CASE_CUBAGE = R1.ITEM_CASE_CUBAGE" & vbCrLf _
            & ",  ITEM_PALLET_QTY = R1.ITEM_PALLET_QTY" & vbCrLf _
            & ",  ITEM_PALLET_LENGTH = R1.ITEM_PALLET_LENGTH" & vbCrLf _
            & ",  ITEM_PALLET_WIDTH = R1.ITEM_PALLET_WIDTH" & vbCrLf _
            & ",  ITEM_PALLET_HEIGHT = R1.ITEM_PALLET_HEIGHT" & vbCrLf _
            & ",  ITEM_PALLET_WEIGHT = R1.ITEM_PALLET_WEIGHT" & vbCrLf _
            & ",  ITEM_PALLET_CUBAGE = R1.ITEM_PALLET_CUBAGE" & vbCrLf _
            & ",  ITEM_STD_PACK_PUR = R1.CARTON_PACK_QTY" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & " End Loop; End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub

#End Region

    Sub Format_grd(GRD As UltraWinGrid.UltraGrid, TABLE_NAME As String)

        Select Case TABLE_NAME
            Case "TATIPIC1"
                grdE = GRD
                grdE.Name = "grdE"
                GRDs.Add("E", grdE)
                ASCMAIN1.grdInitializeLayout(GRD)

                Load_Popup_Menu(grdE, "SSSBBBB", "Show GroupBox", "Show Filter", "Show Pins", "Clear Column", "Paste Value to All", "Copy Value to Clipboard", "Paste Value to Selected")

                With GRD.DisplayLayout.Bands(0)

                    Dim P As Integer = -1
                    grdfmt(.Columns("ITEM_CODE"), P, 100, "Item Code", , , True, True)
                    grdfmt(.Columns("ITEM_DESC"), P, 170, "Description", , , True)
                    grdfmt(.Columns("IPSA_STATUS_CODE"), P, 70, "Status", , , True, True)
                    grdfmt(.Columns("IPSA_ACTION"), P, 80, "Action", , , True)

                    ' SR-6363 - New Item Master creation by Marketing 
                    grdfmt(.Columns("COMPLETED"), P, 80, "Completed", Color.LightPink,, True, True)
                    grdfmt(.Columns("APPROVED"), P, 80, "Approved", Color.LightBlue,, True, True)
                    grdfmt(.Columns("MARKETING_NOTES"), P, 170, "Marketing Notes", ,, True, True)

                    grdfmt(.Columns("ITEM_ALT_SORT"), P, 80, "3PL Item")
                    grdfmt(.Columns("ITEM_EAN_CODE"), P, 120, "EAN")
                    grdfmt(.Columns("ITEM_VALUE"), P, 70, "Value", Color.LightGreen, "#.00")
                    grdfmt(.Columns("ITEM_RETAIL_PRICE"), P, 70, "Retail", Color.LightGreen, "#.00")
                    grdfmt(.Columns("ITEM_COST_STD"), P, 70, "Cost", Color.LightGreen, "#.00")
                    grdfmt(.Columns("PROD_CODE"), P, 70, "Prod", Color.LightBlue, , , , "PROD_CODE")
                    grdfmt(.Columns("COLLECTION_CODE"), P, 100, "Coll", Color.LightBlue, , , , "COLLECTION_CODE")
                    grdfmt(.Columns("SIZE_CODE"), P, 80, "Size", Color.LightBlue, , , , "SIZE_CODE")
                    grdfmt(.Columns("ITEM_CLASS_CODE"), P, 60, "Class", Color.LightBlue, , , , "ITEM_CLASS_CODE")
                    grdfmt(.Columns("COST_CATGY_CODE"), P, 60, "CostCat", Color.LightBlue, , , True, "COST_CATGY_CODE")
                    grdfmt(.Columns("ITEM_CATGY_CODE"), P, 60, "Catgy", Color.LightBlue, , , , "ITEM_CATGY_CODE")
                    grdfmt(.Columns("SEASON_CODE"), P, 90, "Season", Color.LightBlue, , , , "SEASON_CODE")
                    grdfmt(.Columns("LAUNCH_DATE"), P, 100, "Launch", Color.LightBlue, "MM/dd/yy")
                    grdfmt(.Columns("ITEM_BASIC_PROMO"), P, 40, "BP", Color.LightBlue)
                    grdfmt(.Columns("ITEM_WEIGHT"), P, 70, "Wgt", Color.Gold, "#.00")
                    grdfmt(.Columns("ITEM_UNIT_LENGTH"), P, 70, "Len", Color.Gold, "#.00")
                    grdfmt(.Columns("ITEM_UNIT_WIDTH"), P, 70, "Wid", Color.Gold, "#.00")
                    grdfmt(.Columns("ITEM_UNIT_HEIGHT"), P, 70, "Hgt", Color.Gold, "#.00")
                    grdfmt(.Columns("ITEM_STD_PACK_PUR"), P, 70, "Inr-Pur", Color.Orange, "#,##0")
                    grdfmt(.Columns("ITEM_STD_PACK_SLS"), P, 70, "Factor", Color.Orange, "#,##0")
                    grdfmt(.Columns("ITEM_PO_QTY_MIN"), P, 70, "PO Min", Color.Orange, "#,##0")
                    grdfmt(.Columns("ITEM_PO_QTY_MULT"), P, 70, "PO Mult", Color.Orange, "#,##0")
                    grdfmt(.Columns("ITEM_SO_QTY_MIN"), P, 70, "SO Min", Color.Orange, "#,##0")
                    grdfmt(.Columns("ITEM_SO_QTY_MULT"), P, 70, "SO Mult", Color.Orange, "#,##0")
                    grdfmt(.Columns("CARTON_PACK_QTY"), P, 70, "Ctn Pack", Color.Orange, "#,##0")
                    grdfmt(.Columns("CASE_WEIGHT_GRS"), P, 70, "Wgt-Ctn", Color.Orange, "#.00")
                    grdfmt(.Columns("ITEM_CASE_CUBAGE"), P, 70, "Cub-Ctn", Color.Orange, "#.00")
                    grdfmt(.Columns("ITEM_CASE_LENGTH"), P, 70, "Len-Ctn", Color.Orange, "#.00")
                    grdfmt(.Columns("ITEM_CASE_WIDTH"), P, 70, "Wid-Ctn", Color.Orange, "#.00")
                    grdfmt(.Columns("ITEM_CASE_HEIGHT"), P, 70, "Hgt-Ctn", Color.Orange, "#.00")
                    grdfmt(.Columns("ITEM_PALLET_QTY"), P, 70, "Qty-Pal", Color.LightPink, "#,##0")
                    grdfmt(.Columns("ITEM_PALLET_WEIGHT"), P, 70, "Wgt-Pal", Color.LightPink, "#.00")
                    grdfmt(.Columns("ITEM_PALLET_CUBAGE"), P, 70, "Cub-Pal", Color.LightPink, "#.00")
                    grdfmt(.Columns("ITEM_PALLET_LENGTH"), P, 70, "Len-Pal", Color.LightPink, "#.00")
                    grdfmt(.Columns("ITEM_PALLET_WIDTH"), P, 70, "Wid-Pal", Color.LightPink, "#.00")
                    grdfmt(.Columns("ITEM_PALLET_HEIGHT"), P, 70, "Hgt-Pal", Color.LightPink, "#.00")
                    grdfmt(.Columns("ITEM_DESC2"), P, 40, "Desc2")
                    grdfmt(.Columns("COUNTRY_CODE"), P, 40, "Country", , , , True)
                    grdfmt(.Columns("COMMODITY_CODE"), P, 60, "Commodity", Color.Yellow)
                    grdfmt(.Columns("ITEM_CONTAINS_ALCOHOL"), P, 40, "Haz", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_CRITICAL_TO_SHIP"), P, 40, "Crit", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_LOT_CONTROL"), P, 40, "LotCtl", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_WEIGHT_CHECK"), P, 40, "WgtChk", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_APPR_1ST_REC"), P, 40, "ApprRec", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_CRITICAL_TO_SHIP"), P, 40, "Crit", Color.Yellow, "x")
                    grdfmt(.Columns("ITEM_SHELF_LIFE_YRS"), P, 40, "Life", Color.Violet, "#,##0")
                    grdfmt(.Columns("INIT_DATE"), P, 80, "Created", Color.Violet, "MM/dd/yy", , True)
                    grdfmt(.Columns("INIT_OPER"), P, 40, "By", Color.Violet, , , True)
                    grdfmt(.Columns("LAST_DATE"), P, 80, "Changed", Color.Violet, "MM/dd/yy", , True)
                    grdfmt(.Columns("LAST_OPER"), P, 40, "By", Color.Violet, , , True)

                    grdfmt(.Columns("ITEM_ALLOW_HALF_PACK"), P, 60, "Half?", Color.Orange, , True)
                    .Columns("ITEM_ALLOW_HALF_PACK").Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Columns("ITEM_ALLOW_HALF_PACK").Hidden = True

                    ' SR-6363 - New Item Master creation by Marketing 
                    grdfmt(.Columns("COMPLETED_BY"), P, 100, "Completed By", Color.LightPink,,, True)
                    grdfmt(.Columns("COMPLETED_ON"), P, 100, "Completed On", Color.LightPink, "n",, True)

                    grdfmt(.Columns("APPROVED_BY"), P, 100, "Approved By", Color.LightBlue,,, True)
                    grdfmt(.Columns("APPROVED_ON"), P, 100, "Approved On", Color.LightBlue, "n",, True)

                    .Columns("COMPLETED").Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Columns("APPROVED").Style = UltraWinGrid.ColumnStyle.CheckBox
                End With

                ASCMAIN1.Add_Value_List(GRD, "IPSA_STATUS_CODE", Nothing, New String() {":", "P:Pending", "I:Ignore"})
                ASCMAIN1.Add_Value_List(GRD, "IPSA_ACTION", Nothing, New String() {":", "P:Pending", "I:Ignore", "U:Update"})
                ASCMAIN1.Add_Value_List(GRD, "ITEM_BASIC_PROMO", Nothing, New String() {":", "B:Basic", "P:Promo"})
                ASCMAIN1.Add_Value_List(GRD, "COMMODITY_CODE", "Select COMMODITY_CODE, COMMODITY_DESC from ICTCOMM1")
                ' SR-6363 - New Item Master creation by Marketing 
                ASCMAIN1.Add_Value_List(GRD, "COMPLETED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")
                ASCMAIN1.Add_Value_List(GRD, "APPROVED_BY", "SELECT USER_ID, USER_NAME FROM ASTUSER1")

                'With grdE.DisplayLayout.Bands(0)
                '    .Columns("ITEM_CODE").Header.Fixed = True
                '    .Columns("ITEM_DESC").Header.Fixed = True
                'End With


            Case "TATIPITC"
                grdC = GRD
                grdC.Name = "grdC"
                GRDs.Add("C", grdC)

                ' grdC.DisplayLayout.Bands(0).Columns("SEL").Style = UltraWinGrid.ColumnStyle.CheckBox

                ASCMAIN1.grdInitializeLayout(grdC)

                ' Load_Popup_Menu(grdC, "SSSBBBB", "Show GroupBox", "Show Filter")

                Load_Popup_Menu(grdC, "SSSBBBB", "Show GroupBox", "Show Filter", "Show Pins", "Clear Column", "Paste Value to All", "Copy Value to Clipboard", "Paste Value to Selected")


                With grdC.DisplayLayout.Bands(0)

                    Dim P As Integer = -1
                    grdfmt(.Columns("ITEM_CODE"), P, 100, "Item Code", , , True, True)
                    grdfmt(.Columns("ITEM_DESC"), P, 170, "Description", , , True, True)
                    grdfmt(.Columns("IPSA_ACTION"), P, 80, "Action", , , True)
                    'grdfmt(.Columns("SEL"), P, 70, "Sel", , , True, False)
                    grdfmt(.Columns("ITEM_WEIGHT_CURR"), P, 70, "Wgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_WEIGHT"), P, 70, "New Wgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_LENGTH_CURR"), P, 70, "Len Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_LENGTH"), P, 70, "New Len", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_WIDTH_CURR"), P, 70, "Wid Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_WIDTH"), P, 70, "New Wid", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_HEIGHT_CURR"), P, 70, "Hgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_UNIT_HEIGHT"), P, 70, "New Hgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PO_QTY_MULT_CURR"), P, 70, "PO Mult Now", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("ITEM_PO_QTY_MULT"), P, 70, "New PO Mult", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("CARTON_PACK_QTY_CURR"), P, 70, "Ctn Qty Now", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("CARTON_PACK_QTY"), P, 70, "New Ctn Qty", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("CASE_WEIGHT_GRS_CURR"), P, 70, "Ctn Wgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("CASE_WEIGHT_GRS"), P, 70, "New Ctn Wgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_LENGTH_CURR"), P, 70, "Ctn Len Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_LENGTH"), P, 70, "New Ctn Len", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_WIDTH_CURR"), P, 70, "Ctn Wid Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_WIDTH"), P, 70, "New Ctn Wid", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_HEIGHT_CURR"), P, 70, "Ctn Hgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_HEIGHT"), P, 70, "New Ctn Hgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_CUBAGE_CURR"), P, 70, "Ctn Cub Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_CASE_CUBAGE"), P, 70, "New Ctn Cub", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_QTY_CURR"), P, 70, "Plt Qty Now", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("ITEM_PALLET_QTY"), P, 70, "New Plt Qty", Color.Orange, "#,##0", , True)
                    grdfmt(.Columns("ITEM_PALLET_WEIGHT_CURR"), P, 70, "Plt Wgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_WEIGHT"), P, 70, "New Plt Wgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_LENGTH_CURR"), P, 70, "Plt Len Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_LENGTH"), P, 70, "New Plt Len", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_WIDTH_CURR"), P, 70, "Plt Wid Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_WIDTH"), P, 70, "New Plt Wid", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_HEIGHT_CURR"), P, 70, "Plt Hgt Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_HEIGHT"), P, 70, "New Plt Hgt", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_CUBAGE_CURR"), P, 70, "Plt Cub Now", Color.Gold, "#.00", , True)
                    grdfmt(.Columns("ITEM_PALLET_CUBAGE"), P, 70, "New Plt Cub", Color.Gold, "#.00", , True)
                End With

                ASCMAIN1.Add_Value_List(GRD, "IPSA_ACTION", Nothing, New String() {":", "I:Ignore", "U:Update"})
                grdC.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        End Select
    End Sub

    Sub grdfmt(gcol As UltraWinGrid.UltraGridColumn,
               ByRef Position As Integer,
               Width As Integer,
               Caption As String,
               Optional colColor As System.Drawing.Color = Nothing,
               Optional Mask As String = "",
               Optional Fixed As Boolean = False,
               Optional Locked As Boolean = False,
               Optional ViewName As String = "")

        'If gcol.Band.Key = "RSTBOWS1" Then
        '    dst.Tables("RSTBOWSC").Rows.Add(New Object() {gcol.Key, dst.Tables("RSTBOWSC").Rows.Count + 1, gcol.Key.Length})
        'End If

        With gcol
            Position += 1 : .Header.SetVisiblePosition(Position, False)
            .Width = Width
            .Header.Caption = Caption
            If Mask <> "" Then
                If Mask = "x" Then
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .CellAppearance.TextHAlign = HAlign.Center
                Else
                    .Format = Mask
                End If
            End If
            If Fixed Then
                .Header.Fixed = Fixed
            End If
            If Locked Then
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .CellAppearance.BackColor = Color.Beige
            End If
            If colColor = Nothing Then
                colColor = System.Drawing.Color.LightGray
            End If

            With .Header.Appearance
                .BackColor = Color.White
                .BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = colColor
            End With

            If ViewName <> "" Then
                .Style = UltraWinGrid.ColumnStyle.EditButton
                .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            End If
        End With
    End Sub

    Sub Setup_grd(caption As String, Optional TABLE_NAME As String = "")

        tabX2O.Tabs.Add(New UltraWinTabControl.UltraTab)
        Dim grd As UltraWinGrid.UltraGrid = New UltraWinGrid.UltraGrid
        grd.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

        grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
        grd.Parent = tabX2O.Tabs(tabX2O.Tabs.Count - 1).TabPage
        grd.Visible = True
        grd.Dock = DockStyle.Fill

        If caption = "IPSA Item Master" Then
            grdTATIPIT1 = grd
            grdTATIPIT1.Name = "grdTATIPIT1"
            GRDs.Add(Mid(grdTATIPIT1.Name, 4), grdTATIPIT1)
            Load_Popup_Menu(grdTATIPIT1, "B", "Restore as New")
        End If

        tabX2O.Tabs(tabX2O.Tabs.Count - 1).Text = caption

        tabX2O.Tabs(tabX2O.Tabs.Count - 1).Key = caption

        Dim dt As DataTable = ASCDATA1.GetDataTable
        If TABLE_NAME <> "" Then
            dt.TableName = TABLE_NAME


            dt.Columns.Add("IPSA_ACTION")
            If TABLE_NAME <> "TATIPITC" Then
                dt.Columns("IPSA_ACTION").DefaultValue = "I"
            Else
                dt.Columns("IPSA_ACTION").DefaultValue = "0"
            End If


            If TABLE_NAME = "TATIPIC1" Then
                tblE = dt
                '  dst.Tables(TABLE_NAME) = dt
                With tblE
                    .Columns("ITEM_CONTAINS_ALCOHOL").DefaultValue = "0"
                    .Columns("ITEM_CRITICAL_TO_SHIP").DefaultValue = "0"
                    .Columns("ITEM_LOT_CONTROL").DefaultValue = "0"
                    .Columns("ITEM_WEIGHT_CHECK").DefaultValue = "0"
                    .Columns("ITEM_APPR_1ST_REC").DefaultValue = "0"
                    .Columns("ITEM_CRITICAL_TO_SHIP").DefaultValue = "0"
                End With

                For Each row As DataRow In tblE.Select("")
                    row.Item("IPSA_ACTION") = row.Item("IPSA_STATUS_CODE")
                Next
            End If

            If TABLE_NAME = "TATIPITC" Then
                tblC = dt
                With tblC
                    .Columns("SEL").DefaultValue = "0"
                End With
            End If


        End If
        grd.DataSource = dt
        grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

        If TABLE_NAME <> "" Then
            Format_grd(grd, TABLE_NAME)
        End If

        ASCMAIN1.grdInitializeLayout(grd)
        grd.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grd, True)

        For I As Int16 = 0 To dt.Columns.Count - 1
            grd.DisplayLayout.Bands(0).Columns(I).Header.Appearance.TextHAlign = HAlign.Center
        Next
        With grd.DisplayLayout.Override
            .RowSelectors = DefaultableBoolean.True
            .RowSelectorNumberStyle = UltraWinGrid.RowSelectorNumberStyle.VisibleIndex
            .RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
            .RowSelectorAppearance.TextHAlign = HAlign.Center
        End With

        grd.DisplayLayout.CaptionVisible = DefaultableBoolean.True
        grd.DisplayLayout.CaptionAppearance.TextHAlign = HAlign.Left
        grd.Text = caption

        Add_Handlers_grd(grd)

    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATCONVX, "SSS", "Show GroupBox", "Show Filter", "Show Pins")
        'Load_Popup_Menu(grdTATIPIT1, "B", "Restore as New")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        End If

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

            Case "grdTATIPIT1"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing And Not grd.ActiveRow.IsFilterRow And grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Restore as New"), UltraWinToolbars.ButtonTool)

                If grd.Selected.Rows.Count <> 1 Then 'If grd.Selected.Rows.Count = 0 Then

                    tlb_btn.SharedProps.Visible = False
                Else
                    Dim ITEM_CODE As String = grd.Selected.Rows(0).Cells("ITEM_CODE").Value & ""
                    Dim rowictitem1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowictitem1 IsNot Nothing Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        ASCMAIN1.sql = $"Select * from TATIPIC1 where item_code = '{ITEM_CODE}'"
                        Dim rowTATIPIC1 As DataRow = ASCDATA1.GetDataRow
                        If rowTATIPIC1 Is Nothing Then
                            tlb_btn.SharedProps.Visible = False
                        Else
                            If rowTATIPIC1.Item("IPSA_STATUS_CODE") & "" <> "I" Then
                                tlb_btn.SharedProps.Visible = False
                            Else
                                tlb_btn.SharedProps.Visible = True
                            End If
                        End If
                    End If
                    'tlb_btn.SharedProps.Visible = True
                End If

            Case "grdE", "grdC"

                tlb_btn = DirectCast(tlb_pop.Tools("Clear Column"), UltraWinToolbars.ButtonTool)
                If EntryMode <> "E" Or grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow OrElse grd.ActiveCell Is Nothing Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Clear Column " & grd.ActiveCell.Column.Header.Caption
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Value to All"), UltraWinToolbars.ButtonTool)
                If (EntryMode <> "E") Or (COLUMN_NAME_clipboard = "") Or grd.Rows.Count < 1 Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Paste " & COLUMN_CAPTION_clipboard & " = '" & COPY_VALUE_clipboard & "' to All Items"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value to Clipboard"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.Selected.Rows.Count = 0)
                If EntryMode <> "E" Or grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow OrElse grd.ActiveCell Is Nothing OrElse grd.ActiveCell.IsInEditMode OrElse grd.ActiveCell.Value & "" = "" Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Copy " & grd.ActiveCell.Column.Header.Caption & " = '" & grd.ActiveCell.Value & "' to Clipboard"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Value to Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "C" Or EntryMode = "N") And (COLUMN_NAME_clipboard <> "") And grd.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste " & COLUMN_CAPTION_clipboard & " = '" & COPY_VALUE_clipboard & "' to Selected Items"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Paste Value to Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells(COLUMN_NAME_clipboard).Value = COPY_VALUE_clipboard
                    grow.Update()
                Next
                COLUMN_NAME_clipboard = ""
                COLUMN_CAPTION_clipboard = ""
                COPY_VALUE_clipboard = ""

            Case "Paste Value to All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells(COLUMN_NAME_clipboard).Value = COPY_VALUE_clipboard
                        grow.Update()
                    End If
                Next
                COLUMN_NAME_clipboard = ""
                COLUMN_CAPTION_clipboard = ""
                COPY_VALUE_clipboard = ""

        End Select

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Restore as New"
                Dim ITEM_CODE As String = grd.Selected.Rows(0).Cells("ITEM_CODE").Value & ""

                If MsgBox($"Are you sure that you want to Restore Item {ITEM_CODE} as New?", vbYesNo, "Verification") = MsgBoxResult.Yes Then
                    ASCMAIN1.sql = $"Update TATIPIC1 set IPSA_STATUS_CODE = 'P' where  IPSA_STATUS_CODE = 'I' and ITEM_CODE = '{ITEM_CODE}'"
                    ASCDATA1.ExecuteSQL()

                    MsgBox($"Item Code {ITEM_CODE} has been Restored as New", MsgBoxStyle.OkOnly, "Verification")
                End If

            Case "Clear Column"
                If grd.ActiveCell IsNot Nothing Then
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    If COLUMN_NAME = "" Then Exit Sub
                    If COLUMN_NAME = "ITEM_CODE" Then Exit Sub
                    Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)
                    For Each row As DataRow In tbl.Select("", "", DataViewRowState.CurrentRows)
                        row.Item(COLUMN_NAME) = DBNull.Value
                    Next
                End If

            Case "Copy Value To Clipboard"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "ITEM_CODE" Then
                    MsgBox("Cannot Copy And Paste Item Codes")
                    Exit Sub
                End If
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then Exit Sub
                If grd.ActiveRow.Cells(COLUMN_NAME).Value & "" = "" Then
                    Exit Sub
                End If
                COPY_VALUE_clipboard = grd.ActiveRow.Cells(COLUMN_NAME).Value
                COLUMN_NAME_clipboard = COLUMN_NAME
                COLUMN_CAPTION_clipboard = grd.ActiveRow.Cells(COLUMN_NAME).Column.Header.Caption

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

    Sub Load_Workbook(ByVal FILENAME As String, TABLE_NAME As String, ByRef SHEET_NO As Integer)

        Dim INIT_DATE As DateTime
        Dim LAST_DATE As DateTime

        If IPSA_IMPORT_NO = "" Then
            IPSA_IMPORT_NO = ASCMAIN1.Next_Control_No("IPSA_IMPORT_NO")
        End If

        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        Dim fn As String = fi.Name.ToUpper
        Dim ROWCOUNT As Int64 = 0

        Dim rowCFG_FILE As DataRow = dst.Tables("TATCONVX").Rows.Find(Mid(INT2IPSA(TABLE_NAME), 6))

        Dim dt As DataTable = dst.Tables(TABLE_NAME)

        ASCDATA1.ExecuteSQL("Truncate Table " & TABLE_NAME)

        INIT_DATE = Now
        rowCFG_FILE.Item("INIT_DATE") = INIT_DATE

        Dim use_binding As Boolean = True
        Dim batch_rows As Int32 = 100000
        Create_BAs(TABLE_NAME)

        Using sr As New System.IO.StreamReader(FILENAME)
            Do
                Dim lin As String = sr.ReadLine

                Dim delimiter As String = "|"
                'If TABLE_NAME = "ITEM_MASTER" Then delimiter = vbTab
                Dim d() As String = Split(lin, delimiter)

                Dim r As DataRow = dt.NewRow
                ROWCOUNT += 1

                For i As Integer = 0 To dt.Columns.Count - 1
                    If d.Length - 1 >= i Then
                        If Trim(d(i)) <> "" Then
                            r.Item(i) = Trim(d(i))
                        End If
                    End If
                Next

                If TABLE_NAME = "TATIPIT1" Then
                    Dim DESCRIPTION As String = r.Item("DESCRIPTION") & ""
                    If DESCRIPTION.StartsWith(Chr(34)) And DESCRIPTION.EndsWith(Chr(34)) Then
                        DESCRIPTION = Mid(DESCRIPTION, 2, DESCRIPTION.Length - 2)
                    End If
                    Do While InStr(DESCRIPTION, Chr(76) & Chr(63)) <> 0
                        Dim i As Integer = InStr(DESCRIPTION, Chr(76) & Chr(63))
                        DESCRIPTION = Mid(DESCRIPTION, 1, i - 1) & Mid(DESCRIPTION, i + 3)
                    Loop
                    Do While InStr(DESCRIPTION, Chr(63)) <> 0
                        Dim i As Integer = InStr(DESCRIPTION, Chr(63))
                        DESCRIPTION = Mid(DESCRIPTION, 1, i - 1) & Mid(DESCRIPTION, i + 1)
                    Loop
                    'If DESCRIPTION.Length > 40 Then
                    '    DESCRIPTION = Mid(DESCRIPTION, 1, DESCRIPTION.Length - 1)
                    'End If
                    'If Asc(Mid(DESCRIPTION, Len(DESCRIPTION), 1)) > 127 Then
                    '    DESCRIPTION = Mid(DESCRIPTION, 1, DESCRIPTION.Length - 1)
                    'End If
                    r.Item("DESCRIPTION") = DESCRIPTION
                End If

                dt.Rows.Add(r)

                If ROWCOUNT Mod batch_rows = 0 Then
                    If use_binding Then
                        Update_BAs(TABLE_NAME)
                    Else
                        Update_Record_TDA(TABLE_NAME)
                    End If
                    dst.Tables(TABLE_NAME).Rows.Clear()
                End If

            Loop While Not sr.EndOfStream
        End Using

        If use_binding Then
            Update_BAs(TABLE_NAME)
        Else
            Update_Record_TDA(TABLE_NAME)
        End If

        rowCFG_FILE.Item("ROWCOUNT") = ROWCOUNT
        LAST_DATE = Now
        rowCFG_FILE.Item("LAST_DATE") = LAST_DATE
        rowCFG_FILE.Item("SECS") = LAST_DATE.Subtract(INIT_DATE).Seconds


        ASCMAIN1.sql = "Insert into TATIPHIM (HIERARCHY) Select HIERARCHY from TATIPHI1 minus Select HIERARCHY from TATIPHIM"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into TATIPMGM (MERCH_GROUP) Select MERCH_GROUP from TATIPMG1 minus Select MERCH_GROUP from TATIPMGM"
        ASCDATA1.ExecuteSQL()

        My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\archive\" & IPSA_IMPORT_NO & "_" & fi.Name)

    End Sub

    Sub Map_IPSA_to_IPLB()

        ' SR-6363 - New Item Master creation by Marketing 
        ' '0' COMPLETED, NULL COMPLETED_BY, NULL COMPLETED_ON, '0' APPROVED, NULL APPROVED_BY, NULL APPROVED_ON" & v
        ASCMAIN1.sql = "Select " & vbCrLf _
            & "X.ITEM_CODE" & vbCrLf _
            & ", X.DESCRIPTION ITEM_DESC" & vbCrLf _
            & ", SUBSTR(X.EAN,7,6) ITEM_ALT_SORT" & vbCrLf _
            & ", X.GROS_WEIGHT * 2.2046 ITEM_WEIGHT" & vbCrLf _
            & ", NVL(GM.PROD_CODE,'?') PROD_CODE" & vbCrLf _
            & ", NVL(GM.ITEM_CLASS_CODE,'?') ITEM_CLASS_CODE" & vbCrLf _
            & ", NULL ITEM_RETAIL_PRICE" & vbCrLf _
            & ", NVL(HM.COLLECTION_CODE,'?') COLLECTION_CODE" & vbCrLf _
            & ", NVL(ICTPROD1.COST_CATGY_CODE,'?') COST_CATGY_CODE" & vbCrLf _
            & ", 39.3701 * X.ITEM_LENGTH ITEM_UNIT_LENGTH" & vbCrLf _
            & ", 39.3701 * X.WIDTH ITEM_UNIT_WIDTH" & vbCrLf _
            & ", 39.3701 * X.HEIGHT ITEM_UNIT_HEIGHT" & vbCrLf _
            & ", X.QTYCS ITEM_STD_PACK_PUR" & vbCrLf _
            & ", CASE WHEN UPPER(GM.PROD_CODE)='SPGS' THEN X.QTYCS ELSE X.QTYIN END ITEM_STD_PACK_SLS" & vbCrLf _
            & ", X.QTYPL ITEM_PALLET_QTY" & vbCrLf _
            & ", X.PRICE ITEM_COST_STD" & vbCrLf _
            & ", X.QTYCS ITEM_PO_QTY_MIN" & vbCrLf _
            & ", X.QTYCS ITEM_PO_QTY_MULT" & vbCrLf _
            & ", NULL ITEM_SO_QTY_MIN" & vbCrLf _
            & ", NULL ITEM_SO_QTY_MULT" & vbCrLf _
            & ", NULL COMMODITY_CODE" & vbCrLf _
            & ", CASE WHEN UPPER(X.HAZARDOUS)='YES' THEN '1' ELSE '0' END ITEM_CONTAINS_ALCOHOL" & vbCrLf _
            & ", NULL ITEM_DESC2" & vbCrLf _
            & ", X.EAN ITEM_EAN_CODE" & vbCrLf _
            & ", TATCNTRY.COUNTRY_CODE COUNTRY_CODE" & vbCrLf _
            & ", CASE WHEN X.CONTENT IS NOT NULL THEN TRIM(REPLACE(X.CONTENT,'.00','')) || X.CONTENT_UNIT ELSE NULL END SIZE_CODE" & vbCrLf _
            & ", 'E' ITEM_CATGY_CODE" & vbCrLf _
            & ", NULL SEASON_CODE" & vbCrLf _
            & ", NULL LAUNCH_DATE" & vbCrLf _
            & ", X.QTYCS CARTON_PACK_QTY" & vbCrLf _
            & ", NVL(ICTPROD1.PROD_BASIC_PROMO,'B') ITEM_BASIC_PROMO" & vbCrLf _
            & ", '0' ITEM_CRITICAL_TO_SHIP" & vbCrLf _
            & ", CASE WHEN UPPER(X.HAZARDOUS)='YES' THEN '1' ELSE '0' END ITEM_LOT_CONTROL" & vbCrLf _
            & ", '1' ITEM_WEIGHT_CHECK" & vbCrLf _
            & ", '1' ITEM_APPR_1ST_REC" & vbCrLf _
            & ", X.SHELF_LIFE ITEM_SHELF_LIFE_YRS" & vbCrLf _
            & ", 'P' IPSA_STATUS_CODE" & vbCrLf _
            & ", SYSDATE INIT_DATE" & vbCrLf _
            & ", '" & ASCMAIN1.USER_ID & "' INIT_OPER" & vbCrLf _
            & ", SYSDATE LAST_DATE" & vbCrLf _
            & ", '" & ASCMAIN1.USER_ID & "' LAST_OPER" & vbCrLf _
            & ", '0' ITEM_ALLOW_HALF_PACK" & vbCrLf _
            & ", 39.3701 * X.LENCS ITEM_CASE_LENGTH" & vbCrLf _
            & ", 39.3701 * X.WIDCS ITEM_CASE_WIDTH" & vbCrLf _
            & ", 39.3701 * X.HGTCS ITEM_CASE_HEIGHT" & vbCrLf _
            & ", 2.2046 * X.WGTCS CASE_WEIGHT_GRS" & vbCrLf _
            & ", 39.3701 * X.LENCS * 39.3701 * X.WIDCS * 39.3701 * X.HGTCS / 12 / 12 / 12 ITEM_CASE_CUBAGE" & vbCrLf _
            & ", 39.3701 * X.LENPL ITEM_PALLET_LENGTH" & vbCrLf _
            & ", 39.3701 * X.WIDPL ITEM_PALLET_WIDTH" & vbCrLf _
            & ", 39.3701 * X.HGTPL ITEM_PALLET_HEIGHT" & vbCrLf _
            & ", 2.2046 * X.WGTPL ITEM_PALLET_WEIGHT" & vbCrLf _
            & ", 39.3701 * X.LENPL * 39.3701 * X.WIDPL * 39.3701 * X.HGTPL / 12 / 12 / 12 ITEM_PALLET_CUBAGE" & vbCrLf _
            & ", 0 ITEM_VALUE, '0' COMPLETED, NULL COMPLETED_BY, NULL COMPLETED_ON, '0' APPROVED, NULL APPROVED_BY, NULL APPROVED_ON, NULL MARKETING_NOTES" & vbCrLf _
            & " from (" & Get_sql_IPSA_Item_Master() & ") X, TATCNTRY" & vbCrLf _
            & ", TATIPMGM GM, TATIPHIM HM, ICTPROD1" & vbCrLf _
            & " where X.ITEM_CODE in" & vbCrLf _
            & " (Select ITEM_CODE from TATIPIT1 minus Select ITEM_CODE from TATIPIC1)" & vbCrLf _
            & " and TATCNTRY.COUNTRY_CODE2 (+) = X.COUNTRY_ORIGIN" & vbCrLf _
            & " and GM.MERCH_GROUP (+) = X.MERCH_GROUP" & vbCrLf _
            & " and HM.HIERARCHY (+) = X.HIERARCHY" & vbCrLf _
            & " and ICTPROD1.PROD_CODE (+) = GM.PROD_CODE"

        sqlChanges = Replace(ASCMAIN1.sql,
                             "(Select ITEM_CODE from TATIPIT1 minus Select ITEM_CODE from TATIPIC1)",
                             "(Select ITEM_CODE from TATIPIT1)")

        'TATIPHIM IPSA_HIERARCHY_MAP
        'TATIPMGM IPSA_MERCH_GROUP_MAP

        ASCMAIN1.sql = "Insert into TATIPIC1 " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update TATIPIC1 set IPSA_STATUS_CODE = 'X'" & vbCrLf _
            & " where IPSA_STATUS_CODE in ('P','I') and ITEM_CODE in (Select ITEM_CODE from ICTITEM1)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from (" & vbCrLf _
            & Get_sql_IPSA_Item_Master() & vbCrLf _
            & "   ) where ITEM_CODE in (Select ITEM_CODE from TATIPIC1 where IPSA_STATUS_CODE in ('P','I'));" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update TATIPIC1" & vbCrLf _
            & "    Set ITEM_EAN_CODE = R1.EAN" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Function Get_sql_IPLB_Item_Master() As String

        Return "Select" _
            & "  ITEM_CODE" _
            & ", ITEM_DESC" _
            & ", ITEM_ALT_SORT" _
            & ", ITEM_WEIGHT" _
            & ", PROD_CODE" _
            & ", ITEM_CLASS_CODE" _
            & ", ITEM_RETAIL_PRICE" _
            & ", COLLECTION_CODE" _
            & ", COST_CATGY_CODE" _
            & ", ITEM_UNIT_LENGTH" _
            & ", ITEM_UNIT_WIDTH" _
            & ", ITEM_UNIT_HEIGHT" _
            & ", ITEM_STD_PACK_PUR" _
            & ", ITEM_STD_PACK_SLS" _
            & ", ITEM_PALLET_QTY" _
            & ", ITEM_PALLET_LENGTH" _
            & ", ITEM_PALLET_WIDTH" _
            & ", ITEM_PALLET_HEIGHT" _
            & ", ITEM_PALLET_WEIGHT" _
            & ", ITEM_PALLET_CUBAGE" _
            & ", ITEM_COST_STD" _
            & ", ITEM_VALUE" _
            & ", ITEM_PO_QTY_MIN" _
            & ", ITEM_PO_QTY_MULT" _
            & ", ITEM_SO_QTY_MIN" _
            & ", ITEM_SO_QTY_MULT" _
            & ", COMMODITY_CODE" _
            & ", ITEM_CONTAINS_ALCOHOL" _
            & ", ITEM_DESC2" _
            & ", ITEM_EAN_CODE" _
            & ", COUNTRY_CODE" _
            & ", SIZE_CODE" _
            & ", ITEM_CATGY_CODE" _
            & ", SEASON_CODE" _
            & ", LAUNCH_DATE" _
            & ", CARTON_PACK_QTY" _
            & ", ITEM_PALLET_LENGTH" _
            & ", ITEM_PALLET_WIDTH" _
            & ", ITEM_PALLET_HEIGHT" _
            & ", CASE_WEIGHT_GRS" _
            & ", ITEM_CASE_CUBAGE" _
            & ", ITEM_BASIC_PROMO" _
            & ", ITEM_CRITICAL_TO_SHIP" _
            & ", ITEM_LOT_CONTROL" _
            & ", ITEM_WEIGHT_CHECK" _
            & ", ITEM_APPR_1ST_REC" _
            & ", ITEM_SHELF_LIFE_YRS" _
            & " from ICTITEM1"
    End Function

    Function Get_sql_IPSA_Item_Master() As String
        Return "Select M.*, C.PRICE" & vbCrLf _
            & ", P.QTYCS, P.QTYIN, P.QTYPL" & vbCrLf _
            & ", P.LENCS, P.LENIN, P.LENPL" & vbCrLf _
            & ", P.WIDCS, P.WIDIN, P.WIDPL" & vbCrLf _
            & ", P.HGTCS, P.HGTIN, P.HGTPL" & vbCrLf _
            & ", P.WGTCS, P.WGTIN, P.WGTPL" & vbCrLf _
            & " from TATIPIT1 M, TATIPPR1 C, " & vbCrLf _
            & "(Select ITEM_CODE" & vbCrLf _
            & ", MAX(DECODE(PACK,'KAR',QUANTITY,0)) QTYCS" & vbCrLf _
            & ", MAX(DECODE(PACK,'PCB',QUANTITY,0)) QTYIN" & vbCrLf _
            & ", MAX(DECODE(PACK,'PAL',QUANTITY,0)) QTYPL" & vbCrLf _
            & ", MAX(DECODE(PACK,'KAR',LENGTH,0)) LENCS" & vbCrLf _
            & ", MAX(DECODE(PACK,'PCB',LENGTH,0)) LENIN" & vbCrLf _
            & ", MAX(DECODE(PACK,'PAL',LENGTH,0)) LENPL" & vbCrLf _
            & ", MAX(DECODE(PACK,'KAR',WIDTH,0)) WIDCS" & vbCrLf _
            & ", MAX(DECODE(PACK,'PCB',WIDTH,0)) WIDIN" & vbCrLf _
            & ", MAX(DECODE(PACK,'PAL',WIDTH,0)) WIDPL" & vbCrLf _
            & ", MAX(DECODE(PACK,'KAR',HEIGHT,0)) HGTCS" & vbCrLf _
            & ", MAX(DECODE(PACK,'PCB',HEIGHT,0)) HGTIN" & vbCrLf _
            & ", MAX(DECODE(PACK,'PAL',HEIGHT,0)) HGTPL" & vbCrLf _
            & ", MAX(DECODE(PACK,'KAR',GROSS_WGT,0)) WGTCS" & vbCrLf _
            & ", MAX(DECODE(PACK,'PCB',GROSS_WGT,0)) WGTIN" & vbCrLf _
            & ", MAX(DECODE(PACK,'PAL',GROSS_WGT,0)) WGTPL" & vbCrLf _
            & " from TATIPPKG group by ITEM_CODE) P" & vbCrLf _
            & " where P.ITEM_CODE (+) = M.ITEM_CODE" & vbCrLf _
            & "   and C.ITEM_CODE (+) = M.ITEM_CODE"
    End Function

    Private Sub grdE_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdE.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PROD_CODE"
                Dim PROD_CODE As String = e.Cell.Value & ""
                If PROD_CODE <> "" Then
                    Dim rowICTPROD1 As DataRow = LookUp("ICTPROD1", PROD_CODE)
                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = rowICTPROD1.Item("COST_CATGY_CODE") & ""
                End If
                '   If PROD_CODE = "SPGS" Then
                '  e.Cell.Row.Cells("ITEM_STD_PACK_SLS").Value = e.Cell.Row.Cells("CARTON_PACK_QTY").Value
                '  End If

        End Select
    End Sub

    Private Sub grdE_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdE.ClickCellButton

        Dim sql_where As String = ""
        If grdE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True Then

            With e.Cell.Row
                Select Case e.Cell.Column.Key
                    Case "COLLECTION_CODE"
                        grdClickCellButton(grdE, sql_where)
                    Case Else
                        grdClickCellButton(grdE, sql_where)
                End Select
            End With
        End If

    End Sub

    Private Sub grdE_DoubleClickCell(sender As Object, e As UltraWinGrid.DoubleClickCellEventArgs) Handles grdE.DoubleClickCell
        If e.Cell.Column.Key = "IPSA_ACTION" Then

        End If
    End Sub

    Private Sub grdE_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdE.InitializeRow
        With e.Row.Cells("IPSA_ACTION")
            If .Value & "" = "U" Then
                '.Appearance.BackColor = Color.LightGreen
                .Appearance = App_LtGrn
            ElseIf .Value & "" = "I" Then
                .Appearance.BackColor = Color.Yellow
            Else
                .Appearance.BackColor = Color.Empty
            End If
        End With

    End Sub

    Sub Get_Changes()

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, '0' SEL" & vbCrLf _
            & ", ICTITEM1.ITEM_UNIT_LENGTH ITEM_UNIT_LENGTH_CURR, ROUND(X.ITEM_UNIT_LENGTH,3) ITEM_UNIT_LENGTH" & vbCrLf _
            & ", ICTITEM1.ITEM_UNIT_WIDTH ITEM_UNIT_WIDTH_CURR, ROUND(X.ITEM_UNIT_WIDTH,3) ITEM_UNIT_WIDTH" & vbCrLf _
            & ", ICTITEM1.ITEM_UNIT_HEIGHT ITEM_UNIT_HEIGHT_CURR, ROUND(X.ITEM_UNIT_HEIGHT,3) ITEM_UNIT_HEIGHT" & vbCrLf _
            & ", ICTITEM1.ITEM_WEIGHT ITEM_WEIGHT_CURR, ROUND(X.ITEM_WEIGHT,2) ITEM_WEIGHT" & vbCrLf _
            & ", ICTITEM1.CARTON_PACK_QTY CARTON_PACK_QTY_CURR, X.CARTON_PACK_QTY" & vbCrLf _
            & ", ICTITEM1.ITEM_CASE_LENGTH ITEM_CASE_LENGTH_CURR, ROUND(X.ITEM_CASE_LENGTH,3) ITEM_CASE_LENGTH" & vbCrLf _
            & ", ICTITEM1.ITEM_CASE_WIDTH ITEM_CASE_WIDTH_CURR, ROUND(X.ITEM_CASE_WIDTH,3) ITEM_CASE_WIDTH" & vbCrLf _
            & ", ICTITEM1.ITEM_CASE_HEIGHT ITEM_CASE_HEIGHT_CURR, ROUND(X.ITEM_CASE_HEIGHT,3) ITEM_CASE_HEIGHT" & vbCrLf _
            & ", ICTITEM1.ITEM_CASE_CUBAGE ITEM_CASE_CUBAGE_CURR, ROUND(X.ITEM_CASE_CUBAGE,3) ITEM_CASE_CUBAGE" & vbCrLf _
            & ", ICTITEM1.CASE_WEIGHT_GRS CASE_WEIGHT_GRS_CURR, ROUND(X.CASE_WEIGHT_GRS,2) CASE_WEIGHT_GRS" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_QTY ITEM_PALLET_QTY_CURR, X.ITEM_PALLET_QTY" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_LENGTH ITEM_PALLET_LENGTH_CURR, ROUND(X.ITEM_PALLET_LENGTH,3) ITEM_PALLET_LENGTH" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_WIDTH ITEM_PALLET_WIDTH_CURR, ROUND(X.ITEM_PALLET_WIDTH,3) ITEM_PALLET_WIDTH" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_HEIGHT ITEM_PALLET_HEIGHT_CURR, ROUND(X.ITEM_PALLET_HEIGHT,3) ITEM_PALLET_HEIGHT" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_CUBAGE ITEM_PALLET_CUBAGE_CURR, ROUND(X.ITEM_PALLET_CUBAGE,3) ITEM_PALLET_CUBAGE" & vbCrLf _
            & ", ICTITEM1.ITEM_PALLET_WEIGHT ITEM_PALLET_WEIGHT_CURR, ROUND(X.ITEM_PALLET_WEIGHT,2) ITEM_PALLET_WEIGHT" & vbCrLf _
            & ", ICTITEM1.ITEM_PO_QTY_MULT ITEM_PO_QTY_MULT_CURR, X.ITEM_PO_QTY_MULT " & vbCrLf _
            & " from ICTITEM1, " & vbCrLf _
            & IIf(sqlChanges = "", "ICTITEM1", "(" & sqlChanges & ")") & vbCrLf _
            & " X WHERE X.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
            & " and (" & vbCrLf _
            & "ROUND(NVL(X.ITEM_UNIT_LENGTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_UNIT_LENGTH,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.ITEM_UNIT_WIDTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_UNIT_WIDTH,0),3) OR " & vbCrLf _
            & "ROUND(NVL(X.ITEM_UNIT_HEIGHT,0),3) <> ROUND(NVL(ICTITEM1.ITEM_UNIT_HEIGHT,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.ITEM_WEIGHT,0),2) <> ROUND(NVL(ICTITEM1.ITEM_WEIGHT,0),2) OR" & vbCrLf _
            & "NVL(X.CARTON_PACK_QTY,0) <>  NVL(ICTITEM1.CARTON_PACK_QTY,0) OR " & vbCrLf _
            & "ROUND(NVL(X.ITEM_CASE_LENGTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_CASE_LENGTH,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.ITEM_CASE_WIDTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_CASE_WIDTH,0),3) OR " & vbCrLf _
            & "ROUND(NVL(X.ITEM_CASE_HEIGHT,0),3) <> ROUND(NVL(ICTITEM1.ITEM_CASE_HEIGHT,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.CASE_WEIGHT_GRS,0),2) <> ROUND(NVL(ICTITEM1.CASE_WEIGHT_GRS,0),2) OR" & vbCrLf _
            & "NVL(X.ITEM_PALLET_QTY,0) <>  NVL(ICTITEM1.ITEM_PALLET_QTY,0) OR " & vbCrLf _
            & "ROUND(NVL(X.ITEM_PALLET_LENGTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_PALLET_LENGTH,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.ITEM_PALLET_WIDTH,0),3) <> ROUND(NVL(ICTITEM1.ITEM_PALLET_WIDTH,0),3) OR " & vbCrLf _
            & "ROUND(NVL(X.ITEM_PALLET_HEIGHT,0),3) <> ROUND(NVL(ICTITEM1.ITEM_PALLET_HEIGHT,0),3) OR" & vbCrLf _
            & "ROUND(NVL(X.ITEM_PALLET_WEIGHT,0),2) <> ROUND(NVL(ICTITEM1.ITEM_PALLET_WEIGHT,0),2))"

        '& "NVL(X.ITEM_PO_QTY_MULT,0) <>  NVL(ICTITEM1.ITEM_PO_QTY_MULT,0) OR " & vbCrLf _
        '& "NVL(X.ITEM_STD_PACK_PUR,0) <>  NVL(ICTITEM1.ITEM_STD_PACK_PUR,0) OR " & vbCrLf _
        '& "NVL(X.ITEM_PO_QTY_MIN,0) <>  NVL(ICTITEM1.ITEM_PO_QTY_MIN,0))"
        If TATIPITC = "" Then
            TATIPITC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & TATIPITC & " Add Primary Key (ITEM_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & TATIPITC)
            ASCDATA1.ExecuteSQL("Insert into " & TATIPITC & " " & ASCMAIN1.sql)
        End If
    End Sub

    Private Sub btnUpdateAll_Click(sender As Object, e As EventArgs) Handles btnUpdateAll.Click

        For Each row As DataRow In tblC.Select("")
            row.Item("IPSA_ACTION") = "U"
        Next
    End Sub

    Private Sub tabX2O_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabX2O.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        If (EntryMode = "C" Or EntryMode = "N") AndAlso tabX2O.SelectedTab IsNot Nothing AndAlso tabX2O.SelectedTab.Text = "Changes" Then
            UltraExplorerBar1.Groups("Data Controls").Visible = True
        Else
            UltraExplorerBar1.Groups("Data Controls").Visible = False
        End If

        COLUMN_NAME_clipboard = ""
        COLUMN_CAPTION_clipboard = ""
        COPY_VALUE_clipboard = ""
    End Sub

    Private Sub grdTATCONVX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATCONVX.InitializeLayout

    End Sub

    Private Sub grdX2O_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdX2O.BeforeRowUpdate

    End Sub

    Private Sub grdE_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdE.AfterRowUpdate
        If grdE.ActiveRow.Cells("PROD_CODE").Text = "SPGS" Then
            grdE.ActiveRow.Cells("ITEM_STD_PACK_SLS").Value = grdE.ActiveRow.Cells("CARTON_PACK_QTY").Value
        End If
    End Sub

End Class