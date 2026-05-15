Imports ABSolution

Public Class EDF816T1

    Private Const CustomersToProcess As String = "'BBB', 'KOHLS', 'STEINMART'"
    Private Const EDI_PURPOSE_CODE_CHANGE As String = "04"
    Private Const EDI_PURPOSE_CODE_ORIGINAL As String = "00"

    Private Const EDI_MAINT_TYPE_CHANGE As String = "001" ' 001 Change (all data about the location is replaced)
    Private Const EDI_MAINT_TYPE_DELETE As String = "002" '002 Delete (location is to be deleted)
    Private Const EDI_MAINT_TYPE_ADD As String = "021" '021 Addition (new location is to be added)

    Private EDI_DOC_SEQ_NO As String = String.Empty
    Private EDI_ISA_NO As String = String.Empty
    Private EDI_CUST_CODE As String = String.Empty
    Private wktable As String = String.Empty
    Private filterQuery As String = String.Empty



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "EDT816T1", "*")
            Create_TDA(.Tables.Add, "EDT816T2", "*")
            Create_TDA(.Tables.Add, "EDT816T3", "*")

            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            .Tables("ARTCUST2").Columns.Add("MODIFIED", GetType(System.String))

            Create_Relation("ARTCUST1", "EDT816T1", "CUST_CODE", "CUST_CODE")
            Create_Relation("EDT816T1", "EDT816T2", "EDI_DOC_SEQ_NO", "EDI_DOC_SEQ_NO")
            Create_Relation("EDT816T2", "EDT816T3", "EDI_DOC_SEQ_NO,EDI_DTL_SEQ", "EDI_DOC_SEQ_NO,EDI_DTL_SEQ")

            .Tables("EDT816T2").Columns.Add("NUM_CHILDREN", GetType(System.Int16), "COUNT(CHILD.EDI_DOC_SEQ_NO)")
            .Tables("EDT816T1").Columns.Add("NUM_CHILDREN", GetType(System.Int16), "COUNT(CHILD.EDI_DOC_SEQ_NO)")
            .Tables("ARTCUST1").Columns.Add("NUM_CHILDREN", GetType(System.Int16), "COUNT(CHILD.EDI_DOC_SEQ_NO)")

        End With

        grdEDT816T1X.DataSource = dst.Tables("ARTCUST1")
        Create_Summary(grdEDT816T1X, "CUST_CODE", "Count")
        Create_Summary(grdEDT816T1X, "EDI_DOC_SEQ_NO", "Count", grdEDT816T1X.DisplayLayout.Bands(1).Key)
        Create_Summary(grdEDT816T1X, "EDI_DOC_SEQ_NO", "Count", grdEDT816T1X.DisplayLayout.Bands(2).Key)

        ASCMAIN1.Add_Value_List(grdEDT816T1X, "EDI_PURPOSE_CODE", , New String() {":", "04:Change", "00:Original"}, 1)

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")


        grdEDT816T1X.Dock = DockStyle.Fill
        grdARTCUST2.Dock = DockStyle.Fill

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Load"

                If EDI_ISA_NO.Length = 0 AndAlso EDI_CUST_CODE.Length = 0 AndAlso EDI_DOC_SEQ_NO.Length = 0 Then
                    If grdEDT816T1X.Selected.Rows.Count = 0 Then
                        EMsg &= vbCr & "You must select an entry from the list of EDI 816 entries."
                        Exit Select
                    Else
                        Select Case grdEDT816T1X.Selected.Rows(0).Band.Key
                            Case grdEDT816T1X.DisplayLayout.Bands(0).Key
                                EDI_CUST_CODE = grdEDT816T1X.Selected.Rows(0).Cells("CUST_CODE").Text & String.Empty

                            Case grdEDT816T1X.DisplayLayout.Bands(1).Key
                                EDI_CUST_CODE = grdEDT816T1X.Selected.Rows(0).Cells("CUST_CODE").Text & String.Empty
                                EDI_DOC_SEQ_NO = grdEDT816T1X.Selected.Rows(0).Cells("EDI_DOC_SEQ_NO").Text & String.Empty
                                If EDI_CUST_CODE = "KOHLS" Then
                                    EDI_ISA_NO = grdEDT816T1X.Selected.Rows(0).Cells("EDI_ISA_NO").Text & String.Empty
                                End If
                        End Select
                    End If
                End If

                If EDI_ISA_NO.Length > 0 Then
                    If MessageBox.Show("Do you want to Process EDI 816 records for Customer " & EDI_CUST_CODE & ", ISA NO " & EDI_ISA_NO & "?", "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        EDI_DOC_SEQ_NO = String.Empty
                        EDI_ISA_NO = String.Empty
                        EDI_CUST_CODE = String.Empty
                        Exit Sub
                    End If
                ElseIf EDI_DOC_SEQ_NO.Length > 0 Then
                    If MessageBox.Show("Do you want to Process EDI 816 records for Customer " & EDI_CUST_CODE & ", EDI Doc Sequence No " & EDI_DOC_SEQ_NO & "?", "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        EDI_DOC_SEQ_NO = String.Empty
                        EDI_ISA_NO = String.Empty
                        EDI_CUST_CODE = String.Empty
                        Exit Sub
                    End If
                ElseIf EDI_CUST_CODE.Length > 0 Then
                    If MessageBox.Show("Do you want to Process All EDI 816 record(s) for Customer " & EDI_CUST_CODE & "?", "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        EDI_DOC_SEQ_NO = String.Empty
                        EDI_ISA_NO = String.Empty
                        EDI_CUST_CODE = String.Empty
                        Exit Sub
                    End If
                Else
                    EMsg &= "Could not be determined which EDI 816 documents(s) to process."
                End If

                ' Warn the user about any previous 816T1 entries
                If EDI_DOC_SEQ_NO.Length > 0 AndAlso EMsg.Length = 0 Then
                    Dim query As String = "CUST_CODE = '" & EDI_CUST_CODE & "' and EDI_DOC_SEQ_NO < '" & EDI_DOC_SEQ_NO & "'"
                    If EDI_CUST_CODE = "KOHLS" Then
                        query &= " and EDI_ISA_NO <> '" & EDI_ISA_NO & "'"
                    End If

                    Dim numRecords As Int16 = dst.Tables("EDT816T1").Compute("COUNT(EDI_DOC_SEQ_NO)", query)
                    If numRecords > 0 Then
                        If MessageBox.Show("There are " & numRecords & " previous EDI 816 documents. If you continue, older data may overwrite new data. Do you want to continue?", _
                                             "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            EDI_DOC_SEQ_NO = String.Empty
                            EDI_ISA_NO = String.Empty
                            EDI_CUST_CODE = String.Empty
                            Exit Sub
                        End If
                    End If
                End If


            Case "Update"
                If MessageBox.Show("Do you want to Update changes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        If ScreenMode Then

        Else
            Clear_Record()
        End If

        grdEDT816T1X.Visible = Not ScreenMode
        grdARTCUST2.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        For Each table As String In New String() {"ARTCUST1", "EDT816T1", "EDT816T2", "EDT816T3", "ARTCUST2"}
            dst.Tables(table).Rows.Clear()
        Next

        EDI_DOC_SEQ_NO = String.Empty
        EDI_ISA_NO = String.Empty
        EDI_CUST_CODE = String.Empty
        filterQuery = String.Empty

        Try
            ASCMAIN1.Progress("Laoding Data", "")

            Dim sql As String = "Select * from EDT816T1 WHERE EDI_PROCESS_IND = '1' AND CUST_CODE IN (" & CustomersToProcess & ")"
            If wktable.Length = 0 Then
                wktable = ASCMAIN1.Temp_Table(sql)
            Else
                ASCDATA1.ExecuteSQL("Truncate Table " & wktable)
                ASCDATA1.ExecuteSQL("Insert Into " & wktable & " " & sql)
            End If

            ASCMAIN1.Progress("-", "ARTCUST1")
            sql = "Select * from ARTCUST1 WHERE CUST_CODE IN (SELECT CUST_CODE FROM " & wktable & ")"
            Fill_Records("ARTCUST1", String.Empty, True, sql)

            For Each tableName As String In New String() {"EDT816T1", "EDT816T2", "EDT816T3"}
                sql = "Select * from " & tableName & " WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & wktable & ")"
                ASCMAIN1.Progress("-", tableName)
                Fill_Records(tableName, String.Empty, True, sql)
            Next

            Sort_grdColumns(grdEDT816T1X, "CUST_CODE", False, 0)
            Sort_grdColumns(grdEDT816T1X, "EDI_DOC_SEQ_NO", False, 1)
            Sort_grdColumns(grdEDT816T1X, "EDI_DOC_SEQ_NO,EDI_DTL_SEQ", False, 2)
            Sort_grdColumns(grdEDT816T1X, "EDI_DOC_SEQ_NO,EDI_DTL_SEQ,EDI_DTL2_SEQ", False, 3)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            For Each table As String In New String() {"ARTCUST1", "EDT816T1", "EDT816T2", "EDT816T3", "ARTCUST2"}
                dst.Tables(table).Rows.Clear()
            Next
        Finally
            ASCMAIN1.Progress("", "")
        End Try

        EnforceConstraints(True)

    End Sub

    Private Sub Load_Record()

        Try
            ASCMAIN1.Progress("Loading EDI 816 data")

            EnforceConstraints(False)

            Dim listCodes As New List(Of String)

            If EDI_ISA_NO.Length > 0 Then
                filterQuery = "CUST_CODE = '" & EDI_CUST_CODE & "' AND EDI_ISA_NO = '" & EDI_ISA_NO & "'"
            ElseIf EDI_DOC_SEQ_NO.Length > 0 Then
                filterQuery = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ElseIf EDI_CUST_CODE.Length > 0 Then
                filterQuery = "CUST_CODE = '" & EDI_CUST_CODE & "'"
            Else
                Exit Sub
            End If


            Select Case EDI_CUST_CODE

                Case "KOHLS"
                    listCodes.Add("BU")
                    listCodes.Add("WH")

                Case "BBB"
                    listCodes.Add("BY")


                Case "STEINMART"
                    listCodes.Add("BU")
                    listCodes.Add("WH")

                Case Else
                    listCodes.Add("BU")
                    listCodes.Add("WH")

            End Select

            EnforceConstraints(True)

            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Select(filterQuery, "EDI_DOC_SEQ_NO")
                EDI_DOC_SEQ_NO = rowEDT816T1.Item("EDI_DOC_SEQ_NO") & String.Empty

                For Each CUST_ADDR_CODE As String In listCodes
                    For Each rowEDT816T2 As DataRow In dst.Tables("EDT816T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "'", "EDI_DTL_SEQ")
                        Dim EDI_ADDR_CODE As String = rowEDT816T2.Item("EDI_ADDR_CODE") & String.Empty


                        rowEDT816T2.Item("CUST_NAME") = (rowEDT816T2.Item("CUST_NAME") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_ADDR1") = (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_ADDR2") = (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_CITY") = (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_STATE") = (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_ZIP_CODE") = (rowEDT816T2.Item("CUST_ZIP_CODE") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_PHONE") = (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper.Trim
                        rowEDT816T2.Item("CUST_FAX") = (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper.Trim

                        ' Currently the three customers stores are padded with 0's to a lenght of 6
                        EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(6, "0")
                        If EDI_ADDR_CODE = "001375" Then Stop

                        ASCMAIN1.Progress("Processing " & EDI_DOC_SEQ_NO & "/" & EDI_ADDR_CODE)

                        ' Prevents an error if the same door is provided twice
                        Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New Object() {EDI_CUST_CODE, EDI_ADDR_CODE})
                        If rowARTCUST2 Is Nothing Then

                            Fill_Records("ARTCUST2", New Object() {EDI_CUST_CODE, EDI_ADDR_CODE}, False)
                            rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New Object() {EDI_CUST_CODE, EDI_ADDR_CODE})

                            If rowARTCUST2 Is Nothing Then

                                ' Special Handling for a missing Kohls DC.
                                If EDI_CUST_CODE = "KOHLS" AndAlso CUST_ADDR_CODE = "WH" Then
                                    MessageBox.Show("KOHLS DC " & EDI_ADDR_CODE & " cannot be located in Store Master.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Continue For
                                End If

                                rowARTCUST2 = dst.Tables("ARTCUST2").NewRow
                                rowARTCUST2.Item("CUST_CODE") = EDI_CUST_CODE
                                rowARTCUST2.Item("CUST_STORE_NO") = EDI_ADDR_CODE
                                rowARTCUST2.Item("CUST_STORE_STATUS") = "A"
                                rowARTCUST2.Item("CUST_DC_IND") = "0"
                                dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)

                                rowARTCUST2.Item("MODIFIED") = "1"
                            End If
                        End If

                        ' Converts Nulls to zero.
                        rowARTCUST2.Item("MODIFIED") = Val(rowARTCUST2.Item("MODIFIED") & String.Empty)

                        Select Case EDI_CUST_CODE

                            Case "BBB"

                                'EDT816T2.CUST_NAME -> UPPER(ARTCUST2.CUST_STORE_LOCATION)
                                If rowARTCUST2.Item("CUST_STORE_LOCATION") & String.Empty <> (rowEDT816T2.Item("CUST_NAME") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_LOCATION") = (rowEDT816T2.Item("CUST_NAME") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T2.CUST_ADDR1 -> UPPER(ARTCUST2.CUST_STORE_NAME)
                                If rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_NAME") = (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_ADDR2 -> UPPER(ARTCUST2.CUST_STORE_ADDR1)
                                If rowARTCUST2.Item("CUST_STORE_ADDR1") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_ADDR1") = (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_CITY -> UPPER(ARTCUST2.CUST_STORE_CITY)
                                If rowARTCUST2.Item("CUST_STORE_CITY") & String.Empty <> (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_CITY") = (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_STATE -> UPPER(ARTCUST2.CUST_STORE_STATE)
                                If rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty <> (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_STATE") = (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_ZIP_CODE -> ARTCUST2.CUST_STORE_ZIP_CODE (IF 9 WITHOUT -, NEED TO FORMAT)
                                Dim CUST_ZIP_CODE As String = rowEDT816T2.Item("CUST_ZIP_CODE") & String.Empty
                                CUST_ZIP_CODE = CUST_ZIP_CODE.Trim
                                If CUST_ZIP_CODE.Length = 9 AndAlso Not CUST_ZIP_CODE.Contains("-") Then
                                    CUST_ZIP_CODE = CUST_ZIP_CODE.Substring(0, 5) & "-" & CUST_ZIP_CODE.Substring(5)
                                End If

                                If rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & String.Empty <> CUST_ZIP_CODE Then
                                    rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = CUST_ZIP_CODE
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_PHONE -> ARTCUST2.CUST_STORE_PHONE
                                If rowARTCUST2.Item("CUST_STORE_PHONE") & String.Empty <> (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_PHONE") = (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                'EDT816T3.CUST_FAX -> ARTCUST2.CUST_STORE_FAX
                                If rowARTCUST2.Item("CUST_STORE_FAX") & String.Empty <> (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_FAX") = (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                            Case "KOHLS"

                                Select Case CUST_ADDR_CODE
                                    Case "BU"
                                        Dim CUST_STORE_NAME As String = "KOHLS " & rowEDT816T2.Item("CUST_NAME") & " #" & EDI_ADDR_CODE
                                        If rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty <> CUST_STORE_NAME Then
                                            rowARTCUST2.Item("CUST_STORE_NAME") = CUST_STORE_NAME
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_ADDR1") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper Then
                                            rowARTCUST2.Item("CUST_STORE_ADDR1") = (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_ADDR2") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper Then
                                            rowARTCUST2.Item("CUST_STORE_ADDR2") = (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_CITY") & String.Empty <> (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper Then
                                            rowARTCUST2.Item("CUST_STORE_CITY") = (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty <> (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper Then
                                            rowARTCUST2.Item("CUST_STORE_STATE") = (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        Dim CUST_ZIP_CODE As String = rowEDT816T2.Item("CUST_ZIP_CODE") & String.Empty
                                        CUST_ZIP_CODE = CUST_ZIP_CODE.Trim
                                        If CUST_ZIP_CODE.Length = 9 AndAlso Not CUST_ZIP_CODE.Contains("-") Then
                                            CUST_ZIP_CODE = CUST_ZIP_CODE.Substring(0, 5) & "-" & CUST_ZIP_CODE.Substring(5)
                                        End If
                                        If rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & String.Empty <> CUST_ZIP_CODE Then
                                            rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = CUST_ZIP_CODE
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_PHONE") & String.Empty <> (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper _
                                            AndAlso (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.Length > 0 Then
                                            rowARTCUST2.Item("CUST_STORE_PHONE") = (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                        If rowARTCUST2.Item("CUST_STORE_FAX") & String.Empty <> (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper _
                                            AndAlso (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.Length > 0 Then
                                            rowARTCUST2.Item("CUST_STORE_FAX") = (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper
                                            rowARTCUST2.Item("MODIFIED") = "1"
                                        End If

                                    Case "WH"
                                        rowARTCUST2.Item("CUST_DC_IND") = "1"
                                        ' Need to update any stores attached to this DC
                                        Dim EDI_DTL_SEQ As Int16 = rowEDT816T2.Item("EDI_DTL_SEQ")
                                        Dim EDI_DOC_SEQ_NO As String = rowEDT816T2.Item("EDI_DOC_SEQ_NO")
                                        For Each rowEDT816T3 As DataRow In dst.Tables("EDT816T3").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND EDI_DTL_SEQ = " & EDI_DTL_SEQ, "EDI_DTL2_SEQ")
                                            Dim EDI_ADDR_CODE_3 As String = rowEDT816T3.Item("EDI_ADDR_CODE") & String.Empty
                                            EDI_ADDR_CODE_3 = EDI_ADDR_CODE_3.PadLeft(6, "0")

                                            ASCMAIN1.Progress("Processing DC " & EDI_ADDR_CODE, EDI_ADDR_CODE_3)

                                            rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New Object() {EDI_CUST_CODE, EDI_ADDR_CODE_3})
                                            If rowARTCUST2 Is Nothing Then
                                                Fill_Records("ARTCUST2", New Object() {EDI_CUST_CODE, EDI_ADDR_CODE_3}, False)
                                            End If

                                            rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New Object() {EDI_CUST_CODE, EDI_ADDR_CODE_3})
                                            If rowARTCUST2 Is Nothing Then
                                                MessageBox.Show("Cannot Locate Store " & EDI_ADDR_CODE_3 & ". Assignment of DC (" & EDI_ADDR_CODE & ") skipped.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            Else
                                                If rowARTCUST2.Item("CUST_DC_NO") & String.Empty <> EDI_ADDR_CODE Then
                                                    rowARTCUST2.Item("CUST_DC_NO") = EDI_ADDR_CODE
                                                    rowARTCUST2.Item("MODIFIED") = "1"
                                                End If
                                            End If
                                        Next

                                End Select

                            Case "STEINMART"
                                Dim CUST_STORE_NAME As String = "STEIN MART  " & " #" & Val(EDI_ADDR_CODE)
                                If rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty <> CUST_STORE_NAME Then
                                    rowARTCUST2.Item("CUST_STORE_NAME") = CUST_STORE_NAME
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_ADDR1") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_ADDR1") = (rowEDT816T2.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_ADDR2") & String.Empty <> (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_ADDR2") = (rowEDT816T2.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_CITY") & String.Empty <> (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_CITY") = (rowEDT816T2.Item("CUST_CITY") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty <> (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper Then
                                    rowARTCUST2.Item("CUST_STORE_STATE") = (rowEDT816T2.Item("CUST_STATE") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                Dim CUST_ZIP_CODE As String = rowEDT816T2.Item("CUST_ZIP_CODE") & String.Empty
                                CUST_ZIP_CODE = CUST_ZIP_CODE.Trim
                                If CUST_ZIP_CODE.Length = 9 AndAlso Not CUST_ZIP_CODE.Contains("-") Then
                                    CUST_ZIP_CODE = CUST_ZIP_CODE.Substring(0, 5) & "-" & CUST_ZIP_CODE.Substring(5)
                                End If
                                If rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & String.Empty <> CUST_ZIP_CODE Then
                                    rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = CUST_ZIP_CODE
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_PHONE") & String.Empty <> (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper _
                                    AndAlso (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.Length > 0 Then
                                    rowARTCUST2.Item("CUST_STORE_PHONE") = (rowEDT816T2.Item("CUST_PHONE") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                If rowARTCUST2.Item("CUST_STORE_FAX") & String.Empty <> (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper _
                                    AndAlso (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.Length > 0 Then
                                    rowARTCUST2.Item("CUST_STORE_FAX") = (rowEDT816T2.Item("CUST_FAX") & String.Empty).ToString.ToUpper
                                    rowARTCUST2.Item("MODIFIED") = "1"
                                End If

                                Dim CUST_DC_CODE As String = (rowEDT816T2.Item("CUST_DC_CODE") & String.Empty).ToString.ToUpper
                                If CUST_DC_CODE.Length > 0 Then
                                    CUST_DC_CODE = CUST_DC_CODE.PadLeft(6, "0")
                                    If rowARTCUST2.Item("CUST_DC_NO") & String.Empty <> CUST_DC_CODE Then
                                        rowARTCUST2.Item("CUST_DC_NO") = CUST_DC_CODE
                                        rowARTCUST2.Item("MODIFIED") = "1"
                                    End If
                                End If
                        End Select
                    Next
                Next
            Next

            For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("MODIFIED = '0'")
                dst.Tables("ARTCUST2").Rows.Remove(rowARTCUST2)
            Next

            For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("")
                If rowARTCUST2.RowState = DataRowState.Added Then
                    rowARTCUST2.Item("MODIFIED") = "Added"
                Else
                    rowARTCUST2.Item("MODIFIED") = "Modified"
                End If
            Next


            Clear_All_Filters(grdARTCUST2)
            Sort_grdColumns(grdARTCUST2, "CUST_STORE_NO")
            grdARTCUST2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            INIT_LAST("ARTCUST2")
            Update_Record_TDA("ARTCUST2")
            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Select(filterQuery)
                rowEDT816T1.Item("EDI_PROCESS_IND") = "2"
            Next
            Update_Record_TDA("EDT816T1")

            CommitTrans("Update Successful")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT816T1X, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTCUST2, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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


        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdEDT816T1X_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT816T1X.DoubleClickRow

        EDI_CUST_CODE = String.Empty
        EDI_DOC_SEQ_NO = String.Empty
        EDI_ISA_NO = String.Empty

        Select Case e.Row.Band.Key
            Case grdEDT816T1X.DisplayLayout.Bands(0).Key
                EDI_CUST_CODE = e.Row.Cells("CUST_CODE").Text & String.Empty

            Case grdEDT816T1X.DisplayLayout.Bands(1).Key
                EDI_DOC_SEQ_NO = e.Row.Cells("EDI_DOC_SEQ_NO").Text & String.Empty
                EDI_CUST_CODE = e.Row.Cells("CUST_CODE").Text & String.Empty

                If EDI_CUST_CODE = "KOHLS" Then
                    EDI_ISA_NO = e.Row.Cells("EDI_ISA_NO").Text & String.Empty
                End If

            Case Else
                Exit Sub
        End Select


        Click_Command("Load")
    End Sub

    Private Sub grdARTCUST2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST2.InitializeRow

        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Text
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Text

        Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New Object() {CUST_CODE, CUST_STORE_NO})
        If rowARTCUST2.RowState = DataRowState.Added Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        Else
            For Each dcol As DataColumn In rowARTCUST2.Table.Columns
                If rowARTCUST2.Item(dcol.ColumnName, DataRowVersion.Current) & String.Empty <> rowARTCUST2.Item(dcol.ColumnName, DataRowVersion.Original) & String.Empty Then
                    e.Row.Cells(dcol.ColumnName).Appearance.BackColor = Drawing.Color.LightBlue
                End If
            Next
        End If
    End Sub

#End Region

End Class