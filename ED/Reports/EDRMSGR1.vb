
Public Class EDRMSGR1

    Private EDI_DOC_SEQ_NO_860 As String = String.Empty
    Private EDI_DOC_SEQ_NO_816 As String = String.Empty
    Private EDI_DOC_SEQ_NO_864 As String = String.Empty
    Private EDI_DOC_SEQ_NO_824 As String = String.Empty

    Private enable816 As Boolean = False
    Private enable824 As Boolean = False
    Private enable860 As Boolean = False
    Private enable864 As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim tblEDTTRPM1 As DataTable = ASCDATA1.GetDataTable("SELECT DISTINCT EDI_DOC_NO FROM EDTTRPM1")

        chkEDT816.Enabled = False
        chkEDT816.Checked = False

        chkEDT824.Enabled = False
        chkEDT824.Checked = False

        chkEDT860.Enabled = False
        chkEDT860.Checked = False

        chkEDT864.Enabled = False
        chkEDT864.Checked = False

        For Each row As DataRow In tblEDTTRPM1.Rows
            Select Case row.Item("EDI_DOC_NO") & String.Empty

                Case "816"
                    chkEDT816.Enabled = True
                    chkEDT816.Checked = True
                    enable816 = True
                Case "824"
                    chkEDT824.Enabled = True
                    chkEDT824.Checked = True
                    enable824 = True
                Case "860"
                    chkEDT860.Enabled = True
                    chkEDT860.Checked = True
                    enable860 = True
                Case "864"
                    chkEDT864.Enabled = True
                    chkEDT864.Checked = True
                    enable864 = True
            End Select

        Next


    End Sub

    Public Overrides Sub Proceed_PreReq(eItemKey As String)
        MyBase.Proceed_PreReq(eItemKey)

        Select Case eItemKey

            Case "Proceed"
                If chkEDT816.Checked OrElse chkEDT824.Checked OrElse chkEDT860.Checked OrElse chkEDT864.Checked Then
                    ' nothing at this time
                Else
                    EMsg &= vbCr & "You must select at least one EDI table to extract."
                End If

        End Select
    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)
        RWU = "R"

        Dim dataToPrint As Boolean = False

        If chkEDT816.Checked AndAlso dst.Tables("EDT816T1").Rows.Count > 0 Then
            dataToPrint = True
        End If

        If chkEDT824.Checked AndAlso dst.Tables("EDT824T1").Rows.Count > 0 Then
            dataToPrint = True
        End If

        If chkEDT860.Checked AndAlso dst.Tables("EDT860T1").Rows.Count > 0 Then
            dataToPrint = True
        End If

        If chkEDT864.Checked AndAlso dst.Tables("EDT864T1").Rows.Count > 0 Then
            dataToPrint = True
        End If

        If Not dataToPrint Then
            RWU = "0"
            xErrMsg = "No Eligible Records"
        End If

    End Sub

    Public Overrides Sub Print_Report()

        If chkEDT816.Checked Then
            RPT_TITLE = "816 - Store Address Changes"
            'Generate_Report("EDR816R1", RPT_TITLE, String.Empty)
            Generate_Report("EDR816R2", RPT_TITLE, String.Empty)
        End If

        If chkEDT824.Checked Then
            RPT_TITLE = "824 - Advice"
            Generate_Report("EDR824R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT860.Checked Then
            RPT_TITLE = "860 - PO Changes"
            Generate_Report("EDR860R1", RPT_TITLE, String.Empty)
        End If

        If chkEDT864.Checked Then
            RPT_TITLE = "864 - Messages"
            Generate_Report("EDR864R1", RPT_TITLE, String.Empty)
        End If

    End Sub

    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst

            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            Create_TDA(.Tables.Add, "ICTITEM1", "*")

            If enable816 Then
                Create_TDA(.Tables.Add, "EDT816T1", "*")
                Create_TDA(.Tables.Add, "EDT816T2", "*")
                Create_TDA(.Tables.Add, "EDT816T3", "*")
            End If

            If enable824 Then
                Create_TDA(.Tables.Add, "EDT824T1", "*")
                Create_TDA(.Tables.Add, "EDT824T2", "*")
                Create_TDA(.Tables.Add, "EDT824T3", "*")
                Create_TDA(.Tables.Add, "EDT824T4", "*")
            End If

            If enable860 Then
                Create_TDA(.Tables.Add, "EDT860T1", "*")
                Create_TDA(.Tables.Add, "EDT860T2", "*")
                .Tables("EDT860T2").Columns.Add("PO_CHANGE_TYPE", GetType(System.String))
                Create_TDA(.Tables.Add, "EDT860T3", "*")
                Create_TDA(.Tables.Add, "EDT860T4", "*")
            End If

            If enable864 Then
                Create_TDA(.Tables.Add, "EDT864T1", "*")
                '.Tables("EDT864T1").Columns.Add("CUST_CODE", GetType(System.String))
                Create_TDA(.Tables.Add, "EDT864T2", "*")
                Create_TDA(.Tables.Add, "EDT864T3", "*")
            End If

            Create_Lookup("ICTITEM1", "*", "ITEM_UPC_CODE = :PARM1", "V", False)
            Create_Lookup("ARTCUST1")
            Create_Lookup("ARTCUST2")

            sql = "Select CUST_CODE, LPAD(' ', 5) MAINT_TYPE ,CUST_STORE_NO,CUST_STORE_NAME,CUST_STORE_ADDR1,CUST_STORE_ADDR2,CUST_STORE_CITY,"
            sql &= " CUST_STORE_STATE, CUST_STORE_ZIP_CODE, CUST_DC_NO,"
            sql &= " CUST_STORE_NAME CUST_STORE_NAME_OLD, CUST_STORE_ADDR1 CUST_STORE_ADDR1_OLD, "
            sql &= " CUST_STORE_ADDR2 CUST_STORE_ADDR2_OLD, CUST_STORE_CITY CUST_STORE_CITY_OLD, "
            sql &= " CUST_STORE_STATE CUST_STORE_STATE_OLD, CUST_STORE_ZIP_CODE CUST_STORE_ZIP_CODE_OLD, "
            sql &= " CUST_DC_NO CUST_DC_NO_OLD "
            sql &= " FROM ARTCUST2 WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "EDT816TR", sql, 0, False, "", 0)


        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        sql = "Select * from EDTTRPM1"
        Fill_Records("EDTTRPM1", "", True, sql)


        If chkEDT816.Checked Then
            Load816()
        End If

        If chkEDT824.Checked Then
            Load824()
        End If

        If chkEDT860.Checked Then
            Load860()
        End If

        If chkEDT864.Checked Then
            Load864()
        End If

        EnforceConstraints(True)

        RWU = "N"
        If (enable816 AndAlso dst.Tables("EDT816T1").Rows.Count > 0) _
            OrElse (enable860 AndAlso dst.Tables("EDT860T1").Rows.Count > 0) _
            OrElse (enable864 AndAlso dst.Tables("EDT864T1").Rows.Count > 0) _
            OrElse (enable824 AndAlso dst.Tables("EDT824T1").Rows.Count > 0) Then
            RWU = "U"
        End If

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Public Overrides Sub Update_Record()
        MyBase.Update_Record()
        Dim SQL As String = String.Empty

        If chkEDT816.Checked AndAlso EDI_DOC_SEQ_NO_816.Length > 0 Then
            SQL = "UPDATE EDT816T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT824.Checked AndAlso EDI_DOC_SEQ_NO_824.Length > 0 Then
            SQL = "UPDATE EDT824T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT860.Checked AndAlso EDI_DOC_SEQ_NO_860.Length > 0 Then
            SQL = "UPDATE EDT860T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

        If chkEDT864.Checked AndAlso EDI_DOC_SEQ_NO_864.Length > 0 Then
            SQL = "UPDATE EDT864T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
            ASCDATA1.ExecuteSQL(SQL)
        End If

    End Sub

    Private Sub Load860()

        Dim sql As String = String.Empty
        Dim rowICTITEM1 As DataRow = Nothing
        EDI_DOC_SEQ_NO_860 = String.Empty

        For Each TABLE_NAME As String In New String() {"EDT860T1", "EDT860T2", "EDT860T3", "EDT860T4"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ASCMAIN1.Progress("Processing 860's", String.Empty)
        Try

            ASCDATA1.ExecuteSQL("Update EDT860T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT860T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT860T1.EDI_OUR_ID and EDI_TP_ID = EDT860T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT860T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT860T1.EDI_TP_QUAL and EDI_TP_ID = EDT860T1.EDI_TP_ID and EDI_DOC_NO = 860)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            ASCDATA1.ExecuteSQL()

            ' Type KC, KB means Contract and are viewed/processed on another screen.
            sql = "Select * from EDT860T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' and EDI_PO_TYPE NOT IN ('KC', 'KB')"
            Fill_Records("EDT860T1", String.Empty, True, sql)

            For Each rowEDT860T1 As DataRow In dst.Tables("EDT860T1").Rows
                EDI_DOC_SEQ_NO_860 &= ", '" & rowEDT860T1.Item("EDI_DOC_SEQ_NO") & "'"
            Next

            If EDI_DOC_SEQ_NO_860.Length > 0 Then
                EDI_DOC_SEQ_NO_860 = EDI_DOC_SEQ_NO_860.Substring(2).Trim
                sql = "Select * from EDT860T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T2", String.Empty, True, sql)

                sql = "Select * from EDT860T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T3", String.Empty, True, sql)

                sql = "Select * from EDT860T4 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_860 & ")"
                Fill_Records("EDT860T4", String.Empty, True, sql)

            End If

            For Each rowEDT860T2 As DataRow In dst.Tables("EDT860T2").Select("")
                Dim EDI_UPC As String = rowEDT860T2.Item("EDI_UPC") & String.Empty
                If EDI_UPC.Length = 0 Then
                    EDI_UPC = rowEDT860T2.Item("EDI_EAN") & String.Empty
                End If

                ' Needs to print EDT860T2.EDI_SKU on the report as Costco does not transmit EDI_ITEM
                rowICTITEM1 = Nothing
                If EDI_UPC.Length > 0 Then
                    rowICTITEM1 = LookUp("ICTITEM1", New String() {EDI_UPC})
                End If

                If rowICTITEM1 Is Nothing AndAlso EDI_UPC.Length = 0 Then
                    rowEDT860T2.Item("EDI_ITEM") = rowEDT860T2.Item("EDI_SKU") & String.Empty
                End If

                Select Case rowEDT860T2.Item("EDI_CHANGE_TYPE") & String.Empty
                    Case "QI", "QD"
                        If rowEDT860T2.Item("EDI_CHANGE_TYPE") = "QD" Then
                            rowEDT860T2.Item("EDI_QTY_CHANGE") = Val(rowEDT860T2.Item("EDI_QTY_CHANGE") & String.Empty) * -1
                        End If
                    Case Else
                        rowEDT860T2.Item("EDI_QTY_CHANGE") = rowEDT860T2.Item("EDI_QTY_OPEN")
                End Select

                If rowICTITEM1 IsNot Nothing Then
                    rowEDT860T2.Item("EDI_ITEM") = rowICTITEM1.Item("ITEM_CODE")
                End If

            Next

            ' Create dummy records for T2 where it does not exist for a T1
            Dim EDI_DOC_SEQ_NO As String = String.Empty
            For Each rowEDT860T1 As DataRow In dst.Tables("EDT860T1").Select
                EDI_DOC_SEQ_NO = rowEDT860T1.Item("EDI_DOC_SEQ_NO") & String.Empty
                If dst.Tables("EDT860T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 0 Then
                    Dim rowEDT860T2 As DataRow = dst.Tables("EDT860T2").NewRow
                    rowEDT860T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowEDT860T2.Item("EDI_ORIG_DTL_SEQ") = -9
                    dst.Tables("EDT860T2").Rows.Add(rowEDT860T2)
                End If

                If dst.Tables("EDT860T3").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 0 Then
                    Dim rowEDT860T3 As DataRow = dst.Tables("EDT860T3").NewRow
                    rowEDT860T3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowEDT860T3.Item("EDI_ORIG_DTL_SEQ") = -9
                    rowEDT860T3.Item("EDI_SDQ_SEQ") = 1
                    dst.Tables("EDT860T3").Rows.Add(rowEDT860T3)
                End If
            Next

            ' Combine Comments is section T4
            EDI_DOC_SEQ_NO = String.Empty
            Dim EDI_CMT_SEQ As Int32 = 999
            Dim EDI_CMT_REF_QUAL As String = String.Empty
            Dim EDI_CMT_REF As String = String.Empty
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("EDT860T4"), New String() {"EDI_DOC_SEQ_NO", "EDI_CMT_REF_QUAL", "EDI_CMT_REF"}).Select

                If EDI_DOC_SEQ_NO <> row.Item("EDI_DOC_SEQ_NO") & String.Empty _
                    OrElse EDI_CMT_REF_QUAL <> row.Item("EDI_CMT_REF_QUAL") & String.Empty _
                    OrElse EDI_CMT_REF <> row.Item("EDI_CMT_REF") & String.Empty Then
                    EDI_CMT_SEQ += 1
                End If
                EDI_DOC_SEQ_NO = row.Item("EDI_DOC_SEQ_NO") & String.Empty
                EDI_CMT_REF_QUAL = row.Item("EDI_CMT_REF_QUAL") & String.Empty
                EDI_CMT_REF = row.Item("EDI_CMT_REF") & String.Empty

                Dim EDI_CMMNT As String = String.Empty
                Dim EDI_CMMNT_2 As String = String.Empty

                For Each rowEDT860T4 As DataRow In dst.Tables("EDT860T4").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_CMT_REF_QUAL = '" & EDI_CMT_REF_QUAL & "' and EDI_CMT_REF_QUAL = '" & EDI_CMT_REF_QUAL & "'", "EDI_CMT_SEQ")
                    EDI_CMMNT &= " " & rowEDT860T4.Item("EDI_CMMNT")
                    EDI_CMMNT_2 &= " " & rowEDT860T4.Item("EDI_CMMNT_2")
                Next

                While EDI_CMMNT.Contains("  ")
                    EDI_CMMNT = EDI_CMMNT.Replace("  ", " ")
                End While

                While EDI_CMMNT_2.Contains("  ")
                    EDI_CMMNT_2 = EDI_CMMNT_2.Replace("  ", " ")
                End While

                EDI_CMMNT = EDI_CMMNT.Trim
                EDI_CMMNT_2 = EDI_CMMNT_2.Trim

                If EDI_CMMNT.Length > 0 OrElse EDI_CMMNT_2.Length > 0 Then
                    Dim rowEDT860T4x As DataRow = dst.Tables("EDT860T4").NewRow
                    rowEDT860T4x.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowEDT860T4x.Item("EDI_CMT_SEQ") = EDI_CMT_SEQ
                    rowEDT860T4x.Item("EDI_CMMNT") = EDI_CMMNT
                    rowEDT860T4x.Item("EDI_CMMNT_2") = EDI_CMMNT_2
                    rowEDT860T4x.Item("EDI_CMT_REF") = EDI_CMT_REF
                    rowEDT860T4x.Item("EDI_CMT_REF_QUAL") = EDI_CMT_REF_QUAL
                    dst.Tables("EDT860T4").Rows.Add(rowEDT860T4x)
                End If
            Next

            Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT COUNT(*) NUM_RECORDS FROM EDT860T1 WHERE COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' AND NVL(EDI_PROCESS_IND,'0') = '0' and EDI_PO_TYPE IN ('KC', 'KB')")
            Dim numRecords As Int16 = Val(tbl.Rows(0).Item("NUM_RECORDS") & String.Empty)
            If numRecords > 0 Then
                MessageBox.Show("There are " & numRecords & " transactions that need Acknowledgment.", "860 Ack", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 860's: " & ex.Message, "EDI 860", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Load816()
        Dim sql As String = String.Empty
        EDI_DOC_SEQ_NO_816 = String.Empty
        Dim rowARTCUST2 As DataRow = Nothing

        ASCMAIN1.Progress("Processing 816's", String.Empty)

        Try

            ASCDATA1.ExecuteSQL("Update EDT816T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT816T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT816T1.EDI_OUR_ID and EDI_TP_ID = EDT816T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT816T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT816T1.EDI_TP_QUAL and EDI_TP_ID = EDT816T1.EDI_TP_ID and EDI_DOC_NO = 816)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            ASCDATA1.ExecuteSQL()


            For Each TABLE_NAME As String In New String() {"EDT816T1", "EDT816T2", "EDT816T3"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "Select * from EDT816T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            Fill_Records("EDT816T1", String.Empty, True, sql)

            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Rows

                EDI_DOC_SEQ_NO_816 &= ", '" & rowEDT816T1.Item("EDI_DOC_SEQ_NO") & "'"

                Dim EDI_TP_ID As String = rowEDT816T1.Item("EDI_TP_ID") & String.Empty
                If dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                    rowEDT816T1.Item("CUST_CODE") = dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'")(0).Item("CUST_CODE")
                End If
            Next

            If EDI_DOC_SEQ_NO_816.Length > 0 Then
                EDI_DOC_SEQ_NO_816 = EDI_DOC_SEQ_NO_816.Substring(1).Trim
                sql = "Select * from EDT816T2 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
                Fill_Records("EDT816T2", String.Empty, True, sql)

                sql = "Select * from EDT816T3 where EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_816 & ")"
                Fill_Records("EDT816T3", String.Empty, True, sql)
            End If

            For Each rowEDT816T1 As DataRow In dst.Tables("EDT816T1").Rows
                Dim rowEDTTRPM1 As DataRow = Nothing
                Dim EDI_TP_ID As String = rowEDT816T1.Item("EDI_TP_ID") & String.Empty
                EDI_TP_ID = EDI_TP_ID.Trim

                If dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                    rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("EDI_DOC_NO = '816' AND EDI_TP_ID = '" & EDI_TP_ID & "'")(0)
                End If

                If rowEDTTRPM1 IsNot Nothing Then
                    Dim CUST_CODE As String = rowEDTTRPM1.Item("CUST_CODE") & String.Empty
                    For Each rowEDT816T2 As DataRow In dst.Tables("EDT816T2").Select("EDI_DOC_SEQ_NO = '" & rowEDT816T1.Item("EDI_DOC_SEQ_NO") & "' AND CUST_ADDR_CODE <> 'CQ'")
                        Dim CUST_STORE_NO As String = rowEDT816T2.Item("EDI_ADDR_CODE") & String.Empty
                        CUST_STORE_NO = ASCMAIN1.Format_Field(CUST_STORE_NO, "CUST_STORE_NO")

                        Dim rowEDT816TR As DataRow = dst.Tables("EDT816TR").NewRow
                        rowEDT816TR.Item("CUST_CODE") = rowEDTTRPM1.Item("CUST_CODE") & ""
                        rowEDT816TR.Item("MAINT_TYPE") = rowEDT816T2.Item("MAINT_TYPE") & ""
                        rowEDT816TR.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowEDT816TR.Item("CUST_STORE_NAME") = rowEDT816T2.Item("CUST_NAME") & ""
                        rowEDT816TR.Item("CUST_STORE_ADDR1") = rowEDT816T2.Item("CUST_ADDR1") & ""
                        rowEDT816TR.Item("CUST_STORE_ADDR2") = rowEDT816T2.Item("CUST_ADDR2") & ""
                        rowEDT816TR.Item("CUST_STORE_ZIP_CODE") = rowEDT816T2.Item("CUST_ZIP_CODE") & ""
                        rowEDT816TR.Item("CUST_STORE_CITY") = rowEDT816T2.Item("CUST_CITY") & ""
                        rowEDT816TR.Item("CUST_STORE_STATE") = rowEDT816T2.Item("CUST_STATE") & ""
                        rowEDT816TR.Item("CUST_DC_NO") = rowEDT816T2.Item("CUST_DC_CODE") & ""

                        rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        If rowARTCUST2 IsNot Nothing Then
                            rowEDT816TR.Item("CUST_STORE_NAME_OLD") = rowARTCUST2.Item("CUST_STORE_NAME") & ""
                            rowEDT816TR.Item("CUST_STORE_ADDR1_OLD") = rowARTCUST2.Item("CUST_STORE_ADDR1") & ""
                            rowEDT816TR.Item("CUST_STORE_ADDR2_OLD") = rowARTCUST2.Item("CUST_STORE_ADDR2") & ""
                            rowEDT816TR.Item("CUST_STORE_ZIP_CODE_OLD") = rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & ""
                            rowEDT816TR.Item("CUST_STORE_CITY_OLD") = rowARTCUST2.Item("CUST_STORE_CITY") & ""
                            rowEDT816TR.Item("CUST_STORE_STATE_OLD") = rowARTCUST2.Item("CUST_STORE_STATE") & ""
                            rowEDT816TR.Item("CUST_DC_NO_OLD") = rowARTCUST2.Item("CUST_DC_NO") & ""
                        End If

                        dst.Tables("EDT816TR").Rows.Add(rowEDT816TR)
                    Next
                End If

            Next
        Catch ex As Exception
            MessageBox.Show("Error processing 816's: " & ex.Message, "EDI 816", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Load864()
        EDI_DOC_SEQ_NO_864 = String.Empty

        ASCMAIN1.Progress("Processing 864's", String.Empty)

        Try

            ASCDATA1.ExecuteSQL("Update EDT864T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID), EDI_OUR_ID = TRIM(EDI_OUR_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update EDT864T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = EDT864T1.EDI_OUR_ID and EDI_TP_ID = EDT864T1.EDI_TP_ID)" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update EDT864T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = EDT864T1.EDI_TP_QUAL and EDI_TP_ID = EDT864T1.EDI_TP_ID and EDI_DOC_NO = 864)" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() {"EDT864T1", "EDT864T2", "EDT864T3"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "SELECT EDT864T1.* FROM EDT864T1 where NVL(EDI_PROCESS_IND, '0') = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            Fill_Records("EDT864T1", String.Empty, True, sql)

            For Each rowEDT864T1 As DataRow In dst.Tables("EDT864T1").Select()
                EDI_DOC_SEQ_NO_864 = EDI_DOC_SEQ_NO_864 & ", '" & rowEDT864T1.Item("EDI_DOC_SEQ_NO") & "'"
            Next

            If EDI_DOC_SEQ_NO_864.Length > 0 Then
                EDI_DOC_SEQ_NO_864 = EDI_DOC_SEQ_NO_864.Substring(2).Trim

                sql = "SELECT * from EDT864T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
                Fill_Records("EDT864T2", String.Empty, True, sql)

                sql = "SELECT * from EDT864T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_864 & ")"
                Fill_Records("EDT864T3", String.Empty, True, sql)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 864's: " & ex.Message, "EDI 864", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Load824()

        EDI_DOC_SEQ_NO_824 = String.Empty

        ASCMAIN1.Progress("Processing 824's", String.Empty)

        Try
            For Each TABLE_NAME As String In New String() {"EDT824T1", "EDT824T2", "EDT824T3", "EDT824T4"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            sql = "SELECT EDT824T1.*, '' CUST_CODE FROM GEN.EDT824T1 where NVL(EDI_PROCESS_IND, '0') = '0' "
            Fill_Records("EDT824T1", String.Empty, True, sql)

            For Each rowEDT824T1 As DataRow In dst.Tables("EDT824T1").Select()
                Dim EDI_TP_ID As String = (rowEDT824T1.Item("EDI_TP_ID") & String.Empty).ToString.Trim
                Dim EDI_TP_QUAL As String = (rowEDT824T1.Item("EDI_TP_QUAL") & String.Empty).ToString.Trim

                sql = "EDI_DOC_NO = '824' AND EDI_TP_ID = '" & EDI_TP_ID & "' and EDI_TP_QUAL = '" & EDI_TP_QUAL & "'"
                If dst.Tables("EDTTRPM1").Select(sql).Length > 0 Then
                    rowEDT824T1.Item("CUST_CODE") = dst.Tables("EDTTRPM1").Select(sql)(0).Item("CUST_CODE")
                End If

                rowEDT824T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                EDI_DOC_SEQ_NO_824 = EDI_DOC_SEQ_NO_824 & ", '" & rowEDT824T1.Item("EDI_DOC_SEQ_NO") & "'"
            Next

            If EDI_DOC_SEQ_NO_824.Length > 0 Then
                EDI_DOC_SEQ_NO_824 = EDI_DOC_SEQ_NO_824.Substring(2).Trim

                sql = "SELECT * from EDT824T2 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T2", String.Empty, True, sql)

                sql = "SELECT * from EDT824T3 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T3", String.Empty, True, sql)

                sql = "SELECT * from EDT824T4 WHERE EDI_DOC_SEQ_NO IN (" & EDI_DOC_SEQ_NO_824 & ")"
                Fill_Records("EDT824T4", String.Empty, True, sql)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing 824's: " & ex.Message, "EDI 864", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Class