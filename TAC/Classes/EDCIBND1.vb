Public Class EDCIBND1
    Dim MAP1() As strEDTMAPT1
    Dim SEGs As New Dictionary(Of String, strEDTSEGF1())

    Dim rowAtLevel() As DataRow
    Dim STANDARDS_ID As String
    Dim EDI_DOC_NO As String
    Dim SEGMENT_ID As String
    Dim LVLmax As Integer
    Dim LVLrec As Integer

    Public EDI_JRNL_NO As String

    Private frm As New ASCBASE1

    Function Process_File( _
    ByVal EDI_FOLDERNAME As String, _
    ByVal EDI_FILENAME As String, _
    ByVal FILEINFO As System.IO.FileInfo, _
    ByVal EDI_DOC_NO_to_process As String, _
    ByVal FILENAME_archive As String, _
    ByVal INIT_DATE As String, _
    ByVal dup_doc As Boolean) As List(Of String)

        ' transactions
        ' error handling - LIKE DO NOT HAVE STDS SET
        ' CATCH DUPS

        Dim EDI_JRNL_NOs As New List(Of String)
        Dim Delimited As Boolean = True ' there is an unfinished attempt to deal with files that are not delimited (ie, fixed length records)
        Dim FILENAME As String = EDI_FOLDERNAME & "\" & EDI_FILENAME

        Dim raw As String = ""
        Using sr As New System.IO.StreamReader(FILENAME)
            raw = sr.ReadToEnd
        End Using

        Dim EDI_FILE_DATETIME As Date = My.Computer.FileSystem.GetFileInfo(FILENAME).LastWriteTime

        ' note - this method assumes that if there are > 1 ISA / file
        ' that each ISA will be preceded by a vbCrLf

        Dim edi() As String
        Dim ISA As String = ""
        Dim ISAs() As String = Split(vbCrLf & raw, vbCrLf & "ISA")

        For iISA As Integer = 1 To ISAs.Length - 1
            Dim ISAblob As String = "ISA" & ISAs(iISA)
            Dim RS As String = Mid(ISAblob, 106, 1)
            If Mid(ISAblob, 107, 1) = vbLf Then
                RS &= vbLf
            End If
            If Mid(ISAblob, 107, 2) = vbCrLf Then
                RS &= vbCrLf
            End If
            Dim FS As String = Mid(ISAblob, 104, 1)

            Dim GSs() As String = Split(ISAblob, RS & "GS")
            Dim EDI_ISA_RECORD As String = GSs(0)
            edi = Split(EDI_ISA_RECORD, FS)

            Dim rowEDTJRNL1 As DataRow = frm.dst.Tables("EDTJRNL1").NewRow
            EDI_JRNL_NO = ASCMAIN1.Next_Control_No("EDTJRNL1.JOURNAL_NO")
            EDI_JRNL_NOs.Add(EDI_JRNL_NO)

            rowEDTJRNL1.Item("EDI_JRNL_NO") = EDI_JRNL_NO
            rowEDTJRNL1.Item("EDI_JRNL_DATE") = Now + ASCMAIN1.NowTSD
            rowEDTJRNL1.Item("EDI_FILENAME") = EDI_FILENAME
            rowEDTJRNL1.Item("EDI_ISA_NO") = edi(13) ' EDI_ISA_CTL_NO
            Dim Provider As New System.Globalization.DateTimeFormatInfo()
            rowEDTJRNL1.Item("EDI_ISA_DATE") = DateTime.ParseExact(edi(9) & edi(10), "yyMMddHHmm", Provider) ' EDI_ISA_CTL_DATE
            rowEDTJRNL1.Item("EDI_TP_QUAL") = edi(5) ' EDI_SENDER_QUAL
            rowEDTJRNL1.Item("EDI_TP_ID") = edi(6) ' EDI_SENDER_ID
            rowEDTJRNL1.Item("EDI_OUR_QUAL") = edi(7) ' EDI_RECEIVER_QUAL
            rowEDTJRNL1.Item("EDI_OUR_ID") = edi(8) ' EDI_RECEIVER_ID
            rowEDTJRNL1.Item("EDI_ISA_RECORD") = EDI_ISA_RECORD
            rowEDTJRNL1.Item("EDI_FOLDERNAME") = EDI_FOLDERNAME
            rowEDTJRNL1.Item("EDI_DATETIME") = FILEINFO.LastWriteTime
            rowEDTJRNL1.Item("EDI_FILESIZE") = FILEINFO.Length
            frm.dst.Tables("EDTJRNL1").Rows.Add(rowEDTJRNL1)

            For iGS As Integer = 1 To GSs.Length - 1
                Dim GSblob As String = "GS" & GSs(iGS)
                Dim STs() As String = Split(GSblob, RS & "ST")
                Dim EDI_GS_RECORD As String = STs(0)
                edi = Split(EDI_GS_RECORD, FS)

                STANDARDS_ID = Trim(edi(8))

                'GS*FA*6125310421*2122194288*20070126*1254*006272412*X*004010VICS
                Dim rowEDTJRNL2 As DataRow = frm.dst.Tables("EDTJRNL2").NewRow
                rowEDTJRNL2.Item("EDI_JRNL_NO") = EDI_JRNL_NO
                rowEDTJRNL2.Item("EDI_GS_NO") = iGS
                rowEDTJRNL2.Item("EDI_GS_RECORD") = EDI_GS_RECORD
                rowEDTJRNL2.Item("EDI_GS_FUNC_ID") = edi(1)
                rowEDTJRNL2.Item("EDI_GS_SENDER_ID") = edi(2)
                rowEDTJRNL2.Item("EDI_GS_RECEIVER_ID") = edi(3)
                rowEDTJRNL2.Item("EDI_GS_DATE") = edi(4)
                rowEDTJRNL2.Item("EDI_GS_TIME") = edi(5)
                rowEDTJRNL2.Item("EDI_GS_CTL_NO") = edi(6)
                rowEDTJRNL2.Item("EDI_GS_AGENCY") = edi(7)
                rowEDTJRNL2.Item("EDI_GS_STANDARDS_ID") = edi(8)

                frm.dst.Tables("EDTJRNL2").Rows.Add(rowEDTJRNL2)

                For iST As Integer = 1 To STs.Length - 1
                    Dim STblob As String = "ST" & STs(iST)
                    Dim DOC() As String = Split(STblob, RS)
                    Dim EDI_ST_RECORD As String = DOC(0)
                    edi = Split(EDI_ST_RECORD, FS)

                    EDI_DOC_NO = edi(1)

                    Dim rowEDTJRNL3 As DataRow = frm.dst.Tables("EDTJRNL3").NewRow
                    rowEDTJRNL3.Item("EDI_JRNL_NO") = EDI_JRNL_NO
                    rowEDTJRNL3.Item("EDI_GS_NO") = iGS
                    rowEDTJRNL3.Item("EDI_ST_NO") = iST
                    rowEDTJRNL3.Item("EDI_DOC_NO") = edi(1)
                    rowEDTJRNL3.Item("EDI_ST_RECORD") = EDI_ST_RECORD
                    rowEDTJRNL3.Item("EDI_ST_CTL_NO") = edi(2)
                    frm.dst.Tables("EDTJRNL3").Rows.Add(rowEDTJRNL3)

                    If EDI_DOC_NO_to_process = EDI_DOC_NO Then
                        Dim EDI_DOC_SEQ_NO As String = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO")
                        rowEDTJRNL3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                        If Not Prepare_for_Mapping() Then
                            ' raise error
                            Stop
                        Else
                            LVLrec = 0

                            Dim LVLlno() As Integer
                            ReDim LVLlno(LVLmax)

                            Dim reccount As Integer = 0
                            Dim last_SEGMENT_ID As String = ""

                            For iDOC As Integer = 1 To DOC.Length - 1
                                Dim line As String = DOC(iDOC)

                                If InStr(line, "REFZZ") = 1 Then ' Adjust for ADS custom segment
                                    line = "RZZ" & Mid$(line, 6)
                                End If

                                'If line.StartsWith("N1*ST*") Then
                                '    Stop
                                'End If

                                edi = Split(line, FS)

                                reccount += 1
                                If reccount Mod 1000 = 0 Then
                                    Call ASCMAIN1.Progress("-", CStr(reccount))
                                    Application.DoEvents()
                                End If

                                SEGMENT_ID = edi(0)
                                If SEGs.ContainsKey(SEGMENT_ID) Then

                                    Dim SEG_FLDs() As strEDTSEGF1 = SEGs(SEGMENT_ID)
                                    Dim SEGMENT As strEDTSEGF1 = SEG_FLDs(1)

                                    Dim QUALIFIER As String = ""
                                    If SEGMENT.QUALIFIER_FLAG = "1" Then
                                        If Delimited Then
                                            QUALIFIER = edi(1)
                                        Else
                                            QUALIFIER = Mid$(line, 4, SEGMENT.ELEM_LENGTH)
                                        End If
                                    End If

                                    If LVLrec = 0 Then
                                        LVLrec = 1

                                        rowAtLevel(LVLrec) = rowAtLevel(LVLrec).Table.NewRow
                                        rowAtLevel(LVLrec).Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                        If LVLrec > 1 Then
                                            For iLVL As Integer = 2 To LVLrec
                                                If iLVL = LVLrec Then
                                                    LVLlno(LVLrec) += 1
                                                End If
                                                rowAtLevel(LVLrec).Item(iLVL - 1) = LVLlno(iLVL)
                                            Next
                                        End If
                                        rowAtLevel(LVLrec).Table.Rows.Add(rowAtLevel(LVLrec))

                                    End If





                                    For i As Integer = 1 To LVLmax
                                        If MAP1(i).LEAD_SEGMENT & ":" & MAP1(i).LEAD_SEGMENT_QUAL = _
                                           SEGMENT_ID & ":" & QUALIFIER Then

                                            'Stop
                                            last_SEGMENT_ID = ""

                                            Dim LVLnew As Integer = i
                                            If LVLnew < LVLrec Then

                                                For iLVL As Integer = LVLnew + 1 To LVLrec
                                                    LVLlno(iLVL) = 0
                                                Next
                                            End If

                                            LVLrec = LVLnew

                                            rowAtLevel(LVLrec) = rowAtLevel(LVLrec).Table.NewRow
                                            rowAtLevel(LVLrec).Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                            If LVLrec > 1 Then
                                                For iLVL As Integer = 2 To LVLrec
                                                    If iLVL = LVLrec Then
                                                        LVLlno(LVLrec) += 1
                                                    End If
                                                    rowAtLevel(LVLrec).Item(iLVL - 1) = LVLlno(iLVL)
                                                Next
                                            End If
                                            rowAtLevel(LVLrec).Table.Rows.Add(rowAtLevel(LVLrec))
                                            Exit For
                                        End If
                                    Next i

                                    ' NEED TO DEFINE A REPEATING SEGMENT, LIKE SDQ IN AN 852
                                    ' IF LAST SEGMENT = SEGMENT_ID AND "THIS IS A REPEATING SEGMENT" THEN
                                    ' NEW ROW

                                    If last_SEGMENT_ID = SEGMENT_ID And SEGMENT_ID = "SDQ" And LVLrec = LVLmax Then

                                        Dim rowSave As DataRow = rowAtLevel(LVLrec).Table.NewRow
                                        rowSave.ItemArray = rowAtLevel(LVLrec).ItemArray

                                        rowAtLevel(LVLrec) = rowAtLevel(LVLrec).Table.NewRow
                                        rowAtLevel(LVLrec).Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                        If LVLrec > 1 Then
                                            For iLVL As Integer = 2 To LVLrec
                                                If iLVL = LVLrec Then
                                                    LVLlno(LVLrec) += 1
                                                End If
                                                rowAtLevel(LVLrec).Item(iLVL - 1) = LVLlno(iLVL)
                                            Next
                                        End If
                                        rowAtLevel(LVLrec).Table.Rows.Add(rowAtLevel(LVLrec))

                                        ' AND NOW WE NEED TO PACK THE ZA RECORD FIELDS

                                        rowAtLevel(LVLrec).Item("EDI_ZA_TRAN_TYPE") = rowSave.Item("EDI_ZA_TRAN_TYPE")
                                    End If

                                    last_SEGMENT_ID = SEGMENT_ID




                                    Dim ELEM_POSITION As Integer = SEGMENT.ELEM_POSITION
                                    Dim ELEM_LENGTH As Integer = SEGMENT.ELEM_LENGTH

                                    For TABLE_LEVEL As Integer = 1 To LVLrec


                                        Dim TABLE_NAME As String = MAP1(TABLE_LEVEL).TABLE_NAME
                                        Dim sql As String = "STANDARDS_ID = '" & STANDARDS_ID & "'" _
                                            & " and EDI_DOC_NO = '" & EDI_DOC_NO & "'" _
                                            & " and SEGMENT_ID = '" & SEGMENT_ID & "'" _
                                            & " and TABLE_NAME = '" & TABLE_NAME & "'"

                                        ' i think this is not going to work for some documents where the qualifier is embedded in the line somewhere
                                        If QUALIFIER <> "" Then
                                            sql &= " and QUALIFIER_ID = '" & QUALIFIER & "'"
                                        End If

                                        Dim dvwEDTMAPT2 As New DataView(frm.dst.Tables("EDTMAPT2"), _
                                           sql, "", DataViewRowState.CurrentRows)

                                        For Each rowEDTMAPT2 As DataRowView In dvwEDTMAPT2
                                            Dim COLUMN_NAME As String = rowEDTMAPT2.Item("COLUMN_NAME") & ""
                                            Dim QUALIFIER_ID As String = rowEDTMAPT2.Item("QUALIFIER_ID") & ""
                                            Dim SEGMENT_FIELD_NO As Integer = rowEDTMAPT2.Item("SEGMENT_FIELD_NO") & ""

                                            Dim MAX_LENGTH As Integer = frm.dst.Tables(TABLE_NAME).Columns(COLUMN_NAME).MaxLength

                                            If SEGMENT_FIELD_NO = 0 Then SEGMENT_FIELD_NO = 1

                                            Dim zz As String = ""
                                            Dim SFN As Integer = 0

                                            If QUALIFIER_ID <> "" Then
                                                If SEGMENT.QUALIFIER_FLAG = "1" Then ' THE SEGMENT IS ALWAYS QUALIFIED
                                                    SFN = SEGMENT_FIELD_NO + 1
                                                Else ' GO AND FIND THE QUALIFER_ID IN ANY OF THE SEGMENT FIELDS
                                                    SFN = -1
                                                    For i As Integer = 1 To SEG_FLDs.Length - 1
                                                        If i >= edi.Length Then
                                                            Exit For
                                                        End If
                                                        Dim SEG_FLD As strEDTSEGF1 = SEG_FLDs(i)
                                                        If SEG_FLD.QUALIFIER_FLAG = "1" Then
                                                            If edi(i) = QUALIFIER_ID Then
                                                                SFN = i + 1
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            Else
                                                SFN = SEGMENT_FIELD_NO
                                            End If

                                            If SFN = -1 Then
                                            Else


                                                If Delimited Then
                                                    If UBound(edi) >= SFN Then
                                                        zz = edi(SFN)
                                                    Else
                                                        zz = ""
                                                    End If
                                                Else
                                                    Stop ' TEST THIS
                                                    If MAX_LENGTH < ELEM_LENGTH Then
                                                        zz = Mid$(line, 3 + ELEM_POSITION, MAX_LENGTH)
                                                    Else
                                                        zz = Mid$(line, 3 + ELEM_POSITION, ELEM_LENGTH)
                                                    End If
                                                End If

                                                Dim DataType As String = rowAtLevel(TABLE_LEVEL).Table.Columns(COLUMN_NAME).DataType.ToString
                                                Select Case DataType
                                                    Case "System.DateTime"  ' Date
                                                        If Trim$(zz) <> "" Then
                                                            If Len(zz) = 8 Then
                                                                rowAtLevel(TABLE_LEVEL).Item(COLUMN_NAME) = Mid$(zz, 5, 2) & "/" & Mid$(zz, 7, 2) & "/" & Mid$(zz, 1, 4)
                                                            Else
                                                                If Len(zz) >= 6 Then ' WHY >= 6 AND NOT = 6
                                                                    rowAtLevel(TABLE_LEVEL).Item(COLUMN_NAME) = Mid$(zz, 3, 2) & "/" & Mid$(zz, 5, 2) & "/" & Mid$(zz, 1, 2)
                                                                End If
                                                            End If
                                                        End If

                                                    Case "System.Int32", "System.Int16", "System.Double", "System.Decimal"
                                                        Dim multiplier As Integer
                                                        If SEGMENT.ELEM_TYPE = "N" Then
                                                            multiplier = 10 ^ SEGMENT.ELEM_LENGTH
                                                        Else
                                                            multiplier = 1
                                                        End If
                                                        rowAtLevel(TABLE_LEVEL).Item(COLUMN_NAME) = Val(zz) / multiplier

                                                    Case Else
                                                        rowAtLevel(TABLE_LEVEL).Item(COLUMN_NAME) = zz
                                                End Select
                                            End If
                                        Next
                                    Next
                                End If
                            Next

                            For iLVL As Integer = 1 To LVLmax
                                Dim TABLE_NAME As String = MAP1(iLVL).TABLE_NAME
                                frm.Update_Record_TDA(TABLE_NAME)
                            Next
                        End If
                    End If
                Next
            Next
        Next

        frm.Update_Record_TDA("EDTJRNL1")
        frm.Update_Record_TDA("EDTJRNL2")
        frm.Update_Record_TDA("EDTJRNL3")

        '        If dup = False Then
        '            If EDI_JRNL_NOx = "" Then
        '            GoSub End_of_EDI_JRNL_NO
        '            End If
        '        End If
        '        dup_doc = dup
        '        If EDI_JRNL_NOx = "" Then
        '            FileCopy(FILENAME, filename_archive)
        '            Kill(FILENAME)
        '        End If
        '        If EDI_JRNL_NOx = "" Then
        '            OraS.CommitTrans()
        '        End If
        '        Exit Sub

        Return EDI_JRNL_NOs
    End Function

    Function Prepare_for_Mapping() As Boolean
        Dim LVLs As Integer

        ' PROBABLY DON'T NEED TO DO MUCH OF THIS IF WE HAVE ALREADY SEEN THIS STD+DOC

        Dim KEY() As String = {STANDARDS_ID, EDI_DOC_NO}
        If frm.dst.Tables("EDTMAPT0").Rows.Find(KEY) Is Nothing Then
            frm.Fill_Records("EDTMAPT0", KEY, False)
            frm.Fill_Records("EDTMAPT1", KEY, False)
            frm.Fill_Records("EDTMAPT2", KEY, False)
            frm.dst.Tables("EDTSEGFD").PrimaryKey = Nothing
            frm.Fill_Records("EDTSEGFD", KEY, False)
        End If

        ReDim rowAtLevel(0)
        ReDim MAP1(0)

        LVLmax = 0
        LVLs = 1
        For Each rowEDTMAPT1 As DataRow In frm.dst.Tables("EDTMAPT1").Select _
        ("STANDARDS_ID = '" & STANDARDS_ID & "' and EDI_DOC_NO = '" & EDI_DOC_NO & "'", "TABLE_LEVEL")
            Dim TABLE_NAME As String = rowEDTMAPT1.Item("TABLE_NAME")
            If Not frm.dst.Tables.Contains(TABLE_NAME) Then
                frm.Create_TDA(frm.dst.Tables.Add, TABLE_NAME, "*")
            Else
                frm.dst.Tables(TABLE_NAME).Rows.Clear()
            End If

            ReDim Preserve rowAtLevel(LVLs)
            ReDim Preserve MAP1(LVLs)

            Dim row As DataRow = frm.dst.Tables(TABLE_NAME).NewRow
            rowAtLevel(LVLs) = row

            Dim M1 As strEDTMAPT1
            M1.TABLE_NAME = TABLE_NAME
            M1.LEAD_SEGMENT = rowEDTMAPT1.Item("LEAD_SEGMENT") & ""
            M1.LEAD_SEGMENT_QUAL = rowEDTMAPT1.Item("LEAD_SEGMENT_QUAL") & ""
            MAP1(LVLs) = M1

            LVLmax = LVLs
            LVLs = LVLs + 1
        Next

        If LVLmax = 0 Then
            Call ASCMAIN1.Progress("Could Not Determine Mapping for " & STANDARDS_ID & "/" & EDI_DOC_NO, "")
            Prepare_for_Mapping = False
            Stop
            Exit Function
        Else
            Prepare_for_Mapping = True
        End If

        Dim ELEM_POSITION As Integer = 1
        SEGs.Clear()

        For Each row As DataRow In ASCDATA1.SelectDistinct(frm.dst.Tables("EDTSEGFD").Select _
        ("STANDARDS_ID = '" & STANDARDS_ID & "' and EDI_DOC_NO = '" & EDI_DOC_NO & "'",
         "ELEM_SEQ_NO"), "SEGMENT_ID").Rows
            Dim SEGMENT_ID As String = row.Item("SEGMENT_ID")
            Dim SEG_FLDs() As strEDTSEGF1
            ReDim SEG_FLDs(0)
            For Each rowEDTSEGFD As DataRow In frm.dst.Tables("EDTSEGFD").Select _
            ("STANDARDS_ID = '" & STANDARDS_ID & "' and EDI_DOC_NO = '" & EDI_DOC_NO & "' and SEGMENT_ID = '" & SEGMENT_ID & "'",
             "ELEM_SEQ_NO")
                Dim ELEM_SEQ_NO As Integer = Val(rowEDTSEGFD.Item("ELEM_SEQ_NO") & "")
                Dim ELEM_LENGTH As Integer = Val(rowEDTSEGFD.Item("ELEM_LENGTH") & "")
                ReDim Preserve SEG_FLDs(ELEM_SEQ_NO)
                Dim SEG_FLD As strEDTSEGF1
                SEG_FLD.ELEM_DESC = rowEDTSEGFD.Item("ELEM_DESC") & ""
                SEG_FLD.ELEM_POSITION = ELEM_POSITION
                SEG_FLD.ELEM_LENGTH = ELEM_LENGTH
                ELEM_POSITION += ELEM_LENGTH
                SEG_FLD.ELEM_TYPE = rowEDTSEGFD.Item("ELEM_TYPE") & ""
                SEG_FLD.ELEM_REQ = rowEDTSEGFD.Item("ELEM_REQ") & ""
                SEG_FLD.QUALIFIER_FLAG = rowEDTSEGFD.Item("QUALIFIER_FLAG") & ""
                SEG_FLDs(ELEM_SEQ_NO) = SEG_FLD
            Next
            SEGs.Add(SEGMENT_ID, SEG_FLDs)
        Next

    End Function

    Public Sub New()
        'ByVal FF As ASFBASE1
        'frm = FF

        ' THIS CLASS ASSUMES THAT IF EDTJRNL1 IS NOT IN DST THEN ALL OF THE FOLLOWING TABLES NEED TO BE CREATED
        ' MAYBE WE DO A NEW INSTANCE OF ASFBASE1

        If Not frm.dst.Tables.Contains("EDTMAPT0") Then
            'ASCMAIN1.sql = "Select * from EDTJRNL1"
            frm.Create_TDA(frm.dst.Tables.Add, "EDTJRNL1", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "EDTJRNL2", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "EDTJRNL3", "*")

            frm.Create_TDA(frm.dst.Tables.Add, "EDTMAPT0", "*", 2)
            frm.Create_TDA(frm.dst.Tables.Add, "EDTMAPT1", "*", 2)
            frm.Create_TDA(frm.dst.Tables.Add, "EDTMAPT2", "*", 2)

            ASCMAIN1.sql = "Select EDTSEGF1.STANDARDS_ID, X.EDI_DOC_NO" _
            & ", EDTSEGF1.SEGMENT_ID, EDTSEGF1.ELEM_SEQ_NO" _
            & ", EDTSEGF1.ELEM_DESC, EDTSEGF1.ELEM_LENGTH " _
            & ", EDTSEGF1.ELEM_TYPE, EDTSEGF1.ELEM_REQ, EDTSEGF1.QUALIFIER_FLAG " _
            & " from EDTSEGF1, (" _
            & "  Select DISTINCT STANDARDS_ID, SEGMENT_ID, EDI_DOC_NO from " _
            & " (Select STANDARDS_ID, LEAD_SEGMENT SEGMENT_ID, EDI_DOC_NO from EDTMAPT1 " _
            & " where STANDARDS_ID = :PARM1 and EDI_DOC_NO = :PARM2) union " _
            & " (Select STANDARDS_ID, SEGMENT_ID, EDI_DOC_NO from EDTMAPT2 " _
            & " where STANDARDS_ID = :PARM1 and EDI_DOC_NO = :PARM2)" _
            & ") X " _
            & " where EDTSEGF1.STANDARDS_ID = X.STANDARDS_ID" _
            & " and EDTSEGF1.SEGMENT_ID = X.SEGMENT_ID"
            frm.Create_TDA(frm.dst.Tables.Add, "EDTSEGFD", "**", 0, False, "VV", 4)

        Else
            frm.dst.Tables("EDTJRNL1").Rows.Clear()
            frm.dst.Tables("EDTJRNL2").Rows.Clear()
            frm.dst.Tables("EDTJRNL3").Rows.Clear()
            frm.dst.Tables("EDTMAPT0").Rows.Clear()
            frm.dst.Tables("EDTMAPT1").Rows.Clear()
            frm.dst.Tables("EDTMAPT2").Rows.Clear()
            frm.dst.Tables("EDTSEGFD").Rows.Clear()
        End If
    End Sub

    Structure strEDTMAPT1
        Dim TABLE_NAME As String
        Dim LEAD_SEGMENT As String
        Dim LEAD_SEGMENT_QUAL As String

        'Public Function CheckValues() As Boolean
        '    Return True
        'End Function
    End Structure

    Structure strEDTSEGF1
        Dim ELEM_DESC As String
        Dim ELEM_POSITION As Integer
        Dim ELEM_LENGTH As Integer
        Dim ELEM_TYPE As String
        Dim ELEM_REQ As String
        Dim QUALIFIER_FLAG As String
    End Structure
End Class
