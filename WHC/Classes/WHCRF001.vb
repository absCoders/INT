Public Class WHCRF001
    ' Process Wave Instruction (Case Pick)

    Inherits WHCRF000

    Dim CUST_CODE As String
    Dim ORDR_CUST_PO As String
    Dim BAR_CODE As String
    Dim WAVE_INST_NO As String
    Dim WAVE_NO As String
    Dim LOCATION_CODE As String
    Dim CASES As Integer
    Dim WAVE_INST_TEXT As String

    Dim sqlDNA As String = ""
    Dim rowWHTINST1 As DataRow

    Sub New(g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF001"

        AppStates.Add("NEXT_INST", "Enter Wave No or Press Enter for Next|EXIT|")
        AppStates.Add("SCAN_LPN", "Scan Case ID to Pick|DONE|SHOW|EXIT|")
        AppStates.Add("REMOVE_LPN", "Remove Case? (Y/N)|Y|N|")
        AppStates.Add("SHORT_COUNT", "Continue with update? (Y/N)|Y|N|")
        AppStates.Add("VERIFY", "Are You Done (Y/N)|Y|N|CANCEL|")

        AppState = "NEXT_INST"

        With dst
            With .Tables.Add("WHTSCANS").Columns
                .Add("BAR_CODE")
                .Add("BAR_CODE_SCANNED")
                .Add("STYLE_COLOR_QTY_DNA")
            End With
            .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

            Create_TDA(.Tables.Add, "WHTINST2", "*")
            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")
        End With

        tbl = dst.Tables("WHTSCANS") ' New DataTable

        sqlDNA = "Select BAR_CODE," & vbCrLf _
            & " ltrim(sys_connect_by_path(SC ,','),',') STYLE_COLOR_QTY_DNA" & vbCrLf _
            & " from" & vbCrLf _
            & " (select BAR_CODE , STYLE_CODE || COLOR_CODE || TO_CHAR(LOCATION_QTY,'9999999') SC," & vbCrLf _
            & "    row_number() over(partition by BAR_CODE  order by STYLE_CODE || COLOR_CODE || TO_CHAR(LOCATION_QTY,'9999999') ) rn," & vbCrLf _
            & "       row_number() over(partition by BAR_CODE order by STYLE_CODE || COLOR_CODE || TO_CHAR(LOCATION_QTY,'9999999') desc)" & vbCrLf _
            & " rn_desc" & vbCrLf _
            & " from WHTLOCB1 where {sqlDNA_where})" & vbCrLf _
            & "    Where rn_desc = 1" & vbCrLf _
            & "  start with rn = 1" & vbCrLf _
            & "  connect by prior BAR_CODE = BAR_CODE" & vbCrLf _
            & "  and prior rn = rn-1"

    End Sub

    Public Overrides Function Hello() As String
        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        If AppState_initial = "" Then AppState_initial = AppState
        RESPONSE &= vbCrLf & vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "B", "EXIT")
        Else

            Select Case AppState

                Case "NEXT_INST"

                    Dim WAVE_INST_NO_preferred As String = ""
                    If SCANTEXT <> "" And SCANTEXT.StartsWith("I") Then
                        WAVE_INST_NO_preferred = Format(Val(Mid(SCANTEXT, 2)), "0000000000")
                        ASCMAIN1.sql = "Select WAVE_NO from WHTINST1 where WAVE_INST_NO = '" & WAVE_INST_NO_preferred & "'"
                        SCANTEXT = ASCDATA1.GetDataValue
                    End If

                    If SCANTEXT = "" And WAVE_NO <> "" Then
                        SCANTEXT = WAVE_NO
                    End If
                    If SCANTEXT = "" Then
                        CreateResponse("", "R", "You Need to Enter a Wave No")
                        Exit Select
                    End If

                    WAVE_NO = Format(Val(SCANTEXT), "0000000000")

                    ASCMAIN1.sql = "select wave_no from whtlocb1 l1 " _
                        & "join whtinst2 i2 on l1.bar_code = i2.bar_code " _
                        & "join whtinst1 i1 on i1.wave_inst_no = i2.wave_inst_no " _
                        & "where l1.LOCATION_CODE = '" & GunEnv.GUN_LOC & "' " _
                        & "and l1.location_qty > 0 " _
                        & "and rownum =1"

                    Dim WAVE_NO_OPEN As String = ASCDATA1.GetDataValue

                    'If WAVE_NO <> WAVE_NO_OPEN Then
                    '    ' Open Wave exists, they need to deposit or continue with previous wave
                    '    CreateResponse("", "R", "Wave " & WAVE_NO_OPEN & ", is Active, cannot pick wave " & WAVE_NO & " until deposit")
                    '    Exit Select
                    'End If

                    ASCMAIN1.sql = "Select WAVE_INST_NO from ( " & vbCrLf _
                            & "select WHTINST1.WAVE_INST_NO, WHTINST1.LOCATION_CODE" & vbCrLf _
                            & "from WHTINST1,WHTWAVE1" & vbCrLf _
                            & " where WHTINST1.WAVE_PICK_TYPE = '" & GunEnv.PICK_TYPE & "'" & vbCrLf _
                            & "   and WHTINST1.WAVE_INST_STATUS = '0'" & vbCrLf _
                            & "   and WHTWAVE1.WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" & vbCrLf _
                            & "   and WHTINST1.WAVE_NO = WHTWAVE1.WAVE_NO" & vbCrLf _
                            & "   and WHTINST1.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf _
                            & IIf(WAVE_INST_NO_preferred = "", "", "   and WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO_preferred & "'" & vbCrLf) _
                            & "   and WHTINST1.WAVE_INST_NO Not in (Select ENTITY from ASTMTSK2 where ENTITY_TYPE = 'WHTINST1')" & vbCrLf _
                            & "   order by location_code) " & vbCrLf _
                            & "Where RowNum = 1"
                    
                    WAVE_INST_NO = ASCDATA1.GetDataValue
                    If WAVE_INST_NO = "" Then
                        CreateResponse("", "R", "No Picks Available for Wave (" & SCANTEXT & ")")
                        Exit Select
                    End If

                    ASCMAIN1.sql = "Select case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else  SOTORDR0.CUST_CODE END CUST_CODE, " _
                            & "case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else SOTORDR0.ORDR_CUST_PO end ORDR_CUST_PO, " _
                            & " WHTINST1.WAVE_NO, WHTINST1.LOCATION_CODE " _
                            & "from WHTINST1 " _
                            & "join WHTWAVE1 on WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO " _
                            & "left join whtwave3 on WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO " _
                            & "left join SOTSHIP1 on SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO " _
                            & "left join SOTORDR0 on SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                            & " where WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                    Dim rowCUSTPO As DataRow = ASCDATA1.GetDataRow
                    CUST_CODE = rowCUSTPO.Item("CUST_CODE")
                    ORDR_CUST_PO = rowCUSTPO.Item("ORDR_CUST_PO") & ""
                    WAVE_NO = rowCUSTPO.Item("WAVE_NO")
                    LOCATION_CODE = rowCUSTPO.Item("LOCATION_CODE")

                    If Not ASCMAIN1.Logical_Open("WHTWAVE1", WAVE_NO) Then
                        CreateResponse("", "R", "Could Not Access Wave " & WAVE_NO)
                        Exit Select
                    End If
                    If Not ASCMAIN1.Logical_Lock("WHTINST1", WAVE_INST_NO) Then
                        CreateResponse("", "R", "Could Not Access Wave Instruction " & WAVE_INST_NO)
                        Exit Select
                    End If

                    rowWHTINST1 = LookUp("WHTINST1", WAVE_INST_NO)
                    If rowWHTINST1.Item("WAVE_INST_STATUS") & "" <> "0" Then
                        CreateResponse("", "R", "Wave Instruction " & WAVE_INST_NO & " no longer Open")
                        Exit Select
                    End If

                    tbl.Rows.Clear()
                    ASCMAIN1.sql = Replace(sqlDNA, "{sqlDNA_where}", "LOCATION_CODE = '" & LOCATION_CODE & "' and BAR_CODE in (Select Distinct BAR_CODE from WHTINST2 where WAVE_INST_NO = '" & WAVE_INST_NO & "')")
                    For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("", "BAR_CODE")
                        Dim row2 As DataRow = tbl.NewRow
                        row2.Item("BAR_CODE") = ROW.Item("BAR_CODE")
                        row2.Item("STYLE_COLOR_QTY_DNA") = ROW.Item("STYLE_COLOR_QTY_DNA")
                        tbl.Rows.Add(row2)
                    Next

                    For Each TABLE_NAME As String In New String() {"WHTINST2", "WHTLOCB1", "WHTBARC0"}
                        dst.Tables(TABLE_NAME).Rows.Clear()
                    Next

                    CASES = tbl.Rows.Count


                    ASCMAIN1.sql = "Select Distinct WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE, LOCATION_QTY_WAVE" & vbCrLf _
                        & " from WHTINST2,WHTINST1 " & vbCrLf _
                        & "    where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
                        & "      and WHTINST2.WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
                        & "      and WHTINST1.WAVE_INST_STATUS ='0'" & vbCrLf

                    Dim SCQ As String = ""
                    Dim STYLE_COLORs As Integer = 0
                    For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("") ' dst.Tables("WHTINST2").Select("")
                        Dim LOCATION_QTY_WAVE As Integer = Val(ROW.Item("LOCATION_QTY_WAVE") & "")
                        SCQ &= vbCrLf & ROW.Item("STYLE_CODE") & "-" & ROW.Item("COLOR_CODE") & ":" & CStr(LOCATION_QTY_WAVE)
                        STYLE_COLORs += 1
                    Next

                    WAVE_INST_TEXT = "Instruction " & WAVE_INST_NO _
                        & vbCrLf & CUST_CODE & ", PO " & ORDR_CUST_PO _
                        & vbCrLf & "Location: " & LOCATION_CODE _
                        & vbCrLf & "Cases: " & CStr(CASES) _
                        & SCQ ' IIf(STYLE_COLORs = 1, Split(SCQ, ":")(0), SCQ)

                    CreateResponse("SCAN_LPN", "B", WAVE_INST_TEXT, True)

                Case "SCAN_LPN"
                    If SCANTEXT = "DONE" Then
                        If tbl.Select("ISNULL(BAR_CODE_SCANNED,'') <> ''").Length < CASES Then
                            'SHORT_COUNT
                            CreateResponse("SHORT_COUNT", "R", "Missing Cases, ")
                            Exit Select
                        End If
                        CreateResponse("VERIFY", "B", "")
                    Else
                        If SCANTEXT = "SHOW" Then
                            Dim RESPONSE As String = "NEXT "
                            Dim rows As DataRow() = tbl.Select("")
                            If rows.Length = 0 Then
                                RESPONSE = "FINISHED"
                            Else
                                For Each row As DataRow In rows
                                    RESPONSE = RESPONSE & row.Item("BAR_CODE") & ","
                                Next
                                RESPONSE = Left(RESPONSE, RESPONSE.Length - 1)
                            End If
                            CreateResponse("", "B", RESPONSE)
                        Else
                            Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", SCANTEXT)
                            If rowWHTBARC1 Is Nothing Then
                                CreateResponse("", "R", "Invalid Case ID " & SCANTEXT)
                                Exit Select
                            End If
                            'ASCMAIN1.sql = "Select * from WHTLOCB1 where BAR_CODE = '" & SCANTEXT & "'" & vbCrLf _
                            '    & " and LOCATION_QTY > 0" & vbCrLf
                            'Dim rowLocRec As DataRow = ASCDATA1.GetDataRow
                            'If rowLocRec Is Nothing Then
                            '    CreateResponse("", "R", "Invalid Case ID " & SCANTEXT & " no available qty")
                            '    Exit Select
                            'End If
                            'If Val(rowLocRec.Item("LOCATION_QTY")) <> 0 Then

                            'End If


                            BAR_CODE = SCANTEXT
                            If tbl.Select("BAR_CODE_SCANNED = '" & BAR_CODE & "'").Length <> 0 Then
                                CreateResponse("REMOVE_LPN", "R", "Case " & BAR_CODE & " has already been Scanned.")
                                Exit Select
                            End If
                            If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE) Then
                                CreateResponse("", "R", "Could Not Lock Access to Case ID " & BAR_CODE)
                                Exit Select
                            End If
                            Dim row As DataRow = tbl.Rows.Find(BAR_CODE)
                            If row IsNot Nothing Then
                                ' THIS BAR_CODE IS ONE OF THE CASES THAT WAS ORIGINALLY WAVED
                                If row.Item("BAR_CODE_SCANNED") & "" <> "" Then ' IT WAS ALREADY SWAPPED OUT - TRY TO MOVE SWAP TO SOME OTHER
                                    Dim rows As DataRow() = tbl.Select("BAR_CODE_SCANNED IS NULL AND STYLE_COLOR_QTY_DNA = '" & row.Item("STYLE_COLOR_QTY_DNA") & "'")
                                    If rows.Length = 0 Then
                                        CreateResponse("", "R", "Cannot use Case ID " & BAR_CODE)
                                        Exit Select
                                    Else
                                        rows(0).Item("BAR_CODE_SCANNED") = row.Item("BAR_CODE_SCANNED")
                                        row.Item("BAR_CODE_SCANNED") = ""
                                    End If
                                End If
                                row.Item("BAR_CODE_SCANNED") = BAR_CODE
                                CreateResponse("", "G", "Case ID " & BAR_CODE & " Scanned")

                            Else
                                ' THIS BAR_CODE IS NOT ONE OF THE CASES THAT WAS ORIGINALLY WAVED
                                ASCMAIN1.sql = Replace(sqlDNA, "{sqlDNA_where}", "LOCATION_CODE = '" & LOCATION_CODE & "' and BAR_CODE = '" & BAR_CODE & "'")
                                Dim rowScanned As DataRow = ASCDATA1.GetDataRow
                                If rowScanned Is Nothing Then
                                    CreateResponse("", "R", "Cannot determine Contents of Case ID " & BAR_CODE & ", Check Location")
                                    Exit Select

                                Else
                                    Dim STYLE_COLOR_QTY_DNA As String = rowScanned.Item("STYLE_COLOR_QTY_DNA") & ""

                                    Dim rows As DataRow() = tbl.Select("BAR_CODE_SCANNED IS NULL AND STYLE_COLOR_QTY_DNA = '" & STYLE_COLOR_QTY_DNA & "'")
                                    If rows.Length = 0 Then
                                        CreateResponse("", "R", "Cannot use Case ID " & BAR_CODE)
                                        Exit Select
                                    Else
                                        ASCMAIN1.sql = "Select WHTINST1.* from WHTINST1,WHTINST2" & vbCrLf _
                                            & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                                            & "   and WHTINST1.WAVE_INST_STATUS = '0'" & vbCrLf _
                                            & "   and WHTINST2.BAR_CODE = '" & BAR_CODE & "'"
                                        Dim rowWHTINST1 As DataRow = ASCDATA1.GetDataRow
                                        If rowWHTINST1 IsNot Nothing Then
                                            Dim WAVE_INST_NO_OTHER As String = rowWHTINST1.Item("WAVE_INST_NO")
                                            Dim WAVE_NO_OTHER As String = rowWHTINST1.Item("WAVE_NO")
                                            Dim LOCATION_CODE_OTHER As String = rowWHTINST1.Item("LOCATION_CODE")
                                            If Not ASCMAIN1.Logical_Open("WHTWAVE1", WAVE_NO_OTHER) Then
                                                CreateResponse("", "R", "Could Not Access Wave " & WAVE_NO_OTHER)
                                                Exit Select
                                            End If
                                            If Not ASCMAIN1.Logical_Open("WHTINST1", WAVE_INST_NO_OTHER) Then
                                                CreateResponse("", "R", "Could Not Access Wave Instruction " & WAVE_INST_NO_OTHER)
                                                Exit Select
                                            End If
                                        End If
                                        rows(0).Item("BAR_CODE_SCANNED") = BAR_CODE
                                        CreateResponse("", "B", "Case ID+ " & BAR_CODE & " Scanned")
                                        Exit Select
                                    End If
                                End If
                            End If
                        End If
                    End If

                Case "REMOVE_LPN"
                    If SCANTEXT <> "Y" And SCANTEXT <> "N" Then
                        CreateResponse("", "R", "Invalid Response")
                        Exit Select
                    End If

                    If SCANTEXT = "Y" Then
                        Dim row As DataRow = tbl.Select("BAR_CODE_SCANNED = '" & BAR_CODE & "'")(0)
                        row.Item("BAR_CODE_SCANNED") = ""
                        RESPONSE = "Case ID " & BAR_CODE & " has been Removed"
                    Else
                        RESPONSE = "Case ID " & BAR_CODE & " duplicate scan Ignored"
                    End If
                    CreateResponse("SCAN_LPN", "B", RESPONSE & vbCrLf & "Resume Case Scans" & vbCrLf & WAVE_INST_TEXT, True)
                    Exit Select

                Case "SHORT_COUNT"
                    If SCANTEXT = "Y" Then
                        CreateResponse("VERIFY", "B", "")

                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_LPN", "B", "Resume Case Scans" & vbCrLf & WAVE_INST_TEXT, True)

                    End If


                Case "VERIFY"
                    'add missing carton verification 
                    If SCANTEXT = "Y" Then

                        Update_Record()

                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "Work on Inst " & WAVE_INST_NO & " Complete")

                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_LPN", "B", "Resume Case Scans" & vbCrLf & WAVE_INST_TEXT, True)

                    ElseIf SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "Work on Inst " & WAVE_INST_NO & " Cancelled")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Overrides Function GetStatus() As String
        Dim STATUS As String = "Scanned " & tbl.Select("ISNULL(BAR_CODE_SCANNED,'') <> ''").Length & " out of " & CStr(CASES) & " Cases"
        Return STATUS
    End Function

    Sub Update_Record()

        BeginTrans()

        Dim sqlw As String = "BAR_CODE <> BAR_CODE_SCANNED and ISNULL(BAR_CODE_SCANNED,'') <> ''"
        For Each row As DataRow In tbl.Select(sqlw)

            Dim BAR_CODE As String = row.Item("BAR_CODE")
            Dim BAR_CODE_SCANNED As String = row.Item("BAR_CODE_SCANNED")

            ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                & " where WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" & vbCrLf _
                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & "   and BAR_CODE in ('" & BAR_CODE & "','" & BAR_CODE_SCANNED & "')"
            Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)

            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("BAR_CODE = '" & BAR_CODE & "'")
                Dim STYLE_CODE As String = rowWHTLOCB1.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTLOCB1.Item("COLOR_CODE")
                Dim rowWHTLOCB1_SCANNED As DataRow = dst.Tables("WHTLOCB1").Rows.Find _
                                                     (New String() {GunEnv.WHSE_CODE, LOCATION_CODE, BAR_CODE_SCANNED, STYLE_CODE, COLOR_CODE})
                Dim LOCATION_QTY_WAVE As Int64 = Val(rowWHTLOCB1.Item("LOCATION_QTY_WAVE") & "")
                rowWHTLOCB1.Item("LOCATION_QTY_WAVE") = rowWHTLOCB1_SCANNED.Item("LOCATION_QTY_WAVE")
                rowWHTLOCB1_SCANNED.Item("LOCATION_QTY_WAVE") = LOCATION_QTY_WAVE
            Next
            Update_Record_TDA("WHTLOCB1")

            ASCMAIN1.sql = "Select WHTINST2.* from WHTINST2,WHTINST1 " & vbCrLf _
                & "    where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
                & "      and WHTINST2.BAR_CODE in ('" & BAR_CODE & "','" & BAR_CODE_SCANNED & "')" & vbCrLf _
                & "      and WHTINST1.WAVE_INST_STATUS ='0'" & vbCrLf
            Fill_Records("WHTINST2", "", True, ASCMAIN1.sql)

            ' CHANGE TO KEY FIELD BAR_CODE - CANNOT DO WITH ADO.NET
            For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select()
                If rowWHTINST2.Item("WAVE_INST_NO") = WAVE_INST_NO Then
                    ASCMAIN1.sql = "Update WHTINST2 Set BAR_CODE = '" & BAR_CODE_SCANNED & "'" & vbCrLf _
                        & " where WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
                        & "   and BAR_CODE = '" & BAR_CODE & "'"
                Else
                    ASCMAIN1.sql = "Update WHTINST2 Set BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                        & " where WAVE_INST_NO = '" & rowWHTINST2.Item("WAVE_INST_NO") & "'" & vbCrLf _
                        & "   and BAR_CODE = '" & BAR_CODE_SCANNED & "'"
                End If
                ASCDATA1.ExecuteSQL()
            Next
        Next

        Dim LOAD_NO_OTHER As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
        Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
        With rowWHTBARC0
            .Item("LOAD_NO") = LOAD_NO_OTHER
            .Item("WHSE_CODE") = GunEnv.WHSE_CODE
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = GunEnv.USER_ID
            .Item("LOAD_STATUS") = "A"
            .Item("LOCATION_CODE") = GunEnv.GUN_LOC
            .Item("TRAN_TYPE") = "V"
            .Item("TRAN_NO") = WAVE_INST_NO
            .Item("LOAD_DATE") = DATETIME_STAMP.Date
        End With
        dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
        Update_Record_TDA("WHTBARC0")


        ASCMAIN1.sql = "Select WHTINST2.* from WHTINST2,WHTINST1 " & vbCrLf _
            & "    where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
            & "      and WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
            & "      and WHTINST1.WAVE_INST_STATUS ='0'" & vbCrLf
        Fill_Records("WHTINST2", "", True, ASCMAIN1.sql)

        sqlw = "ISNULL(BAR_CODE_SCANNED,'') <> ''"
        For Each row As DataRow In tbl.Select(sqlw)
            Dim BAR_CODE_SCANNED As String = row.Item("BAR_CODE_SCANNED")
            For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("BAR_CODE = '" & BAR_CODE_SCANNED & "'")
                rowWHTINST2.Item("LOCATION_QTY_PICK") = rowWHTINST2.Item("LOCATION_QTY_WAVE")
            Next
            'ASCMAIN1.sql = "Update WHTINST2 Set LOCATION_QTY_PICK = LOCATION_QTY_WAVE" & vbCrLf _
            '    & " where WAVE_INST_NO = '" & WAVE_INST_NO & "' and BAR_CODE = '" & BAR_CODE_SCANNED & "'"
            'ASCDATA1.ExecuteSQL()
        Next
        Update_Record_TDA("WHTINST2")

        Dim partial_pick As Boolean = False
        For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("")
            Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE")
            Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE")
            Dim LOCATION_QTY_WAVE As String = Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
            Dim LOCATION_QTY_PICK As String = Val(rowWHTINST2.Item("LOCATION_QTY_PICK") & "")
            ASCMAIN1.sql = "Update WHTLOCB1 Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) - " & CStr(LOCATION_QTY_WAVE) & vbCrLf _
                & " where WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" & vbCrLf _
                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & "   and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            ASCDATA1.ExecuteSQL()

            If LOCATION_QTY_PICK = 0 Then
                partial_pick = True
            End If
        Next

        If partial_pick Then
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", GunEnv.WHSE_CODE)
            If LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_REC") & "" Then
                ' NEVER LOCK RECEIVING
            Else
                ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED = '1'" _
                    & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2" _
                    & "   and NVL(LOCATION_USE,'?') = '?'"
                Dim R As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {GunEnv.WHSE_CODE, LOCATION_CODE})

                If R = 1 Then
                    ASCMAIN1.Record_Event("WHTLOCM1", GunEnv.WHSE_CODE, LOCATION_CODE, Now, G.USER_ID, "LOCLCK", "CNL WVPICK", WAVE_INST_NO)
                End If
            End If
        End If

        ASCMAIN1.sql = "Update WHTINST1 Set" & vbCrLf _
            & "  WAVE_INST_STATUS = '1'" & vbCrLf _
            & ", LOCATION_CODE_OTHER = '" & GunEnv.GUN_LOC & "'" & vbCrLf _
            & ", LOAD_NO_OTHER = '" & LOAD_NO_OTHER & "'" & vbCrLf _
            & ", INIT_DATE = SYSDATE, INIT_OPER = '" & G.USER_ID & "'" & vbCrLf _
            & " where WAVE_INST_NO = '" & WAVE_INST_NO & "' and WAVE_INST_STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSP("WHPLOCB2", "VVV", _
                           New Object() {"V", WAVE_INST_NO, ASCMAIN1.SESSION_NO}, _
                           New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select WHTINST1.WAVE_NO, WHTINST2.WAVE_LNO, WHTINST1.WAVE_SUB, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE" & vbCrLf _
            & "   , Sum (WHTINST2.LOCATION_QTY_PICK) WAVE_QTY_PICK" & vbCrLf _
            & "   from WHTINST2,WHTINST1" & vbCrLf _
            & "   where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
            & "     and WHTINST2.WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
            & "   group by WHTINST1.WAVE_NO, WHTINST2.WAVE_LNO, WHTINST1.WAVE_SUB, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTWAVE2 Set WAVE_QTY_PICK = NVL(WAVE_QTY_PICK,0) + NVL(R1.WAVE_QTY_PICK,0)" & vbCrLf _
            & "    where WAVE_NO = R1.WAVE_NO" & vbCrLf _
            & "      and WAVE_LNO = R1.WAVE_LNO" & vbCrLf _
            & "      and ((R1.WAVE_SUB = '0' and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE)" & vbCrLf _
            & "       or  (R1.WAVE_SUB = '1' and STYLE_CODE_SUB = R1.STYLE_CODE and COLOR_CODE_SUB = R1.COLOR_CODE));" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into WHTWAVE2 (WAVE_NO, STYLE_CODE, COLOR_CODE, WAVE_QTY_PICK) " & vbCrLf _
            & "     values (R1.WAVE_NO, R1.STYLE_CODE, R1.COLOR_CODE, R1.WAVE_QTY_PICK);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL() ' I KNOW THAT THE INSERT WILL NOT WORK - WAVE_LNO CAME ABOUT AFTER IT WAS CODED - NEED TO DEVELOP IT AFTER WE HAVE AN EXAMPLE

        CommitTrans()
    End Sub

    Overrides Function Get_Anticipated_Next_Response() As String
        Select Case AppState
            Case "NEXT_INST"
                Return ""
            Case "SCAN_LPN"
                Dim rows() As DataRow = tbl.Select("ISNULL(BAR_CODE_SCANNED,'') = ''")
                If rows.Length = 0 Then
                    Return "DONE"
                Else
                    Return rows(0).Item("BAR_CODE")
                End If
            Case "REMOVE_LPN"
                Return ""
            Case "SHORT_COUNT"
                Return "Y"
            Case "VERIFY"
                Return "Y"
            Case Else
                Return ""
        End Select
    End Function
End Class