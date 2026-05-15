Public Class WHCRF002
    ' Move Carton (from Carton's Location) to Gun

    Inherits WHCRF000

    Dim BAR_CODE As String

    Sub New(g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF002"

        AppStates.Add("SCAN_LPN", "Scan Case ID to Move to Gun|DONE|EXIT|")
        AppStates.Add("REMOVE_LPN", "Remove Case? (Y/N)|Y|N|")
        AppStates.Add("VERIFY", "Are You Done (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_LPN"

        With dst
            With .Tables.Add("WHTSCANS").Columns
                .Add("BAR_CODE_SCANNED")
            End With
            .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE_SCANNED")}

            Create_TDA(.Tables.Add, "WHTBARC0", "*")
            Create_TDA(.Tables.Add, "WHTBARC1", "*")
            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
        End With

        tbl = dst.Tables("WHTSCANS") ' New DataTable

    End Sub

    Public Overrides Function Hello() As String
        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        Else
            Select Case AppState
                Case "SCAN_LPN"
                    If SCANTEXT = "DONE" Then
                        CreateResponse("VERIFY", "B", "Scanned " & tbl.Select("").Length & " Cases")
                    Else
                        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", SCANTEXT)
                        If rowWHTBARC1 Is Nothing Then
                            CreateResponse("", "R", "Invalid Case ID " & SCANTEXT)
                            Exit Select
                        End If
                        BAR_CODE = SCANTEXT
                        'Dim row As DataRow = tbl.Rows.Find(BAR_CODE)
                        'If row Is Nothing Then

                        'End If

                        If tbl.Select("BAR_CODE_SCANNED = '" & BAR_CODE & "'").Length <> 0 Then
                            CreateResponse("REMOVE_LPN", "R", "Case " & BAR_CODE & " has already been Scanned.")
                            Exit Select
                        End If

                        ASCMAIN1.sql = "Select Distinct LOCATION_CODE from WHTLOCB1" & vbCrLf _
                            & " where BAR_CODE = '" & BAR_CODE & "'" _
                            & " and WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" _
                            & " and LOCATION_QTY > 0"
                        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                        If rows.Length = 0 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " not found in Warehouse with Qty")
                            Exit Select
                        ElseIf rows.Length > 1 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " found in Multiple Locations with Qty - Call ABS")
                            Exit Select
                        ElseIf rows(0).Item("LOCATION_CODE") = GunEnv.GUN_LOC Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " is already loaded in this Gun Location")
                            Exit Select
                        End If

                        ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                          & " where BAR_CODE = '" & BAR_CODE & "'" _
                          & " and WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" _
                          & " and LOCATION_QTY_WAVE > 0"
                        rows = ASCDATA1.GetDataTable.Select("")
                        If rows.Length <> 0 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " is already Committed to a Wave")
                            Exit Select
                        End If


                        If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE) Then
                            CreateResponse("", "R", "Could not lock access to Case Id " & BAR_CODE)
                            Exit Select
                        End If

                        Dim row2 As DataRow = tbl.NewRow
                        row2.Item("BAR_CODE_SCANNED") = BAR_CODE
                        tbl.Rows.Add(row2)

                        ' tbl.Rows.Add(New Object() {BAR_CODE})

                        CreateResponse("", "B", "Case ID+ " & BAR_CODE & " Scanned")
                        Exit Select
                    End If

                Case "REMOVE_LPN"
                    If SCANTEXT <> "Y" And SCANTEXT <> "N" Then
                        CreateResponse("", "R", "Invalid Response")
                        Exit Select
                    End If

                    If SCANTEXT = "Y" Then
                        Dim row As DataRow = tbl.Rows.Find(BAR_CODE) ' tbl.Select("BAR_CODE_SCANNED = '" & BAR_CODE & "'")(0)
                        row.Delete()
                        RESPONSE = "Case ID " & BAR_CODE & " has been Removed"
                    Else
                        RESPONSE = "Case ID " & BAR_CODE & " duplicate scan Ignored"
                    End If
                    CreateResponse("SCAN_LPN", "B", RESPONSE & vbCrLf & "Resume Case Scans")
                    Exit Select

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        If tbl.Rows.Count = 0 Then
                            CreateResponse("", "R", "No Cases")
                            Exit Select
                        End If

                        Update_Record()
                        ASCMAIN1.MultiTask_Release()

                        CreateResponse("SCAN_LPN", "B", CStr(tbl.Select("").Count) & " Cartons Successfully moved to Gun")
                        tbl.Rows.Clear()

                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_LPN", "B", "Resume Case Scans")

                    ElseIf SCANTEXT = "CANCEL" Then

                        tbl.Rows.Clear()
                        CreateResponse("SCAN_LPN", "B", CStr(tbl.Rows.Count) + " Cartons scanned since last update were cancelled!")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables("WHTBARC1").Rows.Clear()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        Dim LOAD_NO As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = "M"
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("WHSE_CODE") = GunEnv.WHSE_CODE
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = G.USER_ID
            .Item("STATUS") = "U"
        End With
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
        Update_Record_TDA("WHTMOVE1")

        'Dim LOAD_NO_OTHER As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
        Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
        With rowWHTBARC0
            .Item("LOAD_NO") = LOAD_NO
            .Item("WHSE_CODE") = GunEnv.WHSE_CODE
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = GunEnv.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = GunEnv.USER_ID
            .Item("LOAD_STATUS") = "A"
            .Item("LOAD_COMMENT") = "Move Cartons to Gun"
            .Item("LOCATION_CODE") = GunEnv.GUN_LOC
            .Item("TRAN_TYPE") = "M"
            .Item("TRAN_NO") = WHSE_TRAN_NO
            .Item("LOAD_DATE") = DATETIME_STAMP.Date
        End With
        dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
        Update_Record_TDA("WHTBARC0")

        Dim WHSE_TRAN_LNO_ctr As Integer = 0

        For Each row As DataRow In tbl.Select("")

            Dim BAR_CODE As String = row.Item("BAR_CODE_SCANNED")
            Dim rowWHTBARC1 As DataRow = Fill_Record("WHTBARC1", BAR_CODE, , False)

            ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                & " where BAR_CODE = '" & BAR_CODE & "'" _
                & " and WHSE_CODE = '" & GunEnv.WHSE_CODE & "'" _
                & " and LOCATION_QTY > 0"
            Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)

            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("", "BAR_CODE, STYLE_CODE, COLOR_CODE")
                Dim STYLE_CODE As String = rowWHTLOCB1.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTLOCB1.Item("COLOR_CODE")
                Dim LOCATION_QTY As Integer = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "")
                Dim LOCATION_CODE As String = rowWHTLOCB1.Item("LOCATION_CODE")

                Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                With rowWHTMOVE2
                    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                    WHSE_TRAN_LNO_ctr += 1
                    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                    .Item("LOCATION_CODE_FROM") = LOCATION_CODE
                    .Item("LOCATION_CODE_TO") = GunEnv.GUN_LOC
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("WHSE_TRAN_QTY") = LOCATION_QTY
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("INIT_OPER") = G.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = G.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "U"
                    .Item("LOAD_NO_FROM") = rowWHTBARC1.Item("LOAD_NO")
                    .Item("LOAD_NO_TO") = LOAD_NO
                End With
                dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

            Next

            rowWHTBARC1.Item("LOAD_NO") = LOAD_NO
            'rowWHTBARC1.Item("LOCATION_CODE") = G.GUN_LOC
            rowWHTBARC1.Item("LAST_OPER") = G.USER_ID
            rowWHTBARC1.Item("LAST_DATE") = DATETIME_STAMP
        Next

        Update_Record_TDA("WHTBARC1")
        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", _
                       New Object() {WHSE_TRAN_NO, 0, 1}, _
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        CommitTrans()
    End Sub
End Class