Public Class APRCHKP1
    Dim BANK_LAST_CHECK_NO As Int64
    Dim BANK_LAST_CHECK_NO_orig As Int64
    Dim rowGLTBANK1 As DataRow
    Dim rowAPTPYMT1 As DataRow
    Dim APTINVH1 As String = ""
    Dim eCheck As CC.eCheck
    Dim FOLDERNAME_CONTROL As String = ASCMAIN1.Folders("Archive") & "\eChecks\Control\"
    Dim FILENAME_CONTROL As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        If Not System.IO.Directory.Exists(FOLDERNAME_CONTROL) Then System.IO.Directory.CreateDirectory(FOLDERNAME_CONTROL)
    End Sub

    Overrides Sub Clear_Record()

        Load_Drop_Down("BATCH_NO_PYMT")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim BATCH_NO_PYMT As String = Absx1.cmbFor("BATCH_NO_PYMT").Text

        Prepare_dst(True, BATCH_NO_PYMT)

        Check_if_Empty("APTPYMT1")
    End Sub

    Public Overrides Sub Print_Report()

        Dim BATCH_NO_PYMT As String = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")
        Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")

        'If rowGLTBANK1.Item("CHECK_REPORT") & "" <> "" Then
        '    RPT = rowGLTBANK1.Item("CHECK_REPORT")
        'End If

        'If PYMT_METHOD = "DBAUTH" Or PYMT_METHOD = "ECHECK" Or PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
        '    RPT = "APRCHKPE"
        'End If

        'CR_params.Add("COPY", "")
        'Generate_Report(RPT)

        ' If rowGLTBANK1.Item("BANK_PYMT_METHOD") & "" = "DBAUTH" Then
        If PYMT_METHOD = "DBAUTH" Or PYMT_METHOD = "ECHECK" Or PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
            RPT = "APRCHKPE"
            Send_DBAUTHs()
        End If




        If rowGLTBANK1.Item("CHECK_REPORT") & "" <> "" Then
            RPT = rowGLTBANK1.Item("CHECK_REPORT")
        End If

        If PYMT_METHOD = "DBAUTH" Or PYMT_METHOD = "ECHECK" Or PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
            RPT = "APRCHKPE"
        End If

        CR_params.Add("COPY", "")
        Generate_Report(RPT)



        CR_params.Add("PAYMENT_SELECTION", "0")
        Generate_Report("APRPYMT1", "Printed Checks Report", "")

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)
        rowGLTBANK1.Item("BATCH_NO_PYMT") = BATCH_NO_PYMT
        Update_Record_TDA("GLTBANK1")

        Write_Event_Log_Batch("APTINVH1", "Select APTINVH1.VOUCHER_NO, 'Check ' || TRIM(TO_CHAR(TO_NUMBER(SUBSTR(APTINVH1.CHECK_NUM,2)) + " & CStr(BANK_LAST_CHECK_NO_orig) & ",'0000000000')) || ' Printed' from APTINVH1 where APTINVH1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        Write_Event_Log("GLTBANK1", BANK_CODE, "Checks Printed (" & CStr(dst.Tables("APTPYMT2").Rows.Count) & ")")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("BATCH_NO_PYMT").Text = "" Then
                EMsg &= vbCr & "You must pick a Batch"
            Else
                Dim BATCH_NO_PYMT As String = Absx1.cmbFor("BATCH_NO_PYMT").Text
                rowAPTPYMT1 = LookUp("APTPYMT1", BATCH_NO_PYMT)
                Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")

                rowGLTBANK1 = LookUp("GLTBANK1", BANK_CODE)
                If rowGLTBANK1.Item("BATCH_NO_PYMT") & "" <> "" _
                And rowGLTBANK1.Item("BATCH_NO_PYMT") & "" <> BATCH_NO_PYMT Then
                    EMsg &= vbCr & "Bank " & BANK_CODE & " is in Process with Batch " & rowGLTBANK1.Item("BATCH_NO_PYMT")
                End If

                ASCMAIN1.sql = "Select VEND_CODE, VEND_NAME from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & Absx1.cmbFor("BATCH_NO_PYMT").Text & "' union Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & Absx1.cmbFor("BATCH_NO_PYMT").Text & "') and VEND_ON_HOLD = '1'"
                Dim tblAPTVEND1_hold As DataTable = ASCDATA1.GetDataTable
                If tblAPTVEND1_hold.Rows.Count <> 0 Then
                    For Each row As DataRow In tblAPTVEND1_hold.Rows
                        EMsg &= vbCr & "Vendor on Pymt Hold: " & row.Item("VEND_CODE") & ":" & row.Item("VEND_NAME")
                    Next
                End If

                Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")
                Dim SSH_APP_CODE As String = rowGLTBANK1.Item("SSH_APP_CODE") & ""

                If rowGLTBANK1.Item("BANK_PP_IND") & "" = "1" Then
                    If SSH_APP_CODE = "" Then
                        Throw New Exception("No SSH record to use for PP Transmit")
                    End If
                End If

                If PYMT_METHOD = "ECHECK" Then
                    ' NEED TO PROVIDE A WARNING TO THE USER THAT THE ECHECK IS GOING TO BE POSTED TO THE BANK UPON UPDATE
                    FILENAME_CONTROL = BANK_CODE & ".txt"
                    If System.IO.File.Exists(FOLDERNAME_CONTROL & FILENAME_CONTROL) Then
                        EMsg &= vbCr & $"Incomplete ACH Payment Control File for Bank {BANK_CODE} in Control Folder: {FOLDERNAME_CONTROL}"
                    End If
                    If MsgBox("Please Note - upon Update these Payments will be immediately submitted for eCheck Payment" & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        EMsg &= vbCr & "Proceed Action Cancelled"
                    End If
                End If

                If PYMT_METHOD = "ACH" Then
                    If SSH_APP_CODE = "" Then
                        Throw New Exception("No SSH record to use for ACH Transmit")
                    End If
                End If
            End If

        ElseIf eItemKey = "Update" Then
            'Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")
            'If PYMT_METHOD = "ECHECK" Then


            'End If
        End If


    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst
            Create_TDA(.Tables.Add, "APTCHCK1", "*")
            Create_TDA(.Tables.Add, "APTCHCK2", "*")
            Create_TDA(.Tables.Add, "APTCHCK4", "*")

            Create_TDA(.Tables.Add, "APTCHCK5", "*")
            .Tables("APTCHCK5").Columns.Add("VEND_BANK_ACCT_ID_DECRYPTED")

            Create_TDA(.Tables.Add, "APTVEND5", "*")
            Create_TDA(.Tables.Add, "APTPYMT1", "*", 1)
            Create_TDA(.Tables.Add, "APTPYMT2", "*", 1)

            ASCMAIN1.sql = "Select * from APTINVH1 where ROWNUM < 1"
            APTINVH1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APTINVH1 & " Add Primary Key (VOUCHER_NO)")

            ASCMAIN1.sql = "Select * from " & APTINVH1
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, Update_COLUMN_NAMEs:="CHECK_NUM,CHECK_DATE,INV_STATUS,INV_PAYMENTS,INV_DISC_TAKEN,INV_LAST_PMT_DATE,BANK_CODE,INV_BALANCE,BATCH_PYMT,BATCH_DISC")

            ASCMAIN1.sql = "Select * from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from " & APTINVH1 & " union Select Distinct VEND_CODE_AP from " & APTINVH1 & ")"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, 1)

            ASCMAIN1.sql = "Select * from APTVEND2 where VEND_CODE in (Select Distinct VEND_CODE from " & APTINVH1 & " union Select Distinct VEND_CODE_AP from " & APTINVH1 & ")"
            Create_TDA(.Tables.Add, "APTVEND2", "**", 0, False, 2)

            Create_TDA(.Tables.Add, "GLTBANK1", "*")
            With .Tables("GLTBANK1")
                .Columns.Add("BANK_SIGNATURE", GetType(System.Byte()))
            End With

        End With

        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        'sqlw = parms(0)
        Dim BATCH_NO_PYMT As String = parms(0)
        EnforceConstraints(False)

        If BATCH_NO_PYMT <> "" Then
            rowAPTPYMT1 = Fill_Record("APTPYMT1", BATCH_NO_PYMT)
        Else
            ASCMAIN1.sql = "Select NVL(BATCH_NO_PYMT,'000000') BATCH_NO_PYMT, CHECK_DATE, BANK_CODE from APTCHCK1 where BANK_CODE = '" & parms(1) & "' and CHECK_NUM = '" & parms(2) & "'"
            Fill_Records("APTPYMT1", "", True, ASCMAIN1.sql)
            rowAPTPYMT1 = dst.Tables("APTPYMT1").Rows(0)
        End If

        Dim CHECK_DATE As Date = rowAPTPYMT1.Item("CHECK_DATE")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")

        If BATCH_NO_PYMT <> "" Then
            Fill_Records("APTPYMT2", BATCH_NO_PYMT)
        Else
            ASCMAIN1.sql = "Select NVL(BATCH_NO_PYMT,'000000') BATCH_NO_PYMT, CHECK_NUM, VEND_CODE_AP, NULL VOUCHER_NO, CHECK_AMT BATCH_PYMT, 0 BATCH_DISC, VEND_ALT_CODE, VEND_CODE, VEND_NAME from APTCHCK1 where BANK_CODE = '" & parms(1) & "' and CHECK_NUM = '" & parms(2) & "'"
            Fill_Records("APTPYMT2", "", True, ASCMAIN1.sql)
        End If

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)

        Dim BANK_SIGNATURE_IMAGE_FILENAME As String = ASCMAIN1.Folders("Archive") & "\ABS\APRCHKP1\" & BANK_CODE ' & ".PNG"
        If My.Computer.FileSystem.FileExists(BANK_SIGNATURE_IMAGE_FILENAME) Then
            rowGLTBANK1.Item("BANK_SIGNATURE") = ASCMAIN1.GetImageData(BANK_SIGNATURE_IMAGE_FILENAME)
        Else
            BANK_SIGNATURE_IMAGE_FILENAME = ASCMAIN1.Folders("Archive") & "\ABS\BANK_SIGNATURE\" & BANK_CODE & ".JPG"
            If My.Computer.FileSystem.FileExists(BANK_SIGNATURE_IMAGE_FILENAME) Then
                rowGLTBANK1.Item("BANK_SIGNATURE") = ASCMAIN1.GetImageData(BANK_SIGNATURE_IMAGE_FILENAME)
            End If
        End If

        ASCDATA1.ExecuteSQL("Delete from " & APTINVH1)

        If BATCH_NO_PYMT <> "" Then
            ASCDATA1.ExecuteSQL("Insert into " & APTINVH1 & " Select APTINVH1.* from APTINVH1 where APTINVH1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        Else
            Dim sql2 As String = "APTCHCK2.BANK_CODE = '" & parms(1) & "' and APTCHCK2.CHECK_NUM = '" & parms(2) & "'"
            ASCDATA1.ExecuteSQL("Insert into " & APTINVH1 & " Select APTINVH1.* from APTINVH1,APTCHCK2 where APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO and " & sql2)
            ASCMAIN1.sql = "Begin Declare Cursor C1 is Select * from APTCHCK2 where " & sql2 & "; Begin for R1 in C1 Loop Update " & APTINVH1 & " Set BATCH_NO_PYMT = NVL(BATCH_NO_PYMT,'000000'), BATCH_PYMT = NVL(R1.INV_AMT_APPLIED,0) - NVL(R1.INV_DISC_TAKEN,0), BATCH_DISC = R1.INV_DISC_TAKEN where VOUCHER_NO = R1.VOUCHER_NO; End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If
        Fill_Records("APTINVH1")

        Fill_Records("APTVEND1")
        Fill_Records("APTVEND2")

        If BATCH_NO_PYMT <> "" Then
            Generate_Checks(BANK_CODE, CHECK_DATE)
        End If

        EnforceConstraints(True)
    End Sub

    Sub Generate_Checks(BANK_CODE As String, CHECK_DATE As Date)

        dst.Tables("APTCHCK5").Rows.Clear()

        sql = "Select FIELD_LENGTH from ASTFFMT1 where COLUMN_NAME = 'CHECK_NUM'"
        Dim CHECK_NUM_length As Integer = Val(ASCDATA1.GetDataValue(sql))
        If CHECK_NUM_length = 0 Then
            CHECK_NUM_length = 10
        End If

        BANK_LAST_CHECK_NO = Val(rowGLTBANK1.Item("BANK_LAST_CHECK_NO") & "")
        BANK_LAST_CHECK_NO_orig = BANK_LAST_CHECK_NO
        Dim CHECK_NUM As String = ""

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("", "CHECK_NUM")
            Do
                BANK_LAST_CHECK_NO = BANK_LAST_CHECK_NO + 1
                CHECK_NUM = Format(BANK_LAST_CHECK_NO, "".PadLeft(CHECK_NUM_length, "0"))
                LookUp("APTCHCK1", New String() {BANK_CODE, CHECK_NUM})
            Loop While cdr IsNot Nothing
            Dim CHECK_NUM_X As String = rowAPTPYMT2.Item("CHECK_NUM")
            rowAPTPYMT2.Item("CHECK_NUM") = CHECK_NUM

            For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("CHECK_NUM = '" & CHECK_NUM_X & "'", "")
                rowAPTINVH1.Item("CHECK_NUM") = CHECK_NUM
            Next
        Next
    End Sub

    Overrides Sub Update_Record()

        Dim SEQ_NUM As Integer = 0
        Dim CHECK_NUM As String = ""

        Dim CHECK_DATE As Date = rowAPTPYMT1.Item("CHECK_DATE")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")
        Dim BATCH_NO_PYMT As String = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")

        Dim SSH_APP_CODE As String = rowGLTBANK1.Item("SSH_APP_CODE") & ""
        Dim BANK_PYMT_METHOD As String = "CHECK"
        If rowGLTBANK1.Item("BANK_PYMT_METHOD") & "" <> "" Then
            BANK_PYMT_METHOD = rowGLTBANK1.Item("BANK_PYMT_METHOD")
        End If

        ASCMAIN1.Progress("Now Processing AP Items")

        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("", "CHECK_NUM, VOUCHER_NO")

            Dim INV_AMT_APPLIED As Decimal = Val(rowAPTINVH1.Item("BATCH_PYMT") & "") + Val(rowAPTINVH1.Item("BATCH_DISC") & "")
            Dim INV_DISC_TAKEN As Decimal = Val(rowAPTINVH1.Item("BATCH_DISC") & "")

            With rowAPTINVH1
                .Item("CHECK_DATE") = CHECK_DATE
                .Item("INV_STATUS") = "P"
                .Item("INV_PAYMENTS") = Val(.Item("INV_PAYMENTS") & "") + Val(.Item("BATCH_PYMT") & "")
                .Item("INV_DISC_TAKEN") = Val(.Item("INV_DISC_TAKEN") & "") + Val(.Item("BATCH_DISC") & "")
                .Item("INV_LAST_PMT_DATE") = CHECK_DATE
                .Item("BANK_CODE") = BANK_CODE
                .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - INV_AMT_APPLIED
                .Item("BATCH_PYMT") = 0
                .Item("BATCH_DISC") = 0

                If CHECK_NUM <> .Item("CHECK_NUM") Then
                    SEQ_NUM = 0
                    CHECK_NUM = .Item("CHECK_NUM")
                End If
            End With


            Dim rowAPTCHCK2 As DataRow = dst.Tables("APTCHCK2").NewRow
            With rowAPTCHCK2
                .Item("BANK_CODE") = BANK_CODE
                .Item("CHECK_NUM") = CHECK_NUM
                SEQ_NUM = SEQ_NUM + 1
                .Item("SEQ_NUM") = SEQ_NUM
                .Item("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")
                .Item("INV_NUM") = rowAPTINVH1.Item("INV_NUM")
                .Item("INV_DATE") = rowAPTINVH1.Item("INV_DATE")
                .Item("VOUCHER_NO") = rowAPTINVH1.Item("VOUCHER_NO")
                .Item("INV_AMT_APPLIED") = INV_AMT_APPLIED
                .Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
            End With
            dst.Tables("APTCHCK2").Rows.Add(rowAPTCHCK2)
        Next
        Update_Record_TDA("APTCHCK2")
        Update_Record_TDA("APTINVH1")

        ASCMAIN1.Progress("Now Processing Checks")

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Rows
            Call ASCMAIN1.Progress("", rowAPTPYMT2.Item("CHECK_NUM"))

            Dim rowAPTVEND5 = Fill_Record("APTVEND5", rowAPTPYMT2.Item("VEND_CODE"), True, False)
            With rowAPTVEND5
                .ITEM("VEND_PAYMENTS_MTD") = Val(.ITEM("VEND_PAYMENTS_MTD") & "") + Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")
                .ITEM("VEND_PAYMENTS_YTD") = Val(.ITEM("VEND_PAYMENTS_YTD") & "") + Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")
                .ITEM("VEND_DISC_TAKEN_MTD") = Val(.ITEM("VEND_DISC_TAKEN_MTD") & "") + Val(rowAPTPYMT2.Item("BATCH_DISC") & "")
                .ITEM("VEND_DISC_TAKEN_YTD") = Val(.ITEM("VEND_DISC_TAKEN_YTD") & "") + Val(rowAPTPYMT2.Item("BATCH_DISC") & "")
                .ITEM("VEND_NUM_CHKS_MTD") = Val(.ITEM("VEND_NUM_CHKS_MTD") & "") + 1
                .ITEM("VEND_NUM_CHKS_YTD") = Val(.ITEM("VEND_NUM_CHKS_YTD") & "") + 1
                .ITEM("VEND_LAST_PMT_DATE") = CHECK_DATE
                .ITEM("VEND_LAST_PMT_AMT") = Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")
            End With

            Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").NewRow
            With rowAPTCHCK1
                .Item("BANK_CODE") = BANK_CODE
                .Item("CHECK_NUM") = rowAPTPYMT2.Item("CHECK_NUM")
                .Item("CHECK_DATE") = CHECK_DATE
                .Item("CHECK_AMT") = rowAPTPYMT2.Item("BATCH_PYMT")
                .Item("PYMT_METHOD") = PYMT_METHOD
                .Item("VEND_CODE") = rowAPTPYMT2.Item("VEND_CODE")
                .Item("VEND_CODE_AP") = rowAPTPYMT2.Item("VEND_CODE_AP")
                .Item("VEND_ALT_CODE") = rowAPTPYMT2.Item("VEND_ALT_CODE")
                .Item("VEND_NAME") = rowAPTPYMT2.Item("VEND_NAME")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("CHECK_STATUS") = "I"
                .Item("REGISTER_IND") = "0"
                .Item("BATCH_NO_PYMT") = rowAPTPYMT2.Item("BATCH_NO_PYMT")
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                If PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
                    If ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" = "1" Then
                        .Item("ACH_PAY_STATUS_IND") = "P" ' Pending Transmission
                    End If
                ElseIf PYMT_METHOD = "ECHECK" Then
                    If ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" = "1" Then
                        .Item("ECK_PAY_STATUS_IND") = "P" ' Pending Transmission
                    End If
                ElseIf PYMT_METHOD = "CHECK" AndAlso rowGLTBANK1.Item("BANK_PP_IND") & "" = "1" Then
                    .Item("POS_PAY_STATUS_IND") = "P" ' Pending Transmission
                End If
            End With

            dst.Tables("APTCHCK1").Rows.Add(rowAPTCHCK1)
        Next

        Update_Record_TDA("APTCHCK1")
        Update_Record_TDA("APTCHCK5")
        Update_Record_TDA("APTVEND5")

        Write_Event_Log_Batch("APTINVH1", "Select APTCHCK2.VOUCHER_NO, 'Check ' || APTCHCK2.CHECK_NUM || ' Updated' from APTCHCK2,APTCHCK1 where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM and APTCHCK1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")

        ASCDATA1.ExecuteSQL("Delete from APTPYMT1 where BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        ASCDATA1.ExecuteSQL("Delete from APTPYMT2 where BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)
        If Val(rowGLTBANK1.Item("BANK_LAST_CHECK_NO") & "") <> BANK_LAST_CHECK_NO_orig Then
            Throw New Exception("Error updating GL Bank Table Last Check - please call ABS")
            Stop ' NEED TO ROLLBACK
        End If
        rowGLTBANK1.Item("BANK_LAST_CHECK_NO") = BANK_LAST_CHECK_NO
        rowGLTBANK1.Item("BATCH_NO_PYMT") = ""
        Update_Record_TDA("GLTBANK1")

        Select Case PYMT_METHOD
            Case "ECHECK"
                Pay_using_ECHECK
            Case "ACH"
                ' Pay_using_ACH
            Case "WIRE"
                ' Pay_using_WIRE
        End Select

    End Sub

    Sub Pay_using_WIRE()
        Stop
    End Sub

    Sub Pay_using_ACH()
        Stop
    End Sub

    Sub Pay_using_ECHECK()

        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")

        Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("APTCHCK4.TRAN_NO")
        Dim FOLDERNAME As String = ASCMAIN1.Folders("Archive") & "\eChecks\Responses\" & TRAN_NO & "\"
        System.IO.Directory.CreateDirectory(FOLDERNAME)

        eCheck = New CC.eCheck()

        eCheck.TRAN_NO = TRAN_NO
        eCheck.FOLDERNAME = FOLDERNAME
        eCheck.INIT_OPER = ASCMAIN1.USER_ID
        eCheck.INIT_DATE = DATETIME_STAMP

        With eCheck
            .rowGLTBANK1 = rowGLTBANK1
            .tblAPTCHCK1 = dst.Tables("APTCHCK1")
            .tblAPTCHCK4 = dst.Tables("APTCHCK4")
            .tblAPTVEND1 = dst.Tables("APTVEND1")
            .tblAPTVEND2 = dst.Tables("APTVEND2")
        End With

        Dim errorList As List(Of String) = eCheck.ValidatePayees()

        If errorList.Count > 0 Then
            For Each errorMsg As String In errorList
                If errorMsg.Length = 0 Then Continue For
                EMsg &= vbCr & errorMsg
            Next
        End If

        ' NEED TO SET A BLOCK AND THEN SET A RELEASE ON USING BANK
        Using sw As New System.IO.StreamWriter(FOLDERNAME_CONTROL & FILENAME_CONTROL)
            sw.WriteLine(TRAN_NO & " Transaction No")
        End Using

        Dim errorOccured As Boolean = False

        For Each rowAPTCHCK1 As DataRow In dst.Tables("APTCHCK1").Select("", "VEND_CODE") ' Select("CHECK_AMT <> 0", "VEND_CODE")

            Dim CHECK_NUM As String = rowAPTCHCK1.Item("CHECK_NUM")
            Dim CHECK_AMT As Decimal = Val(rowAPTCHCK1.Item("CHECK_AMT") & "")

            If CHECK_AMT <= 0 Then
                rowAPTCHCK1.Item("ECK_PAY_STATUS_IND") = "S"
            Else

                Using sw As New System.IO.StreamWriter(FOLDERNAME_CONTROL & FILENAME_CONTROL, True)
                    sw.WriteLine(CHECK_NUM & " Sending")
                End Using

                If eCheck.Send_eChecks(rowAPTCHCK1, CC.eCheck.eCheckTypes.Credit) Then

                    rowAPTCHCK1.Item("ECK_PAY_STATUS_IND") = "S"

                    Using sw As New System.IO.StreamWriter(FOLDERNAME_CONTROL & FILENAME_CONTROL, True)
                        sw.WriteLine(CHECK_NUM & " Sent")
                    End Using

                    Stop ' The process worked

                Else
                    Stop ' The process failed
                    errorOccured = True
                    Using sw As New System.IO.StreamWriter(FOLDERNAME_CONTROL & FILENAME_CONTROL, True)
                        sw.WriteLine(eCheck.CheckResponse.ERRORMESSAGE)
                    End Using
                End If
            End If

        Next

        If Not errorOccured Then
            System.IO.File.Move(FOLDERNAME_CONTROL & BANK_CODE & ".txt", FOLDERNAME & FILENAME_CONTROL)
        End If

        Update_Record_TDA("APTCHCK1")
        Update_Record_TDA("APTCHCK4")


        ' IS THERE PROTECTION AGAINST DUP
    End Sub

    Sub Send_DBAUTHs()

        Dim BATCH_NO_PYMT As String = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")
        Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")

        Dim CHECK_DATE As Date = rowAPTPYMT1.Item("CHECK_DATE")

        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)
        Dim USER_SIGNATURE As String =
          rowASTUSER1.Item("USER_NAME") & vbCrLf _
        & rowASTUSER1.Item("USER_TITLE") & vbCrLf _
        & rowASTUSER1.Item("USER_COMPANY") & vbCrLf _
        & "Tel: " & rowASTUSER1.Item("USER_TELEPHONE") & vbCrLf _
        & "Fax: " & rowASTUSER1.Item("USER_FAX") & vbCrLf _
        & rowASTUSER1.Item("USER_EMAIL") & vbCrLf

        Dim VEND_CODE As String
        Dim VEND_NAME As String
        Dim VEND_ACCT_NO As String = ""
        Dim VEND_CODE_AP As String
        Dim VEND_ALT_CODE As String

        Dim VEND_EMAIL As String

        Dim VEND_CONTACT As String
        Dim VEND_PHONE As String

        Dim VEND_BANK_ACCT_ID As String
        Dim VEND_BANK_ROUTING_NO As String
        Dim VEND_BANK_SWIFT_NO As String
        Dim VEND_BANK_ACCT_CLASS As String
        Dim VEND_BANK_ACCT_TYPE As String

        Dim VEND_BANK_NAME As String
        Dim VEND_BANK_ADDR1 As String
        Dim VEND_BANK_ADDR2 As String
        Dim VEND_BANK_ADDR3 As String

        Dim VEND_BANK_CITY As String
        Dim VEND_BANK_STATE As String
        Dim VEND_BANK_ZIP_CODE As String

        Dim VEND_BANK_COUNTRY As String
        Dim VEND_BANK_CONTACT As String
        Dim VEND_REMIT_EMAIL As String

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Rows
            ASCMAIN1.Progress("Now Generating emails")

            VEND_CODE = rowAPTPYMT2.Item("VEND_CODE")
            VEND_NAME = rowAPTPYMT2.Item("VEND_NAME")

            Dim CHECK_NUM As String = rowAPTPYMT2.Item("CHECK_NUM") & ""

            Dim rowAPTCHCK5 As DataRow = dst.Tables("APTCHCK5").NewRow
            rowAPTCHCK5.Item("BANK_CODE") = BANK_CODE
            rowAPTCHCK5.Item("CHECK_NUM") = CHECK_NUM
            dst.Tables("APTCHCK5").Rows.Add(rowAPTCHCK5)

            VEND_CODE_AP = rowAPTPYMT2.Item("VEND_CODE_AP")
            VEND_ALT_CODE = rowAPTPYMT2.Item("VEND_ALT_CODE") & ""

            If VEND_CODE <> VEND_CODE_AP Then
                LookUp("APTVEND1", VEND_CODE_AP)
                VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                VEND_ACCT_NO = cdr.Item("VEND_ACCT_NO") & ""
                VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                VEND_PHONE = cdr.Item("VEND_PHONE") & ""

                VEND_BANK_ACCT_ID = cdr.Item("VEND_BANK_ACCT_ID") & ""
                VEND_BANK_ROUTING_NO = cdr.Item("VEND_BANK_ROUTING_NO") & ""
                VEND_BANK_SWIFT_NO = cdr.Item("VEND_BANK_SWIFT_NO") & ""
                VEND_BANK_ACCT_CLASS = cdr.Item("VEND_BANK_ACCT_CLASS") & ""
                VEND_BANK_ACCT_TYPE = cdr.Item("VEND_BANK_ACCT_TYPE") & ""

                VEND_BANK_NAME = cdr.Item("VEND_BANK_NAME") & ""
                VEND_BANK_ADDR1 = cdr.Item("VEND_BANK_ADDR1") & ""
                VEND_BANK_ADDR2 = cdr.Item("VEND_BANK_ADDR2") & ""
                VEND_BANK_ADDR3 = cdr.Item("VEND_BANK_ADDR3") & ""

                VEND_BANK_CITY = cdr.Item("VEND_BANK_CITY") & ""
                VEND_BANK_STATE = cdr.Item("VEND_BANK_STATE") & ""
                VEND_BANK_ZIP_CODE = cdr.Item("VEND_BANK_ZIP_CODE") & ""
                VEND_BANK_COUNTRY = cdr.Item("VEND_BANK_COUNTRY") & ""
                VEND_BANK_CONTACT = cdr.Item("VEND_BANK_CONTACT") & ""
                VEND_REMIT_EMAIL = cdr.Item("VEND_REMIT_EMAIL") & ""

            Else
                If VEND_ALT_CODE <> "VENDOR" And VEND_ALT_CODE <> "" Then
                    LookUp("APTVEND2", New String() {VEND_CODE_AP, VEND_ALT_CODE})
                    VEND_EMAIL = cdr.Item("VEND_ALT_EMAIL") & ""
                    VEND_CONTACT = cdr.Item("VEND_ALT_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_ALT_PHONE") & ""

                    VEND_BANK_ACCT_ID = cdr.Item("VEND_ALT_BANK_ACCT_ID") & ""
                    VEND_BANK_ROUTING_NO = cdr.Item("VEND_ALT_BANK_ROUTING_NO") & ""
                    VEND_BANK_SWIFT_NO = cdr.Item("VEND_ALT_BANK_SWIFT_NO") & ""
                    VEND_BANK_ACCT_CLASS = cdr.Item("VEND_ALT_BANK_ACCT_CLASS") & ""
                    VEND_BANK_ACCT_TYPE = cdr.Item("VEND_ALT_BANK_ACCT_TYPE") & ""

                    VEND_BANK_NAME = cdr.Item("VEND_ALT_BANK_NAME") & ""
                    VEND_BANK_ADDR1 = cdr.Item("VEND_ALT_BANK_ADDR1") & ""
                    VEND_BANK_ADDR2 = cdr.Item("VEND_ALT_BANK_ADDR2") & ""
                    VEND_BANK_ADDR3 = cdr.Item("VEND_ALT_BANK_ADDR3") & ""

                    VEND_BANK_CITY = cdr.Item("VEND_ALT_BANK_CITY") & ""
                    VEND_BANK_STATE = cdr.Item("VEND_ALT_BANK_STATE") & ""
                    VEND_BANK_ZIP_CODE = cdr.Item("VEND_ALT_BANK_ZIP_CODE") & ""
                    VEND_BANK_COUNTRY = cdr.Item("VEND_ALT_BANK_COUNTRY") & ""
                    VEND_BANK_CONTACT = cdr.Item("VEND_ALT_BANK_CONTACT") & ""
                    VEND_REMIT_EMAIL = cdr.Item("VEND_ALT_REMIT_EMAIL") & ""

                Else
                    LookUp("APTVEND1", VEND_CODE_AP)
                    VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                    VEND_ACCT_NO = cdr.Item("VEND_ACCT_NO") & ""
                    VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_PHONE") & ""

                    VEND_BANK_ACCT_ID = cdr.Item("VEND_BANK_ACCT_ID") & ""
                    VEND_BANK_ROUTING_NO = cdr.Item("VEND_BANK_ROUTING_NO") & ""
                    VEND_BANK_SWIFT_NO = cdr.Item("VEND_BANK_SWIFT_NO") & ""
                    VEND_BANK_ACCT_CLASS = cdr.Item("VEND_BANK_ACCT_CLASS") & ""
                    VEND_BANK_ACCT_TYPE = cdr.Item("VEND_BANK_ACCT_TYPE") & ""

                    VEND_BANK_NAME = cdr.Item("VEND_BANK_NAME") & ""
                    VEND_BANK_ADDR1 = cdr.Item("VEND_BANK_ADDR1") & ""
                    VEND_BANK_ADDR2 = cdr.Item("VEND_BANK_ADDR2") & ""
                    VEND_BANK_ADDR3 = cdr.Item("VEND_BANK_ADDR3") & ""

                    VEND_BANK_CITY = cdr.Item("VEND_BANK_CITY") & ""
                    VEND_BANK_STATE = cdr.Item("VEND_BANK_STATE") & ""
                    VEND_BANK_ZIP_CODE = cdr.Item("VEND_BANK_ZIP_CODE") & ""
                    VEND_BANK_COUNTRY = cdr.Item("VEND_BANK_COUNTRY") & ""
                    VEND_BANK_CONTACT = cdr.Item("VEND_BANK_CONTACT") & ""
                    VEND_REMIT_EMAIL = cdr.Item("VEND_REMIT_EMAIL") & ""
                End If
            End If

            With rowAPTCHCK5
                .Item("VEND_BANK_ACCT_ID") = VEND_BANK_ACCT_ID
                .Item("VEND_BANK_ROUTING_NO") = VEND_BANK_ROUTING_NO
                .Item("VEND_BANK_SWIFT_NO") = VEND_BANK_SWIFT_NO
                .Item("VEND_BANK_ACCT_CLASS") = VEND_BANK_ACCT_CLASS
                .Item("VEND_BANK_ACCT_TYPE") = VEND_BANK_ACCT_TYPE

                .Item("VEND_BANK_NAME") = VEND_BANK_NAME
                .Item("VEND_BANK_ADDR1") = VEND_BANK_ADDR1
                .Item("VEND_BANK_ADDR2") = VEND_BANK_ADDR2
                .Item("VEND_BANK_ADDR3") = VEND_BANK_ADDR3

                .Item("VEND_BANK_CITY") = VEND_BANK_CITY
                .Item("VEND_BANK_STATE") = VEND_BANK_STATE
                .Item("VEND_BANK_ZIP_CODE") = VEND_BANK_ZIP_CODE
                .Item("VEND_BANK_COUNTRY") = VEND_BANK_COUNTRY
                .Item("VEND_BANK_CONTACT") = VEND_BANK_CONTACT
                .Item("VEND_REMIT_EMAIL") = VEND_REMIT_EMAIL
                .Item("VEND_BANK_ACCT_ID_DECRYPTED") = ASCMAIN1.DecryptAES(VEND_BANK_ACCT_ID).ToString
            End With

            If PYMT_METHOD = "ECHECK" Or PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
                VEND_EMAIL = VEND_REMIT_EMAIL
            End If

            Dim RecordSelectionFormula As String
            RecordSelectionFormula = "{APTINVH1.CHECK_NUM} = '" & rowAPTPYMT2.Item("CHECK_NUM") & "'"
            CR_params.Add("COPY", "0")
            'Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("APRCHKP1.REMIT")
            'Generate_Report(RPT, , , RecordSelectionFormula, "PDF", FILENAME)
            Dim REPORT_NO As String = Generate_Report(RPT, , , RecordSelectionFormula, "PDF")
            Dim FILENAME As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & "." & "PDF"
            Dim MailCCList As String = rowGLTBANK1.Item("ACCT_EMAIL") & ""
            If rowGLTBANK1.Item("BANK_EMAIL") & "" <> "" Then
                If MailCCList <> "" Then
                    MailCCList &= ";"
                End If
                MailCCList &= rowGLTBANK1.Item("BANK_EMAIL") & ""
            End If

            email_to_Self(FILENAME, VEND_CODE, VEND_NAME, VEND_EMAIL, VEND_ACCT_NO, MailCCList,
                          VEND_BANK_ACCT_ID, CHECK_NUM, CHECK_DATE)

        Next
    End Sub

    Sub email_to_Self(FILENAME As String, VEND_CODE As String, VEND_NAME As String, VEND_EMAIL As String,
                      VEND_ACCT_NO As String, MailCCList As String,
                      VEND_BANK_ACCT_ID As String, CHECK_NUM As String, CHECK_DATE As Date)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing email")

        If Not ASCMAIN1.Running_in_VS Then Stop
        ' If Not ASCMAIN1.Running_in_VS Then Exit Sub
        MailCCList = "" ' while testing

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add("wjz@absolution.com", ASCMAIN1.USER_NAME)
        Else
            If VEND_EMAIL <> "" Then
                EMAIL_ADDRESSs.Add(VEND_EMAIL, VEND_NAME)
            Else
                EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            End If
            ' EMAIL_ADDRESSs.Add("wjz@absolution.com", ASCMAIN1.USER_NAME)
            ' EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            EMAIL_ADDRESSs.Add("financepay.nyc@interparfums.com", "Financepay.NYC")
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add($"Payment Remittance Advice", FILENAME)

        Dim EMAIL_SUBJECT As String = $"Payment Remittance Advice - {VEND_NAME}"

        Dim strHtml As String = ""

        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)

        Dim yellow As String = " style='background-color:yellow'"
        Dim blue As String = " style='color:blue'"
        Dim margin As String = " style='margin-left:20px'"

        strHtml = "<h1>Payment Remittance Advice</h1>"
        strHtml &= $"<br/><div>Vendor Contacts who should receive this email<br/>{Replace(Mid(VEND_EMAIL, 1), ";", "<br/>")}</div><br/>"

        If Now.Hour < 12 Then
            strHtml &= $"<br/><div>Good Morning</div>"
        Else
            strHtml &= $"<br/><div>Good Afternoon</div>"
        End If

        strHtml &= $"<br/><div>Attached is a report detailing Invoices which have been paid:</div>"
        strHtml &= $"<ul>"
        strHtml &= $"<li><strong>Payment was made on {Format(CHECK_DATE, "MM/dd/yyyy")}</strong>.</li>"

        Dim VEND_BANK_ACCT_ID_decrypted = ASCMAIN1.DecryptAES(VEND_BANK_ACCT_ID)
        If VEND_BANK_ACCT_ID_decrypted.LENGTH < 5 Then
            VEND_BANK_ACCT_ID_decrypted = "*****" & VEND_BANK_ACCT_ID_decrypted
        End If

        Dim VEND_BANK_ACCOUNT_ID_prt As String = "******" & Mid(VEND_BANK_ACCT_ID_decrypted, VEND_BANK_ACCT_ID_decrypted.Length - 3, 4)
        strHtml &= $"<li>Payment was made to your Bank Account {VEND_BANK_ACCOUNT_ID_prt}</li>"
        strHtml &= $"</ul>"

        If VEND_ACCT_NO <> "" Then
            strHtml &= $"<br/><div>Account No <strong{blue}>{VEND_ACCT_NO}</strong></div>"
        End If

        strHtml &= $"<br/><div>Thank You,</div><br/>" ' & TAC.POCMAIN1.GetUserSignature(rowASTUSER1)

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "REMITA", True, False, CHECK_NUM, "CHECK_NUM", "Payment Remittance Advice", strHtml)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'MsgBox("email Sent", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Private Sub cmbBATCH_NO_PYMT_ValueChanged(sender As Object, e As EventArgs) Handles cmbBATCH_NO_PYMT.ValueChanged

        txtBankCode.Clear()
        txtPaymentMethod.Clear()

        Dim BATCH_NO_PYMT As String = cmbBATCH_NO_PYMT.Text
        Dim rowAPTPYMT1 As DataRow = LookUp("APTPYMT1", BATCH_NO_PYMT)
        If rowAPTPYMT1 IsNot Nothing Then
            txtBankCode.Text = rowAPTPYMT1.Item("BANK_CODE") & String.Empty
            txtPaymentMethod.Text = rowAPTPYMT1.Item("PYMT_METHOD") & String.Empty
        End If

    End Sub
End Class