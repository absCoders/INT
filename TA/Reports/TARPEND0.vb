Public Class TARPEND0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "U"

    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" = "1" Then
                EMsg = EMsg & vbCr & "Period-End Already Initialized"
            Else
                Dim YYYY As String
                For i As Integer = 0 To 1
                    YYYY = Format$(Val(Mid$(ASCMAIN1.CYP, 1, 4)) + i, "0000")
                    ASCMAIN1.sql = "Select Count (*) from GLTPARM2 where OPS_YYYYPP like '" & YYYY & "%'"
                    If Val(ASCDATA1.GetDataValue & "") <> 12 Then
                        EMsg = EMsg & vbCr & "Please Check Operations Calendar for " & YYYY
                    End If
                Next
            End If

            If EMsg = "" Then

                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

                ASCMAIN1.sql = "Select ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                    & " from ARTPYMT2,ARTPYMT1 where ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                    & " and ARTPYMT1.PYMT_APPL_ONLY = '1' and ARTPYMT2.PYMT_STATUS = '1'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim PYMT_BATCH_NO As String = row.Item("PYMT_BATCH_NO")
                    Dim PYMT_BATCH_LNO As Integer = Val(row.Item("PYMT_BATCH_LNO") & "")
                    ARCMAIN1.Clean_Out_ARTPYMT2_Started_NOT_Completed(PYMT_BATCH_NO, PYMT_BATCH_LNO)
                Next

                Check_for_Records("APTPYMT1", "AP Payment Selection Batch")
                Check_for_Records("ARTPYMT1", "AR Payment Application Journal", "STATUS = '1'")
                Check_for_Records("APTINVH1", "AP Vendor Invoice Register", "REGISTER_IND = '0'")
                Check_for_Records("APTCHCK1", "AP Check Register", "REGISTER_IND  = '0' or (OPS_YYYYPP_F is Not Null and REGISTER_IND_F = '0')")
                Check_for_Records("SOTINVH1", "Sales Journal", "ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "' and NVL(REGISTER_XNO,'0') = '0'")

                Check_for_Records("ICTITEM1", "Items without Standard Costs", "ITEM_COST_STATUS = 'P' and ITEM_CODE in (Select Distinct ITEM_CODE from ICTSTAT1 where OPS_YYYYPP >= '" & ASCMAIN1.CYP & "')")
                'month-end - require idoc import

                ' Check_for_Records("ICTIREC1", "PO Receipts Journal", "REGISTER_IND = '0'")
                ' Check_for_Records("ICTIADJ1", "Inventory Adjustments Journal", "REGISTER_IND = '0'")
                ' Check_for_Records("ICTIXFR1", "Warehouse Transfer Journal", "REGISTER_IND = '0'")

                ' CHECK FOR ICTWHSE1 WITH PI IN PROCESS

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Check_for_Records("EDTTRXN1", "3PL Receipts Pending Update", "PROCESS_IND = '0' and TRANS_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'")

                    '  Check_for_Records("EDT180T1", "EDI 180 Documents Pending", "EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' and TRANS_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'")
                    Check_for_Records("EDT812T1", "EDI 812 Documents Pending", "EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' and EDI_CLAIM_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'")
                    ASCMAIN1.sql = "Select .* from EDT180T1 where "

                    TAC.TACMAIN1.Get_Unprocessed_IDOCs(Me)
                    If dst.Tables("TATIDOCU").Select("INV_DATE <= #" & Format(PRD_END_DATE, "MM/dd/yyyy") & "#").Length > 0 Then
                        EMsg &= vbCr & "IDOCs awaiting Import (see IDOCs tab of PO Anticipated Receipts Entry)"
                    End If

                    'Dim sftp_folder As String = "" _
                    '  & IIf(ASCMAIN1.Running_in_VS, "C:\Users\wjz\Desktop\Interparfums\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
                    '  & "\IPSA\" _
                    '  & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
                    '  & "\FROM_IPSA\IDOC\"

                    'If My.Computer.FileSystem.GetFiles(sftp_folder).Count > 0 Then
                    '    EMsg &= vbCr & "IDOCs awaiting Import (see IDOCs tab of PO Anticipated Receipts Entry)"
                    'End If


                    'Dim FOLDERNAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\" & "COWORX" & "\" & "FromCoworx"
                    'If My.Computer.FileSystem.GetFiles(FOLDERNAME).Count > 0 Then
                    '    EMsg &= vbCr & "RSC or Invoice Files awaiting to be Loaded"
                    'End If
                
                End If


                TAC.TACMAIN1.Update_Forex()

                If EMsg <> "" Then
                    EMsg = "Cannot Proceed because a Clean Cut-off has not been established as follows:" & vbCr & EMsg
                End If
            End If
        End If

    End Sub

    Sub Check_for_Records( _
    ByVal TABLE_NAME As String, _
    ByVal TABLE_DESC As String, _
    Optional ByVal where_clause As String = "", _
    Optional ByVal custom_sql As String = "")

        If custom_sql <> "" Then
            ASCMAIN1.sql = custom_sql
        Else
            ASCMAIN1.sql = "Select count (*) from " & TABLE_NAME
            If where_clause <> "" Then
                ASCMAIN1.sql &= " where " & where_clause
            End If
        End If

        Dim sql As String = ASCMAIN1.sql
        Dim r As Long = Val(ASCDATA1.GetDataValue() & "")
        If r <> 0 Then
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(r) & " Records"
        End If

    End Sub

    Overrides Sub Update_Record()

        'ASCMAIN1.sql = "Select INV_TYPE, INV_NO from SOTINVH1 where ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "' and NVL(REGISTER_IND,'0') = '0'"

        ASCMAIN1.sql = "Update ASTPCTL1 set PRD_CLOSE_IND = '1'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN SAPSSUMX('" & ASCMAIN1.CYP & "'); END;"
        ASCDATA1.ExecuteSQL()

        ' Create_Monthly_Backup_Script()
        Create_Monthly_Backup_Script2()
    End Sub

    Sub Create_Monthly_Backup_Script()

        'Dim FOLDER As String = "S:\" & ASCMAIN1.DBS_COMPANY & "\MONTHLY\"
        Dim FOLDER As String = ASCMAIN1.Folders("SharedRoot") & "MONTHLY\"
        If ASCMAIN1.Running_in_VS Then
            FOLDER = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "\MONTHLY\"
            Stop
        End If
        If Not My.Computer.FileSystem.DirectoryExists(FOLDER) Then
            My.Computer.FileSystem.CreateDirectory(FOLDER)
        End If

        Dim FILENAME As String = FOLDER & "MONTHLY.BAT"
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            My.Computer.FileSystem.DeleteFile(FILENAME)
        End If

        'after bat file exported we could not run bat file when it was pointing to e drive until
        Dim MMM As String = Mid(Format(Now, "MMMM"), 1, 3).ToUpper
        MMM = ASCMAIN1.CYP
        Dim AS_PARM_PEND_BACKUP_FOLDER As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_PEND_BACKUP_FOLDER") & ""
        If AS_PARM_PEND_BACKUP_FOLDER = "" Then
            AS_PARM_PEND_BACKUP_FOLDER = "C:\"
        Else
            If Not AS_PARM_PEND_BACKUP_FOLDER.EndsWith("\") Then
                AS_PARM_PEND_BACKUP_FOLDER &= "\"
            End If
        End If
        Dim F As String = AS_PARM_PEND_BACKUP_FOLDER & ASCMAIN1.DBS_COMPANY & "\" & MMM
        Using SW As New System.IO.StreamWriter(FILENAME)
            SW.WriteLine("exp " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_COMPANY & " file=" & F & ".dmp log=" & F & ".log")
            SW.WriteLine(Chr(34) & "C:\Program Files\WinZip\winzip64" & Chr(34) & " -m " & F & ".zip " & F & ".dmp " & F & ".log")
        End Using
    End Sub

    Sub Create_Monthly_Backup_Script2()

        If ASCMAIN1.Running_in_VS Then
            Stop

        End If
        Dim FOLDER As String = ASCMAIN1.Folders("SharedRoot") & "MONTHLY\"
        Dim FILENAME As String = FOLDER & "MONTHLY.BAT"

        If My.Computer.FileSystem.FileExists(FILENAME) Then
            My.Computer.FileSystem.DeleteFile(FILENAME)
        End If

        Using SW As New System.IO.StreamWriter(FILENAME)
            SW.WriteLine(FOLDER & "MONTHLY_BACKUP.BAT " & ASCMAIN1.CYP)
            SW.WriteLine("PAUSE")
            SW.WriteLine("EXIT")
        End Using
    End Sub

End Class