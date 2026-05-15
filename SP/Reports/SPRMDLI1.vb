Imports nsoftware.IPWorksSFTP

Public Class SPRMDLI1

    ' MT

#Region "General Declarations"
    Dim SPTDCOMA As String ' Drives Report
    Dim SPTDCOMB As String ' Detailed to Store
    ' Dim SPTDCOMC As String ' Credit Accrual
    Dim sqlSPTDCOMA As String
    Dim sqlSPTDCOMB As String
    Dim sqlSPTDCOMC As String

    Dim files As New List(Of String)

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)

        ' Me.Sftp1.SSHHost.
    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"

        Retrieve_Files()

        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("SPTDCOMA")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = ""

        CR_params.Add("SUBT", SUBT)
        Generate_Report(RPT, , SUBT)

        Print_GL()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.sql = "Insert into SPTDCOMB Select * from " & SPTDCOMB
        ASCDATA1.ExecuteSQL(sql)

        For Each row As DataRow In dst.Tables("SPTDCOMC").Select
            row.Item("ACC_CTL_NO") = ASCMAIN1.Next_Control_No("SPTDCOMC.ACC_CTL_NO")
            row.AcceptChanges()
            row.SetAdded()
        Next
        Update_Record_TDA("SPTDCOMC")

        'ASCMAIN1.sql = "Insert into SPTDCOMC Select * from " & SPTDCOMC
        'ASCDATA1.ExecuteSQL(sql)

        GL_Update()
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select SPTDCOMA.*, ARTCUST1.TRADE_CLASS_CODE from " & SPTDCOMA & " SPTDCOMA, ARTCUST1 where ARTCUST1.CUST_CODE = SPTDCOMA.CUST_CODE"
            Create_TDA(.Tables.Add, "SPTDCOMA", "**", 0, False)

            For Each TABLE_NAME As String In New String() _
                {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1"}
            Fill_Records(TABLE_NAME)
        Next

        If perform_fill Then
            Fill_Records_RPT(sqlw)
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        If parms.Length > 0 Then
            sqlw = parms(0)
        End If

        Fill_Records("GLTPARM3", RYP)
        Fill_Records("GLTPARM2", RYP)

        EnforceConstraints(False)
        Fill_Records("SPTDCOMA")
        EnforceConstraints(True)

        'Prepare_GL_Interface("SPMA")

    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Vehicle, Trade Class and Brand

        Dim TOTAL_DEMO_COMM As Decimal = 0

        For Each row As DataRow In
            ASCDATA1.SelectDistinct(dst.Tables("SPTDCOMA"), New String() {"TRADE_CLASS_CODE", "COLLECTION_CODE"}).Select

            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

            Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

            Dim BRAND_CODE As String = rowICTCOLL1.Item("BRAND_CODE")
            Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(BRAND_CODE)

            Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("SPTDCOMA").Compute("SUM(AMT_COMM)", "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "' and COLLECTION_CODE = '" & COLLECTION_CODE & "'") & "")

            If DETL_POSTING_AMT <> 0 Then
                TOTAL_DEMO_COMM += DETL_POSTING_AMT
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("OPS_YYYYPP") = RYP
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_DEMO_ACCT_CODE_EXP")
                If rowSOTTCLS1 IsNot Nothing Then
                    If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then
                        rowGLTINTF1.Item("SEG3_CODE") = rowSOTTCLS1.Item("SEG3_CODE")
                    Else
                        rowGLTINTF1.Item("SEG3_CODE") = TRADE_CLASS_CODE
                    End If
                Else
                    rowGLTINTF1.Item("SEG3_CODE") = "?"
                End If
                'If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then rowGLTINTF1.Item("SEG4_CODE") = rowSOTTCLS1.Item("SEG4_CODE")
                If rowICTBRAN1 IsNot Nothing Then
                    If rowICTBRAN1.Item("SEG4_CODE") & "" <> "" Then
                        rowGLTINTF1.Item("SEG4_CODE") = rowICTBRAN1.Item("SEG4_CODE")
                    Else
                        rowGLTINTF1.Item("SEG4_CODE") = BRAND_CODE ' TEMP UNTIL WE MAKE SURE WE HAVE SEG4S CREATED IN ICTBRAN1 OR ICTCOLL1, AND THEN IN GLTSEGM1
                    End If
                Else
                    rowGLTINTF1.Item("SEG4_CODE") = "?"
                End If
                If rowGLTINTF1.Item("SEG2_CODE") & "" = "" Then rowGLTINTF1.Item("SEG2_CODE") = "?"
                If rowGLTINTF1.Item("SEG3_CODE") & "" = "" Then rowGLTINTF1.Item("SEG3_CODE") = "?"
                If rowGLTINTF1.Item("SEG4_CODE") & "" = "" Then rowGLTINTF1.Item("SEG4_CODE") = "?"
                rowGLTINTF1.Item("DETL_CVX_NO") = COLLECTION_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"
            End If
        Next

        ' Accrual


        If TOTAL_DEMO_COMM <> 0 Then
            Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * TOTAL_DEMO_COMM)
            rowGLTINTF1.Item("OPS_YYYYPP") = RYP
            rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_DEMO_ACCT_CODE_ACC")
        End If

        Return JOURNAL_NO
    End Function

    Function Write_GLTINTF1(JOURNAL_TYPE As String, JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer, DETL_CTL_DATE As Date, DETL_POSTING_AMT As Decimal)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP0
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ""
        rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
        rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables()

    End Sub

    Sub Retrieve_Files()
        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
        Sftp1.SSHUser = "IPLB715"
        Sftp1.SSHPassword = "wVH37Ev"
        Sftp1.SSHHost = "sftp.coworx.net"
        ' Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SCPSSHAuthModes.amPassword

        'Sftp1.SSHAcceptAnyServerHostKey = True
        Sftp1.RemotePath = "FromCoworx"

        Sftp1.SSHLogon(Sftp1.SSHHost, Sftp1.SSHPort)
        files.Clear()
        Sftp1.ListDirectory()

        For Each file As String In Sftp1.DirList
            'Debug.Print(e.FileName)
            'Debug.Print(e.DirEntry)

            If file.EndsWith(".csv") Then
                files.Add(file)
            End If
        Next

        If files.Count > 0 Then
            For Each filename As String In files
                Sftp1.LocalFile = ASCMAIN1.Folders("Work") & filename
                Sftp1.RemoteFile = filename
                Sftp1.Download()
            Next
        End If

        Sftp1.SSHLogoff()

    End Sub

    'Private Sub Sftp1_OnDirList(sender As Object, e As nsoftware.IPWorksSSH.SftpDirListEventArgs) Handles Sftp1.OnDirList
    '    'Debug.Print(e.FileName)
    '    'Debug.Print(e.DirEntry)

    '    If e.FileName.EndsWith(".csv") Then
    '        files.Add(e.FileName)
    '    End If
    'End Sub

    Private Sub Sftp1_OnSSHServerAuthentication(sender As Object, e As nsoftware.IPWorks.SCPSSHServerAuthenticationEventArgs) Handles Sftp1.O
        e.Accept = True
    End Sub

End Class