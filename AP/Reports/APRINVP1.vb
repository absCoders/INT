Imports System.Math

Public Class APRINVP1
    Dim APTACRC1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        With dst
            RWU = "R"

            sql = ""
            ' sql &= SQL_in("VEND_CODE", "APTACRC1.VEND_CODE_ACC")

            ' APTACRC1

            ASCMAIN1.sql = "SELECT APTACRC1.* " _
             & " FROM APTACRC1 where INV_PRINT_IND = '0'"
            APTACRC1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("ALTER TABLE " & APTACRC1 & " ADD PRIMARY KEY (CTL_NO)")

            ASCMAIN1.sql = "SELECT APTACRC1.*, 'Z' AR_PARM_KEY FROM " & APTACRC1 & " APTACRC1"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTACRC1", 1))
            ' .Tables("APTACRC1").Columns.Add("AR_PARM_KEY")


            ASCMAIN1.sql = "SELECT APTINVH1.* " _
             & " from APTINVH1 where VOUCHER_NO in (Select distinct VOUCHER_NO_ORIG from " & APTACRC1 & ")"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTINVH1", 1))

            ASCMAIN1.sql = "SELECT * FROM APTVEND1 WHERE VEND_CODE IN (SELECT DISTINCT VEND_CODE_ACC FROM " & APTACRC1 & ")"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTVEND1", 1))

      
            'For Each TABLE_NAME As String In New String() _
            '{"TATTERM1", "ICTWHSE1", "SOTSVIA1", "SOTSREP1", "SOTREAS1", "ARTREAS1", "SOTSDIV1", "TATCNTRY"}
            '    Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
            '    Fill_Records(TABLE_NAME)
            'Next

            With .Tables.Add("SOTINVP0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("ADDRESS0")
                .Columns.Add("ADDRESS1")
                .Columns.Add("ADDRESS2")
                .Columns.Add("ADDRESS3")
                .Columns.Add("AR_PARM_DUNS_NO")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() {.Columns("AR_PARM_KEY")}
            End With
        End With

        Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
        With ROWs("ARTPARM1")
            rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
            rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_ADDR2") & ""
            rowSOTINVP0.Item("REMIT3") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            ' rowSOTINVP0.Item("REMIT3") = "Tel " & .Item("AR_PARM_REMIT_PHONE") & " Fax " & .Item("AR_PARM_REMIT_FAX")
            ' rowSOTINVP0.Item("REMIT3") = ""
            rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
            rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
        End With

        With ASCMAIN1.rowASTPARM1
            rowSOTINVP0.Item("ADDRESS0") = .Item("AS_PARM_INST_NAME") & ""
            rowSOTINVP0.Item("ADDRESS1") = .Item("AS_PARM_INST_ADDR1") & ""
            rowSOTINVP0.Item("ADDRESS2") = .Item("AS_PARM_INST_ADDR2") & ""
            rowSOTINVP0.Item("ADDRESS3") = .Item("AS_PARM_INST_CITY") & ", " _
                    & .Item("AS_PARM_INST_STATE") & " " _
                    & .Item("AS_PARM_INST_ZIP_CODE") & " " _
                    & .Item("AS_PARM_INST_COUNTRY")

            Dim TEL As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE")
            If TEL.Length = 10 Then
                TEL = "(" & Mid(TEL, 1, 3) & ")" & Mid(TEL, 4, 3) & "-" & Mid(TEL, 7, 4)
            End If
            Dim FAX As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX")
            If FAX.Length = 10 Then
                FAX = "(" & Mid(FAX, 1, 3) & ")" & Mid(FAX, 4, 3) & "-" & Mid(FAX, 7, 4)
            End If
            rowSOTINVP0.Item("ADDRESS3") = "P " & TEL & " F " & FAX
            ' rowSOTINVP0.Item("ADDRESS3") = ""
        End With

        rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)
 

        If dst.Tables("APTACRC1").Rows.Count = 0 Then
            RWU &= "0"
            xErrMsg = "No Eligible Records"
        End If
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        RPT = "APRINVP1"

        'Dim rowAPTPARM1 As DataRow = LookUp("APTPARM1", "Z")
        'Dim Telephone As String = IIf(rowAPTPARM1.Item("AP_PARM_REMIT_PHONE") & "" = "", "", rowAPTPARM1.Item("AP_PARM_REMIT_PHONE"))
        'If Telephone <> "" Then
        '    Telephone = Mid(Telephone, 1, 3) & "." & Mid(Telephone, 4, 3) & "." & Mid(Telephone, 7, 4)
        'End If
        'Dim Fax As String = IIf(rowAPTPARM1.Item("AP_PARM_REMIT_FAX") & "" = "", "", rowAPTPARM1.Item("AP_PARM_REMIT_FAX"))
        'If Fax <> "" Then
        '    Fax = Mid(Fax, 1, 3) & "." & Mid(Fax, 4, 3) & "." & Mid(Fax, 7, 4)
        'End If


        'CR_params.Add("AP_PARM_REMIT_NAME", rowAPTPARM1.Item("AP_PARM_REMIT_NAME") & "")
        'CR_params.Add("AP_PARM_REMIT_ADDR1", rowAPTPARM1.Item("AP_PARM_REMIT_ADDR1") & "")
        'CR_params.Add("AP_PARM_REMIT_ADDR2", rowAPTPARM1.Item("AP_PARM_REMIT_ADDR2") & "")
        'CR_params.Add("AP_PARM_REMIT_CITY", rowAPTPARM1.Item("AP_PARM_REMIT_CITY") & "")
        'CR_params.Add("AP_PARM_REMIT_STATE", rowAPTPARM1.Item("AP_PARM_REMIT_STATE") & "")
        'CR_params.Add("AP_PARM_REMIT_ZIP_CODE", rowAPTPARM1.Item("AP_PARM_REMIT_ZIP_CODE") & "")
        'CR_params.Add("AP_PARM_REMIT_PHONE_FAX", "Telephone: " & Telephone & "  Fax: " & Fax)

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.sql = "UPDATE APTACRC1" _
            & " SET INV_PRINT_IND = '1', INV_PRINT_DATE = SYSDATE, INV_PRINT_USER = '" & ASCMAIN1.USER_ID & "'" _
            & "  WHERE CTL_NO IN (SELECT CTL_NO FROM " & APTACRC1 & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub

End Class