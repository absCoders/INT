Public Class TACMAIN1

    ' nSoftware License Keys
    'Public nSoftwareZipkey As String = "315A4E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004A424A545848315458354B300000"
    Public Shared nSoftwareZipkey As String = "315A4E46414431535542323032333033313352415331544531414D48313432360000000000000000574A5A3156305333000044365934463431435057394D0000"
    Public Shared nSoftwareIPWorksV9Key As String = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
    Public Shared nSoftwareIPWorksV2020Key As String = "31504E46414431535542323032333033313352415331544531414D48313432360000000000000000574A5A31563053330000334D4256385A5656584735450000"
    Public Shared nSoftwareftpkey As String = nSoftwareIPWorksV2020Key
    Public Shared nSoftwareipportkey As String = nSoftwareIPWorksV2020Key
    Public Shared nSoftwarepopkey As String = nSoftwareIPWorksV2020Key
    Public Shared nSoftwarehttpkey As String = nSoftwareIPWorksV2020Key
    'Public Shared nSoftwareInship As String = "42584E354141315355425241533154453345383933333331580000000000000000000000000000004A52344B5057583900003059573859305A4A545958520000"
    Public Shared s4DPaymentsShippingSDK As String = "44584E4A41414E58524633545447303532320052534A5A514744574E49514157524D434E43424F003432454E54583931000048464A365A47525055394A5A0000"
    Public Shared nSoftwareEncryptkey As String = "31454E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004E4B54534E383157353733320000"
    'Public Shared nSoftwaresftpkey As String = "31484E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D5446000044483650384E5454444B4D4B0000"
    Public Shared nSoftwaresftpkey As String = "31484E46414431535542323032333033313352415331544531414D48313432360000000000000000574A5A31563053330000445A304339303631355534390000"
    'Public Shared nSoftwareEncryptionkey As String = "31454E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004E4B54534E383157353733320000"
    Public Shared nSoftwareEncryptionkey As String = "31454E46414431535542323032333033313352415331544531414D48313432360000000000000000574A5A31563053330000385A58474E46455A375755420000"
    Public Shared nSoftwareInPay As String = "42504E364141315355425241533154453345383933333331580000000000000000000000000000004532594A58424252000032455A374B3543414D5841330000"

    Public Overridable Sub Site_Specific_Settings()

    End Sub

    Public Overridable Sub Get_Column_Expression_Exceptions(ByVal FORM_NAME As String, ByVal DATA_SOURCE As String, ByVal COLUMN_NAME As String, ByRef sql_SELECT_col As String) ' , ByRef sql_GROUP_BY_col As String)

    End Sub

    Public Overridable Function Get_Code_SQL_X(ByVal FORM_NAME As String, ByVal COLUMN_NAME As String, ByRef GROUP_KEY As String) As String
        Return Nothing
    End Function

    Public Overridable Sub Write_Group_Record_X(ByVal GROUP_KEY As String, ByVal COLUMN_NAME As String, ByVal GROUP_CODEs As ArrayList, ByVal GROUP_DESCs As ArrayList)

    End Sub

    Public Overridable Function CodeValues(ByVal TABLE_COLUMN As String) As Dictionary(Of String, String)
        Return Nothing
    End Function

    Public Overridable Function Send_email(ByVal frmASFBASE0 As ASFBASE0,
                                 ByVal EMAIL_ADDRESSs As Dictionary(Of String, String),
                                 ByVal ATTACHMENTs As Dictionary(Of String, String),
                                 ByVal SUBJECT As String,
                                 ByVal EMAIL_KEY As String,
                                 Optional ByVal auto_send As Boolean = False,
                                 Optional SEND_CC_to_USER_ID As Boolean = False,
                                 Optional ENTITY_KEY As String = "",
                                 Optional ENTITY_NAME As String = "",
                                 Optional ENTITY_CAPTION As String = "",
                                 Optional EMAIL_BODY As String = "") As String
        Return Nothing
    End Function

    Public Overridable Sub Application_Initialization()

    End Sub

    Public Overridable Sub Maintain_Contacts(ByVal frmASFBASE1 As ASFBASE1, _
                                           ByVal CONTACT_ENTITY_TABLE As String, _
                                           ByVal CONTACT_ENTITY_KEY As String, _
                                           ByVal CONTACT_ENTITY_NAME As String)

    End Sub

    Public Overridable Function Custom_sqlwhere( _
    ByVal sqlwhere As String, _
    ByVal grd As UltraWinGrid.UltraGrid, _
    ByVal COLUMN_NAME As String) As String
        Return sqlwhere
    End Function

    Public Sub Record_Event( _
        ByVal TABLE_NAME As String, _
        ByVal TABLE_KEY As String, _
        ByVal INIT_DATE As Date, _
        ByVal INIT_OPER As String, _
        ByVal EVENT_TYPE As String, _
        ByVal EVENT_DESC As String, _
        Optional ByVal EVENT_KEY As String = "", _
        Optional FORM_NAME As String = "")

        If FORM_NAME = "" Then
            FORM_NAME = ASCMAIN1.ActiveForm.Name
        End If

        Dim SELECTION_NO As String = ""
        Dim XNO As String = ""

        If ASCMAIN1.ActiveForm IsNot Nothing Then
            SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
            XNO = ASCMAIN1.ActiveForm.XNO
        End If

        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME, SESSION_NO, SELECTION_NO, XNO) " _
                             & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8,:PARM9,:PARM10,:PARM11)", _
                             "VVDVVVVVVVV", _
                             New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME, ASCMAIN1.SESSION_NO, SELECTION_NO, XNO})

    End Sub
End Class
