Imports System.Text

Public Class ASTUSER1
    Dim SECURITY_CODEs As New List(Of String)
    Dim rowASTPARMP As DataRow
    Dim USER_PASSWORD_orig As String

    ' LIST IF USERS IN ASTUSERX - WE PROB NEED SIMILAR FOR GROUPS

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "TATALRT1", "*")

            ASCMAIN1.sql = "Select USER_ID, USER_NAME, USER_STATUS from ASTUSER1"
            Create_TDA(.Tables.Add, "ASTUSERX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SECURITY_CODE from ASTSECM1 order by SECURITY_CODE"

            For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTSECM1").Rows
                Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
                Dim dc As New DataColumn
                dc.ColumnName = SECURITY_CODE
                dc.DataType = GetType(System.Boolean)
                dc.DefaultValue = False
                dst.Tables("ASTUSERX").Columns.Add(dc)
                SECURITY_CODEs.Add(SECURITY_CODE)
            Next

            grdASTUSERX.DataSource = dst.Tables("ASTUSERX")

            For Each SECURITY_CODE As String In SECURITY_CODEs
                With grdASTUSERX.DisplayLayout.Bands(0).Columns(SECURITY_CODE)
                    .Width = 50
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .CellAppearance.TextHAlign = HAlign.Center
                    .Header.Appearance.TextHAlign = HAlign.Center
                End With
            Next

            Fill_ASTUSERX()

            ASCMAIN1.sql = "SELECT ASTUSER2.*, ASTSECM1.SECURITY_DESC, '1' SEL " _
            & " FROM ASTUSER2,ASTSECM1 where ASTSECM1.SECURITY_CODE = ASTUSER2.SECURITY_CODE"

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER2.SECURITY_CODE,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
                & " FROM ASTUSER1, ASTSECM1 " _
                & " LEFT OUTER JOIN ASTUSER2 ON ASTUSER2.USER_ID = @PARM1" _
                & " AND ASTUSER2.SECURITY_CODE = ASTSECM1.SECURITY_CODE" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER2.USER_ID,NULL,'0','1') SEL " _
                & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
                & " FROM ASTSECM1, ASTUSER2, ASTUSER1 " _
                & " WHERE ASTUSER2.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER2.SECURITY_CODE (+) = ASTSECM1.SECURITY_CODE " _
                & " AND ASTUSER1.USER_ID = :PARM2"
            End If
            Create_TDA(.Tables.Add, "ASTUSER2", "**", 0, True, "VV", -1)



            ASCMAIN1.sql = "SELECT ASTUSER3.*, ASTUSER1.USER_NAME USER_GROUP_NAME, '1' SEL " _
            & " FROM ASTUSER3,ASTUSER1 where ASTUSER3.USER_GROUP_ID = ASTUSER1.USER_ID and ASTUSER1.USER_STATUS = 'G'"

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER3.USER_GROUP_ID,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", ASTUSERG.USER_ID USER_GROUP_ID, ASTUSERG.USER_NAME USER_GROUP_NAME " _
                & " FROM ASTUSER1, ASTUSER1 ASTUSERG " _
                & " LEFT OUTER JOIN ASTUSER3 ON ASTUSER3.USER_ID = @PARM1" _
                & " AND ASTUSER3.USER_GROUP_ID = ASTUSERG.USER_ID" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER3.USER_ID,NULL,'0','1') SEL " _
                & ", ASTUSERG.USER_ID USER_GROUP_ID, ASTUSERG.USER_NAME USER_GROUP_NAME" _
                & " FROM ASTUSER1 ASTUSERG, ASTUSER3, ASTUSER1 " _
                & " WHERE ASTUSER3.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER3.USER_GROUP_ID (+) = ASTUSERG.USER_ID " _
                & " AND ASTUSERG.USER_STATUS = 'G'" _
                & " AND ASTUSER1.USER_STATUS <> 'G'" _
                & " AND ASTUSER1.USER_ID = :PARM2"
            End If
            Create_TDA(.Tables.Add, "ASTUSER3", "**", 0, True, "VV", -1)




            ASCMAIN1.sql = "SELECT ASTUSER4.*, GLTCOMP1.COMPANY_NAME, '1' SEL " _
            & " FROM ASTUSER4,GLTCOMP1 where GLTCOMP1.COMPANY_CODE = ASTUSER4.COMPANY_CODE"

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", CASE ISNULL(ASTUSER4.USER_ID,'') WHEN '' THEN '0' ELSE '1' END SEL " _
                & ", GLTCOMP1.COMPANY_CODE, GLTCOMP1.COMPANY_NAME " _
                & " FROM ASTUSER1, GLTCOMP1 " _
                & " LEFT OUTER JOIN ASTUSER4 ON ASTUSER4.USER_ID = @PARM1" _
                & " AND ASTUSER4.COMPANY_CODE = GLTCOMP1.COMPANY_CODE" _
                & " WHERE ASTUSER1.USER_ID = @PARM2"
            Else
                ASCMAIN1.sql = "SELECT ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER4.USER_ID,NULL,'0','1') SEL " _
                & ", GLTCOMP1.COMPANY_CODE, GLTCOMP1.COMPANY_NAME " _
                & " FROM GLTCOMP1, ASTUSER4, ASTUSER1 " _
                & " WHERE ASTUSER4.USER_ID (+) = :PARM1 " _
                & " AND ASTUSER4.COMPANY_CODE (+) = GLTCOMP1.COMPANY_CODE " _
                & " AND ASTUSER1.USER_ID = :PARM2" _
                & " AND NVL(GLTCOMP1.NO_LOGIN,'0') = '0'"
            End If
            Create_TDA(.Tables.Add, "ASTUSER4", "**", 0, True, "VV", -1)


            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE " _
                & " FROM ASTOPST1 where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 GROUP BY CONVERT(DATE,INIT_DATE)"
            Else
                ASCMAIN1.sql = "SELECT DISTINCT TRUNC(INIT_DATE) STAT_DATE" _
                & " FROM ASTOPST1 where USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE)" '  and SELECTION_NO = 0 
            End If
            Create_TDA(.Tables.Add, "ASTOPST0", "**", 0, False, "VDD", 1)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
                & " , ASTOPST1.INIT_DATE, ASTOPST1.LAST_DATE " _
                & " FROM ASTOPST1 " _
                & " where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3" '  and SELECTION_NO = 0
            Else
                ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
                & " , MIN (ASTOPST1.INIT_DATE) INIT_DATE, MAX (ASTOPST1.LAST_DATE) LAST_DATE " _
                & " FROM ASTOPST1 " _
                & " where ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE), ASTOPST1.SESSION_NO" '  and SELECTION_NO = 0
            End If
            Create_TDA(.Tables.Add, "ASTOPST1", "**", 0, False, "VDD", 0)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.* " _
                & " , ASTMENU1.MENU_ITEM_DESC " _
                & " FROM ASTOPST1,ASTMENU1 " _
                & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
                & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
                & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
                & "   and ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 and SELECTION_NO <> 0"
            Else
                ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.* " _
                & " , ASTMENU1.MENU_ITEM_DESC " _
                & " FROM ASTOPST1,ASTMENU1 " _
                & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
                & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
                & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
                & "   and ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 and SELECTION_NO <> 0"
            End If
            Create_TDA(.Tables.Add, "ASTOPSTX", "**", 0, False, "VDD", 0)

            Create_Relation("ASTOPST0", "ASTOPST1", "STAT_DATE")
            Create_Relation("ASTOPST1", "ASTOPSTX", "STAT_DATE,SESSION_NO")

            '.Relations.Add("ASTOPST1", _
            'New DataColumn() {.Tables("ASTOPST0").Columns("STAT_DATE")}, _
            'New DataColumn() {.Tables("ASTOPST1").Columns("STAT_DATE")})

            '.Relations.Add("ASTOPSTX", _
            'New DataColumn() {.Tables("ASTOPST1").Columns("SESSION_NO")}, _
            'New DataColumn() {.Tables("ASTOPSTX").Columns("SESSION_NO")})

            .Tables("ASTOPST0").Columns.Add("SESSIONS", GetType(System.Int64), "COUNT(CHILD.SESSION_NO)")
            .Tables("ASTOPST1").Columns.Add("SELECTIONS", GetType(System.Int64), "COUNT(CHILD.SELECTION_NO)")
            .Tables("ASTOPST0").Columns.Add("SELECTIONS", GetType(System.Int64), "SUM(CHILD.SELECTIONS)")


            ASCMAIN1.sql = "SELECT X.*, Y.MENU_ITEM_DESC from" _
                & " (Select MENU_ITEM_OBJECT, MIN (MENU_ITEM_DESC) MENU_ITEM_DESC" _
                & " from ASTMENU1 group by MENU_ITEM_OBJECT) Y," _
                & " (Select MENU_ITEM_OBJECT, COUNT (*) RUNS, MIN (INIT_DATE) INIT_DATE, MAX (INIT_DATE) LAST_DATE" _
                & ", MIN (MENU_ID) MENU_ID1, MAX (MENU_ID) MENU_ID2" _
                & " from ASTOPST1" _
                & " where USER_ID = :PARM1 and INIT_DATE > :PARM2 and INIT_DATE < :PARM2" _
                & " group by MENU_ITEM_OBJECT) X where Y.MENU_ITEM_OBJECT (+) = X.MENU_ITEM_OBJECT" _
                & " and X.MENU_ITEM_OBJECT is Not Null"
            Create_TDA(.Tables.Add, "ASTOPSTF", "**", 0, False, "VDD", 1)
            .Tables("ASTOPSTF").Columns("RUNS").DataType = GetType(System.Int64)

        End With

        grdASTUSER2.DataSource = dst.Tables("ASTUSER2")
        grdASTUSER3.DataSource = dst.Tables("ASTUSER3")
        grdASTUSER4.DataSource = dst.Tables("ASTUSER4")

        grdASTOPST1.DataSource = dst.Tables("ASTOPST0")
        grdASTOPSTF.DataSource = dst.Tables("ASTOPSTF")

        With grdASTOPST1.DisplayLayout
            .Bands(0).SortedColumns.Clear()
            .Bands(0).SortedColumns.Add(.Bands(0).Columns("STAT_DATE"), True)
            .Bands(1).SortedColumns.Clear()
            .Bands(1).SortedColumns.Add(.Bands(1).Columns("SESSION_NO"), False)
            .Bands(2).SortedColumns.Clear()
            .Bands(2).SortedColumns.Add(.Bands(2).Columns("INIT_DATE"), False)

        End With
        grdASTOPST1.DisplayLayout.Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.OncePerRowIsland

        rowASTPARMP = ASCDATA1.GetDataRow("Select * from ASTPARMP where AS_PARM_KEY = 'Z'")

        With grdASTUSERX.DisplayLayout.Bands(0)
            .Columns("USER_ID").Header.Fixed = True
            .Columns("USER_NAME").Header.Fixed = True
            .Columns("USER_STATUS").Header.Fixed = True
        End With

        grdASTUSER2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.Default

        dte1.Value = Now.Date.AddDays(-90)
        dte2.Value = Now.Date

        ReParent_Tabs(tabASTUSER1)
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdASTUSERX, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

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

            '    Case "grdATEVNTX"
            '        tlb_btn = DirectCast(tlb_pop.Tools("Refresh"), UltraWinToolbars.ButtonTool)
            '        tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Refresh"
            '    If grd.Name = "grdTATEVNTX" Then
            '        Fill_TATEVNTX()
            '    End If
        End Select
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sql As String = ""

        sql = "Delete from ASTUSER2 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)

        Dim security_codes_added As New List(Of String)
        Dim security_codes_deleted As New List(Of String)


        For Each row As DataRow In dst.Tables("ASTUSER2").Select("", "", DataRowState.Modified)
            Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
            If row.Item("SEL") = "1" Then
                If row.Item("SEL", DataRowVersion.Original) <> "1" Then
                    row.AcceptChanges()
                    row.SetAdded()
                    security_codes_added.Add(SECURITY_CODE)
                End If
            Else
                If row.Item("SEL", DataRowVersion.Original) = "1" Then
                    row.Delete()
                    security_codes_deleted.Add(SECURITY_CODE)
                End If
            End If
        Next

        Dim AS_PARM_SEC_ALERT_EMAIL As String = rowASTPARMP.Item("AS_PARM_SEC_ALERT_EMAIL") & ""

        Dim userStatusDesc As New Dictionary(Of String, String)
        userStatusDesc.Add("I", "Inactive")
        userStatusDesc.Add("A", "Active")
        For Each row As DataRow In dst.Tables("ASTUSER1").Select("", "", DataRowState.Modified)
            If row.Item("USER_STATUS", DataRowVersion.Original) <> row.Item("USER_STATUS") Then
                If AS_PARM_SEC_ALERT_EMAIL <> "" Then
                    Dim ALERT_MESSAGE As String = "User Status changed from " & userStatusDesc(row.Item("USER_STATUS", DataRowVersion.Original)) & " to " & userStatusDesc(row.Item("USER_STATUS"))
                    TAC.TACMAIN1.Send_alert_email(Me, AS_PARM_SEC_ALERT_EMAIL, Absx1.txtFor("USER_ID").Text, "", ALERT_MESSAGE)
                End If
            End If
        Next

        If security_codes_added.Count > 0 Or security_codes_deleted.Count > 0 Then
            If AS_PARM_SEC_ALERT_EMAIL <> "" Then
                Dim ALERT_MESSAGE As String = ""
                If EntryMode = "New" Then
                    ALERT_MESSAGE &= vbCrLf & "New User: " & Absx1.txtFor("USER_ID").Text & " " & Absx1.txtFor("USER_NAME").Text
                End If
                If security_codes_added.Count > 0 Then
                    ALERT_MESSAGE &= vbCrLf & "Security Codes Added: " & Join(security_codes_added.ToArray, ",")
                End If
                If security_codes_deleted.Count > 0 Then
                    ALERT_MESSAGE &= vbCrLf & "Security Codes Deleted: " & Join(security_codes_deleted.ToArray, ",")
                End If
                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
                '    AS_PARM_SEC_ALERT_EMAIL = "wjz@absolution.com"
                'End If

                TAC.TACMAIN1.Send_alert_email(Me, AS_PARM_SEC_ALERT_EMAIL, Absx1.txtFor("USER_ID").Text, "", ALERT_MESSAGE)
            End If
        End If


        WriteAuditTrail("ASTUSER2")

        dst.Tables("ASTUSER2").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER2").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next

        Update_Record_TDA("ASTUSER2")

        sql = "Delete from ASTUSER3 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER3").AcceptChanges()
        If Absx1.optFor("USER_STATUS").Value <> "G" Then
            For Each row As DataRow In dst.Tables("ASTUSER3").Rows
                If row.Item("SEL") = "1" Then
                    row.SetAdded()
                End If
            Next
        End If
        Update_Record_TDA("ASTUSER3")

        sql = "Delete from ASTUSER4 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER4").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER4").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next
        Update_Record_TDA("ASTUSER4")


        If rowASFBASE1.Item("USER_PASSWORD") & "" <> USER_PASSWORD_orig Then
            Dim USER_PASSWORD As String = rowASFBASE1.Item("USER_PASSWORD") & ""
            If rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1" Then
                ' - should not encrypt here since we now encrypt these fields in ASFBASE0
                USER_PASSWORD = ASCMAIN1.EncryptAES(USER_PASSWORD)
                rowASFBASE1.Item("USER_PASSWORD") = USER_PASSWORD
            End If
            rowASFBASE1.Item("USER_PASSWORD_LAST_DATE") = DATETIME_STAMP
        End If

        If (EntryMode = "New") And Absx1.chkFor("USER_ADMIN").Checked Then

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC)" & vbCrLf _
                & " values ('ASTUSER1', :PARM1, SYSDATE, :PARM2,'ADDADMIN','New Admin created: " & Absx1.txtFor("USER_ID").Text & "')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {Absx1.txtFor("USER_ID").Text, ASCMAIN1.USER_ID})

            If AS_PARM_SEC_ALERT_EMAIL <> "" Then
                Dim ALERT_MESSAGE As String = "New Admin created: " & Absx1.txtFor("USER_ID").Text & " by " & ASCMAIN1.USER_ID
                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
                '    AS_PARM_SEC_ALERT_EMAIL = "wjz@absolution.com"
                'End If
                TAC.TACMAIN1.Send_alert_email(Me, AS_PARM_SEC_ALERT_EMAIL, Absx1.txtFor("USER_ID").Text, "", ALERT_MESSAGE)
            End If
        End If

        If chkResetPassword.Checked Then

            ' Send email to user
            Dim em As New TAC.ASCNOTEE(ASCMAIN1.Folders, "ASTUSER1", Nothing)
            em.CreateComponents()
            em.SetEmailSubject(“Your ABSolution password has been reset to a temporary password”)
            Dim USER_EMAIL As String = Absx1.txtFor("USER_EMAIL").Text

            'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            '    USER_EMAIL = "wjz@absolution.com"
            'End If

            em.SetEmailTo(USER_EMAIL)

            'Dim objASCNOTE1 As New TAC.ASCNOTE1("CONF_ADS", Nothing)
            'objASCNOTE1.Note = note
            'objASCNOTE1.CreateComponents()
            'objASCNOTE1.EmailDocument()

            Dim GUID As String = System.Guid.NewGuid.ToString()

            Dim NEW_PASSWORD = Mid(GUID, 1, 10)
            Dim USER_NAME As String = Absx1.txtFor("USER_NAME").Text & ""

            Dim AS_PARM_TEMP_PWD_HRS As Integer = Val(rowASTPARMP.Item("AS_PARM_TEMP_PWD_HRS") & "")

            Dim emailParms As New Dictionary(Of String, String)
            emailParms.Add("NEW_PASSWORD", NEW_PASSWORD)
            emailParms.Add("USER_NAME", USER_NAME)
            emailParms.Add("AS_PARM_TEMP_PWD_HRS", AS_PARM_TEMP_PWD_HRS)

            Dim strHtml As String = Build_Html_Email(emailParms, AS_PARM_TEMP_PWD_HRS)

            em.SetDocumentBody(strHtml)

            ' em.SaveEmail = True
            em.EmailDocument()
            ' Dim SEND_NO As String = em.EmailSendNo

            Dim USER_PASSWORD As String = NEW_PASSWORD
            USER_PASSWORD = NEW_PASSWORD ' ASCMAIN1.EncryptAES(NEW_PASSWORD) - should not encrypt here since we now encrypt these fields in ASFBASE0

            rowASFBASE1.Item("USER_PASSWORD") = USER_PASSWORD
            rowASFBASE1.Item("USER_PASSWORD_TEMP") = "1"
            rowASFBASE1.Item("USER_MUST_CHG_PWD") = "1"
            rowASFBASE1.Item("USER_PASSWORD_LAST_DATE") = DATETIME_STAMP

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC)" & vbCrLf _
                & " values ('ASTUSER1', :PARM1, SYSDATE, :PARM2,'PWDRESET','Password was Reset')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {Absx1.txtFor("USER_ID").Text, ASCMAIN1.USER_ID})

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC)" & vbCrLf _
                & $" values ('ASTUSER1', :PARM1, SYSDATE, :PARM2,'TEMPPWD','Temp Password emailed to {USER_EMAIL}')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {Absx1.txtFor("USER_ID").Text, ASCMAIN1.USER_ID})

            MsgBox("Password Reset", MsgBoxStyle.OkOnly, "Verification")
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        chkResetPassword.Checked = False

        EnforceConstraints(False)
        Get_ASTOPST1()

        Fill_Records("ASTUSER2", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER2, "SECURITY_CODE")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from ASTSECM1 order by SECURITY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER2 As DataRow = dst.Tables("ASTUSER2").NewRow
                rowASTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER2.Item("SEL") = "0"
                rowASTUSER2.Item("SECURITY_CODE") = row.Item("SECURITY_CODE")
                rowASTUSER2.Item("SECURITY_DESC") = row.Item("SECURITY_DESC")
                dst.Tables("ASTUSER2").Rows.Add(rowASTUSER2)
            Next
            dst.Tables("ASTUSER2").AcceptChanges()
        End If

        Fill_Records("ASTUSER3", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER3, "USER_GROUP_ID")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from ASTUSER1 where USER_STATUS = 'G' order by USER_ID"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER3 As DataRow = dst.Tables("ASTUSER3").NewRow
                rowASTUSER3.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER3.Item("SEL") = "0"
                rowASTUSER3.Item("USER_GROUP_ID") = row.Item("USER_ID")
                rowASTUSER3.Item("USER_GROUP_NAME") = row.Item("USER_NAME")
                dst.Tables("ASTUSER3").Rows.Add(rowASTUSER3)
            Next
        End If

        Fill_Records("ASTUSER4", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER4, "COMPANY_CODE")
        If EntryMode = "New" Then
            ASCMAIN1.sql = "Select * from GLTCOMP1 order by COMPANY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER4 As DataRow = dst.Tables("ASTUSER4").NewRow
                rowASTUSER4.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER4.Item("SEL") = "0"
                rowASTUSER4.Item("COMPANY_CODE") = row.Item("COMPANY_CODE")
                rowASTUSER4.Item("COMPANY_NAME") = row.Item("COMPANY_NAME")
                dst.Tables("ASTUSER4").Rows.Add(rowASTUSER4)
            Next
        End If

        If dst.Tables("ASTUSER4").Select("SEL='1'").Length = 0 Then
            chkAllCompanies.Checked = True
        Else
            chkAllCompanies.Checked = False
        End If

        USER_PASSWORD_orig = Absx1.txtFor("USER_PASSWORD").Text

        txtUSER_PASSWORD.Visible = Not (rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1")
        chkResetPassword.Visible = (rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1")

        EnforceConstraints(True)
    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            'dst.EnforceConstraints = False
            'dst.Tables("ASTOPST0").Rows.Clear()
            'dst.Tables("ASTOPST1").Rows.Clear()
            'dst.Tables("ASTOPSTX").Rows.Clear()
            'dst.EnforceConstraints = True
            Fill_ASTUSERX()

        End If
    End Sub

    Sub Fill_ASTUSERX()

        Fill_Records("ASTUSERX")
        Sort_grdColumns(grdASTUSERX, "USER_ID")

        ASCMAIN1.sql = "Select * from ASTUSER2"
        For Each row As DataRow In ASCDATA1.GetDataTable("", "ASTUSER2").Rows
            Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
            Dim USER_ID As String = row.Item("USER_ID")
            Dim rowASTUSERX As DataRow = dst.Tables("ASTUSERX").Rows.Find(USER_ID)
            If SECURITY_CODEs.Contains(SECURITY_CODE) Then
                If Not rowASTUSERX Is Nothing Then
                    rowASTUSERX.Item(SECURITY_CODE) = True
                End If
            End If
        Next

    End Sub


    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTUSERX.Visible = Not tf
        UltraTabControl1.Visible = tf
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                Dim USER_ID As String = Absx1.txtFor("USER_ID").Text

                If USER_ID.Length > 0 Then
                    If USER_ID <> USER_ID.ToLower Then
                        EMsg &= vbCr & "User ID should use lowercase letters only"
                    Else
                        For i As Int16 = 1 To USER_ID.Length
                            Dim z As String = USER_ID.Substring(i - 1, 1)
                            If z < "a" Or z > "z" Then
                                If InStr("0123456789", z) = 0 Then
                                    EMsg &= vbCr & "User ID should use lowercase letters and numbers only"
                                End If
                            End If
                        Next
                    End If
                End If

            Case "Edit"
                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    If Absx1.txtFor("USER_ID").Text = ASCMAIN1.USER_ID Then
                        EMsg &= EMsg & "You may NOT maintain your own record"
                    End If
                End If

            Case "Update"

                If Absx1.optFor("USER_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "User Status must be selected"
                Else
                    If (EntryMode = "New") Then
                        If Absx1.optFor("USER_STATUS").Value & "" <> "A" Then
                            EMsg &= vbCr & "User Status should be set to Active for a new User"
                        End If
                    End If
                End If


                If Absx1.txtFor("USER_NAME").Text = "" Then
                    EMsg &= vbCr & "User Name is Mandatory"
                End If

                If chkResetPassword.Checked Then
                    If Absx1.txtFor("USER_EMAIL").Text = "" Then
                        EMsg &= vbCr & "User must have an email address in order to Reset Password"
                    End If

                End If

                Dim SY_assigned As Boolean = False

                Dim SECURITY_CODESs_not_permitted As New List(Of String)
                For Each row As DataRow In dst.Tables("ASTUSER2").Select("SEL = '1'", "")
                    Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
                    If SECURITY_CODE = "SY" Then
                        SY_assigned = True
                    End If
                    If Absx1.chkFor("USER_ADMIN").Checked Then
                        If SECURITY_CODE <> "SY" Then
                            SECURITY_CODESs_not_permitted.Add(SECURITY_CODE)
                        End If
                    Else
                        If SECURITY_CODE = "SY" Then
                            SECURITY_CODESs_not_permitted.Add(SECURITY_CODE)
                        End If
                    End If
                Next
                If SECURITY_CODESs_not_permitted.Count <> 0 Then
                    EMsg &= vbCr & vbCrLf & "Admin users are constrained to Security Code SY." _
                        & vbCrLf & "Non-Admin users are not permitted to have Security Code SY" _
                        & vbCrLf _
                        & vbCrLf & " - The following Security Codes are not permitted for this user: " _
                        & Join(SECURITY_CODESs_not_permitted.ToArray, ",")
                End If
                If Absx1.chkFor("USER_ADMIN").Checked And Not SY_assigned Then
                    EMsg &= vbCr & "Admin User must have Security Code SY assinged"
                End If

                'Dim password_error_checks As String = _
                'ASCMAIN1.Validate_User_Password( _
                'False, _
                'Absx1.txtFor("USER_ID").Text, _
                'Absx1.txtFor("USER_PASSWORD").Text, _
                'rowASTPARMP)

                If Not chkAllCompanies.Checked Then
                    If dst.Tables("ASTUSER4").Select("SEL='1'").Length = 0 Then
                        EMsg &= vbCr & "You must select at least 1 company if not granting All Companies"
                    End If
                End If

                'If password_error_checks <> "" Then
                '    EMsg &= vbCr & "Password Errors:" & vbCr & vbTab & Replace(password_error_checks, vbCr, vbCr & vbTab)
                'End If

        End Select

    End Sub

    Public Overrides Sub Mode_Settings(tf As Boolean, Optional MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        Set_Read_Only_for_ctl(Absx1.chkFor("USER_ADMIN"), Not (EntryMode = "New"))


        If (EntryMode = "New") Then
            chkResetPassword.Checked = True
            Set_Read_Only_for_ctl(chkResetPassword, True)
        Else
            Set_Read_Only_for_ctl(chkResetPassword, Not (EntryMode = "Edit"))
        End If
    End Sub

#End Region

    Private Sub grdASTUSERX_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdASTUSERX.DoubleClickCell
        If grdASTUSERX.ActiveCell Is Nothing Then
            Exit Sub
        End If
        Absx1.txtFor("USER_ID").Text = grdASTUSERX.ActiveCell.Row.Cells("USER_ID").Text
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Call Click_Command("Edit")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        EnforceConstraints(False)
        Get_ASTOPST1()
        EnforceConstraints(True)
    End Sub

    Sub Get_ASTOPST1()

        Dim caption As String = String.Format _
        ("Operator Statistics for {0} for the {1} days from {2} to {3}",
         Absx1.txtFor("USER_ID").Text,
         CStr(1 + DateDiff("d", dte1.DateTime, dte2.DateTime)),
         Format(dte1.Value, "MM/dd/yyyy"),
         Format(dte2.Value, "MM/dd/yyyy"))

        grdASTOPST1.Text = caption

        caption = String.Format _
            ("Function Statistics for {0} for the {1} days from {2} to {3}",
             Absx1.txtFor("USER_ID").Text,
             CStr(1 + DateDiff("d", dte1.DateTime, dte2.DateTime)),
             Format(dte1.Value, "MM/dd/yyyy"),
             Format(dte2.Value, "MM/dd/yyyy"))

        grdASTOPSTF.Text = caption

        Fill_Records("ASTOPST0", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPST1", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPSTX", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})
        Fill_Records("ASTOPSTF", New Object() {Absx1.txtFor("USER_ID").Text, dte1.Value, dte2.Value})

        Sort_grdColumns(grdASTOPST1, "STAT_DATE".ToLower)
        Sort_grdColumns(grdASTOPSTF, "MENU_ITEM_OBJECT")

    End Sub

    Private Sub chkAllCompanies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        grdASTUSER4.Enabled = Not chkAllCompanies.Checked
        For Each row As DataRow In dst.Tables("ASTUSER4").Rows
            row.Item("SEL") = "0"
        Next
    End Sub

    Function Build_Html_Email(eParms As Dictionary(Of String, String), AS_PARM_TEMP_PWD_HRS As Integer) As String
        Dim eBody As StringBuilder = New StringBuilder()

        With eBody
            .Append("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
            .Append("<html xmlns=""http://www.w3.org/1999/xhtml"">")
            .Append("<head>")
            .Append("  <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
            .Append("  <!--[if !mso]><!-->")
            .Append("  <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />")
            .Append("  <!--<![endif]-->")
            .Append("  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />")
            .Append("  <meta name=""color-scheme"" content=""light"">")
            .Append("  <meta name=""supported-color-schemes"" content=""light"">")
            .Append("  <title></title>")
            .Append("  <style type=""text/css"">")
            .Append(".ReadMsgBody { width: 100%; background-color: #ffffff; }")
            .Append(".ExternalClass { width: 100%; background-color: #ffffff; }")
            .Append(".ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div { line-height: 100%; }")
            .Append("html { width: 100%; }")
            .Append("body { -webkit-text-size-adjust: none; -ms-text-size-adjust: none; margin: 0; padding: 0; }")
            .Append("table { border-spacing: 0; table-layout: auto; margin: 0 auto; }")
            .Append(".yshortcuts a { border-bottom: none !important; }")
            .Append("img:hover { opacity: 0.9 !important; }")
            .Append("a { color: #3cb2d0; text-decoration: none; }")
            .Append(".textbutton a { font-family: 'open sans', arial, sans-serif !important; }")
            .Append(".btn-link a { color: #FFFFFF !important; }")
            .Append("@media only screen and (max-width: 479px) {")
            .Append("body { width: auto !important; font-family: 'Open Sans', Arial, Sans-serif !important;}")
            .Append(".table-inner{ width: 90% !important; text-align: center !important;}")
            .Append(".table-full { width: 100%!important; max-width: 100%!important; text-align: center !important;}")
            .Append("/*gmail*/")
            .Append("u + .body .full { width:100% !important; width:100vw !important;}")
            .Append("}")
            .Append("</style>")
            .Append("</head>")
            .Append("<body class=""body"">")
            .Append("  <table class=""full"" width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
            .Append("    <tr>")
            .Append("      <td  bgcolor=""#494c50"" valign=""top"" style=""background-size: cover; background-position: center;"">")
            .Append("        <table class=""table-inner"" align=""center"" width=""600"" style=""max-width: 600px;"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            .Append("          <tr>")
            .Append("            <td height=""40""></td>")
            .Append("          </tr>")
            .Append("          <tr>")
            .Append("            <td bgcolor=""#FFFFFF"" style=""border-top-left-radius: 4px;border-top-right-radius: 4px;"" align=""center"">")
            .Append("              <table width=""90%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
            .Append("                <tr>")
            .Append("                  <td height=""50""></td>")
            .Append("                </tr>")
            .Append("                <!-- logo -->")
            .Append("                <tr>")
            .Append("                  <td align=""center"" style=""line-height: 0px;""><img style=""display:block; line-height:0px; font-size:0px; border:0px;"" src=""https://portal.interparfums.com/assets/admin/layout/img/logo400.png"" alt=""logo"" /></td>")
            .Append("                </tr>")
            .Append("                <!-- end logo -->")
            .Append("                <tr>")
            .Append("                  <td height=""15""></td>")
            .Append("                </tr>")
            .Append("                <tr>")
            .Append("                  <td height=""40""></td>")
            .Append("                </tr>")
            .Append("              </table>")
            .Append("            </td>")
            .Append("          </tr>")
            .Append("          <tr>")
            .Append("            <td align=""center"" bgcolor=""#f3f3f3"">")
            .Append("              <table width=""90%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
            .Append("                <tr>")
            .Append("                  <td height=""50"">##USER_NAME##</td>")
            .Append("                </tr>")
            .Append("                <!-- title -->")
            .Append("                <tr>")
            .Append("                  <td align=""center"" style=""font-family: 'Open Sans', Arial, sans-serif; font-size:30px; color:#3b3b3b; font-weight: bold; letter-spacing:4px;"">Your ABSolution password has been reset</td>")
            .Append("                </tr>")
            .Append("                <!-- end title -->")
            .Append("                <tr>")
            .Append("                  <td align=""center"">")
            .Append("                    <table width=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            .Append("                      <tr>")
            .Append("                        <td height=""15"" style=""border-bottom:2px solid ##333333;""></td>")
            .Append("                      </tr>")
            .Append("                    </table>")
            .Append("                  </td>")
            .Append("                </tr>")
            .Append("                <tr>")
            .Append("                  <td height=""20""></td>")
            .Append("                </tr>")
            .Append("                <!-- content -->")
            .Append("                <tr>")
            .Append("                  <td align=""center"" style=""font-family: 'Open Sans', Arial, sans-serif; font-size:14px; color:#7f8c8d; line-height:29px;""> Please use the new temporary password below to log into ABSolution. If you did not request a password reset and believe you have received this email in error, please contact your system administrator. </td>")
            .Append("                </tr>")
            .Append("                <!-- end content -->")
            .Append("                <tr>")
            .Append("                  <td height=""50""></td>")
            .Append("                </tr>")
            If AS_PARM_TEMP_PWD_HRS <> 0 Then
                .Append("                <tr>")
                .Append("                  <td align=""center"" style=""font-family: 'Open Sans', Arial, sans-serif; font-size:14px; color:#7f8c8d; line-height:29px;""> The password reset link is valid for ##AS_PARM_TEMP_PWD_HRS## hour" & IIf(AS_PARM_TEMP_PWD_HRS = 1, "", "s.") & " </td>")
                .Append("                </tr>")
            End If
            .Append("              </table>")
            .Append("            </td>")
            .Append("          </tr>")
            .Append("          <tr>")
            .Append("            <td bgcolor=""#FFFFFF"" style=""border-bottom-left-radius: 4px;border-bottom-right-radius: 4px;"" align=""center"">")
            .Append("              <table width=""90%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
            .Append("                <tr>")
            .Append("                  <td height=""40""></td>")
            .Append("                </tr>")
            .Append("                <!-- button -->")
            .Append("                <tr>")
            .Append("                  <td align=""center"">")
            .Append("                    <table class=""textbutton"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            .Append("                      <tr>")
            .Append("                        <td class=""btn-link"" bgcolor=""#333333"" height=""55"" align=""center"" style=""font-family: 'Open Sans', Arial, sans-serif; font-size:16px; color:#FFFFFF;font-weight: bold;padding-left: 25px;padding-right: 25px;border-radius:4px;""><a href=""#"">##NEW_PASSWORD##</a></td>")
            .Append("                      </tr>")
            .Append("                    </table>")
            .Append("                  </td>")
            .Append("                </tr>")
            .Append("                <!-- end button -->")
            .Append("                <tr>")
            .Append("                  <td height=""25""></td>")
            .Append("                </tr>")
            .Append("                <tr>")
            .Append("                  <td height=""30""></td>")
            .Append("                </tr>")
            .Append("              </table>")
            .Append("            </td>")
            .Append("          </tr>")
            .Append("          <tr>")
            .Append("            <td height=""25""></td>")
            .Append("          </tr>")
            .Append("          <!-- copyright -->")
            .Append("          <tr>")
            .Append("            <td align=""center"" style=""font-family: 'Open Sans', Arial, sans-serif; font-size:13px; color:#ffffff;""> © 2020 Interparfums Luxury Brands </td>")
            .Append("          </tr>")
            .Append("          <!-- end copyright -->")
            .Append("          <tr>")
            .Append("            <td height=""25""></td>")
            .Append("          </tr>")
            .Append(" ")
            .Append("          <tr>")
            .Append("            <td height=""45""></td>")
            .Append("          </tr>")
            .Append("        </table>")
            .Append("      </td>")
            .Append("    </tr>")
            .Append("  </table>")
            .Append("</body>")
            .Append("</html>")

        End With

        Dim strHtml As String = eBody.ToString

        Dim eParm As KeyValuePair(Of String, String)
        For Each eParm In eParms
            Dim parmName As String = "##" & eParm.Key & "##"
            strHtml = Replace(strHtml, parmName, eParm.Value)
        Next
        Return strHtml
    End Function

    Private Sub btnGenerateExtract_Click(sender As Object, e As EventArgs)
        Dim AUDIT_TABLES As New Dictionary(Of String, String)
        AUDIT_TABLES.Add("ASTOPST1", "Operator Statistics")
        AUDIT_TABLES.Add("ASTPARM1", "System Parameters")
        AUDIT_TABLES.Add("ASTPARMP", "Password Parameters")
        AUDIT_TABLES.Add("ASTSECM1", "Security Codes")
        AUDIT_TABLES.Add("ASTUSER2", "User Security Codes")
        AUDIT_TABLES.Add("TATUSER1", "User Profiles")
        AUDIT_TABLES.Add("ASTUSER1", "Users")
        AUDIT_TABLES.Add("ASTMENU1", "Menu")
        For Each kvp As KeyValuePair(Of String, String) In AUDIT_TABLES

        Next

    End Sub
End Class