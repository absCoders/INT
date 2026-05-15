Imports System.Text.RegularExpressions
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Security.Cryptography
Imports System.Windows.Forms

Public Class ASTUSERW

    Dim rowASTPARMP As DataRow
    Dim USER_PASSWORD_orig As String

    Private randomBytes() As Byte
    Private randomInt32Value As Integer
    Private possibleChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()"
    Private len As Int32 = 8
    Private GetRandomInt32Value As New RandomInt32Value
    Private GetPasswordGenProfiler As New PasswordGenProfiler

    Private m_emailRegEx As Regex

    Private ReadOnly Property EmailRegEx() As Regex
        Get
            If Me.m_emailRegEx Is Nothing Then
                Me.m_emailRegEx = New Regex("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b", RegexOptions.IgnoreCase)
            End If

            Return Me.m_emailRegEx
        End Get
    End Property

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
        With dst
            ASCMAIN1.sql = "Select USER_EMAIL, USER_NAME, USER_STATUS from ASTUSERW"
            Create_TDA(.Tables.Add, "ASTUSERX", "**", 0, False, "", 1)

            grdASTUSERX.DataSource = dst.Tables("ASTUSERX")

            Call Fill_ASTUSERX()

            ASCMAIN1.sql = "SELECT CLIENT_IP, 'Authentication' EVENT_TYPE" _
            & ", EVENT_DATE, EVENT_DESC FROM ASTWEBL1 " _
            & " WHERE ASTWEBL1.USER_EMAIL = :PARM1 " _
            & " AND ASTWEBL1.EVENT_DATE >= :PARM2 AND ASTWEBL1.EVENT_DATE -1 <= :PARM3" _
            & " AND EVENT_TYPE = 'A'"
            Create_TDA(.Tables.Add, "ASTWEBLA", "**", 0, False, "VDD")

            ASCMAIN1.sql = "SELECT CLIENT_IP, EVENT_TYPE" _
            & ", EVENT_DATE, EVENT_DESC FROM ASTWEBL1 " _
            & " WHERE ASTWEBL1.USER_EMAIL = :PARM1 " _
            & " AND ASTWEBL1.EVENT_DATE >= :PARM2 AND ASTWEBL1.EVENT_DATE -1 <= :PARM3" _
            & " AND EVENT_TYPE <> 'A'"
            Create_TDA(.Tables.Add, "ASTWEBL1", "**", 0, False, "VDD")

            'If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            '    ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE " _
            '    & " FROM ASTOPST1 where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 GROUP BY CONVERT(DATE,INIT_DATE)"
            'Else
            '    ASCMAIN1.sql = "SELECT DISTINCT TRUNC(INIT_DATE) STAT_DATE" _
            '    & " FROM ASTOPST1 where USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE)" '  and SELECTION_NO = 0 
            'End If
            'Create_TDA(.Tables.Add, "ASTOPST0", "**", 0, False, "VDD", 1)

            'If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            '    ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
            '    & " , ASTOPST1.INIT_DATE, ASTOPST1.LAST_DATE " _
            '    & " FROM ASTOPST1 " _
            '    & " where ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3" '  and SELECTION_NO = 0
            'Else
            '    ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.SESSION_NO " _
            '    & " , MIN (ASTOPST1.INIT_DATE) INIT_DATE, MAX (ASTOPST1.LAST_DATE) LAST_DATE " _
            '    & " FROM ASTOPST1 " _
            '    & " where ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 GROUP BY TRUNC(INIT_DATE), ASTOPST1.SESSION_NO" '  and SELECTION_NO = 0
            'End If
            'Create_TDA(.Tables.Add, "ASTOPST1", "**", 0, False, "VDD", 0)

            'If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            '    ASCMAIN1.sql = "SELECT CONVERT(DATE,INIT_DATE) STAT_DATE, ASTOPST1.* " _
            '    & " , ASTMENU1.MENU_ITEM_DESC " _
            '    & " FROM ASTOPST1,ASTMENU1 " _
            '    & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
            '    & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
            '    & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
            '    & "   and ASTOPST1.USER_ID = @PARM1 AND ASTOPST1.INIT_DATE >= @PARM2 AND ASTOPST1.INIT_DATE -1 <= @PARM3 and SELECTION_NO <> 0"
            'Else
            '    ASCMAIN1.sql = "SELECT TRUNC(INIT_DATE) STAT_DATE, ASTOPST1.* " _
            '    & " , ASTMENU1.MENU_ITEM_DESC " _
            '    & " FROM ASTOPST1,ASTMENU1 " _
            '    & " where ASTOPST1.MENU_ID = ASTMENU1.MENU_ID (+) " _
            '    & "   and ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE (+) " _
            '    & "   and ASTOPST1.MENU_ITEM_OBJECT = ASTMENU1.MENU_ITEM_OBJECT (+) " _
            '    & "   and ASTOPST1.USER_ID = :PARM1 AND ASTOPST1.INIT_DATE >= :PARM2 AND ASTOPST1.INIT_DATE -1 <= :PARM3 and SELECTION_NO <> 0"
            'End If
            'Create_TDA(.Tables.Add, "ASTOPST2", "**", 0, False, "VDD", 0)

            'Create_Relation("ASTOPST0", "ASTOPST1", "STAT_DATE")
            'Create_Relation("ASTOPST1", "ASTOPST2", "STAT_DATE,SESSION_NO")

            '.Tables("ASTOPST0").Columns.Add("SESSIONS", GetType(System.Int64), "COUNT(CHILD.SESSION_NO)")
            '.Tables("ASTOPST1").Columns.Add("SELECTIONS", GetType(System.Int64), "COUNT(CHILD.SELECTION_NO)")
            '.Tables("ASTOPST0").Columns.Add("SELECTIONS", GetType(System.Int64), "SUM(CHILD.SELECTIONS)")
        End With

        grdASTWEBLA.DataSource = dst.Tables("ASTWEBLA")
        grdASTWEBL1.DataSource = dst.Tables("ASTWEBL1")

        rowASTPARMP = ASCDATA1.GetDataRow("Select * from ASTPARMP where AS_PARM_KEY = 'Z'")

        With grdASTUSERX.DisplayLayout.Bands(0)
            .Columns("USER_EMAIL").Header.Fixed = True
            .Columns("USER_NAME").Header.Fixed = True
            .Columns("USER_STATUS").Header.Fixed = True
        End With

        dte1.Value = Now.Date.AddDays(-90)
        dte2.Value = Now.Date

        dte3.Value = Now.Date.AddDays(-90)
        dte4.Value = Now.Date

        ReParent_Tabs(tabASTUSER1)

    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If SELECTION_NO = 0 Then Exit Sub
    End Sub


#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTUSERX, "SS", "Show Filter", "Show GroupBox")
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
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sql As String = ""

        If rowASFBASE1.Item("USER_PASSWORD") & "" <> USER_PASSWORD_orig Then
            Dim USER_PASSWORD As String = rowASFBASE1.Item("USER_PASSWORD") & ""
            If rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1" Then
                Dim MD5 As New ASCSCMD5
                USER_PASSWORD = MD5.DigestStrToHexStr(USER_PASSWORD)
                rowASFBASE1.Item("USER_PASSWORD") = USER_PASSWORD
                MD5 = Nothing
            End If
            rowASFBASE1.Item("USER_PASSWORD_LAST_DATE") = DATETIME_STAMP
        End If

        If EntryMode = "New" Then
            Dim WEB_USER_TOKEN As String = ASCMAIN1.Next_Control_No("ASTUSERW.WEB_USER_TOKEN")
            rowASFBASE1.Item("WEB_USER_TOKEN") = WEB_USER_TOKEN
        End If

    End Sub

    Overrides Sub Show_Record_Special()

        EnforceConstraints(False)
        Get_ASTWEBLA()
        Get_ASTWEBL1()
        USER_PASSWORD_orig = Absx1.txtFor("USER_PASSWORD").Text
        EnforceConstraints(True)

    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            'dst.EnforceConstraints = False
            'dst.Tables("ASTOPST0").Rows.Clear()
            'dst.Tables("ASTOPST1").Rows.Clear()
            'dst.Tables("ASTOPST2").Rows.Clear()
            'dst.EnforceConstraints = True
            Call Fill_ASTUSERX()
        End If
    End Sub

    Sub Fill_ASTUSERX()

        Fill_Records("ASTUSERX")
        Sort_grdColumns(grdASTUSERX, "USER_EMAIL")

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTUSERX.Visible = Not tf
        UltraTabControl1.Visible = tf
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                Dim USER_EMAIL As String = Absx1.txtFor("USER_EMAIL").Text

                If USER_EMAIL.Length > 0 Then
                    If USER_EMAIL <> USER_EMAIL.ToLower Then
                        EMsg &= vbCr & "User Email should use lowercase letters only"
                    Else
                        If Not EmailRegEx.IsMatch(USER_EMAIL) Then
                            EMsg &= vbCr & "User Email is not Valid"
                        End If

                    End If
                End If

            Case "Update"

                Dim password_error_checks As String = _
                ASCMAIN1.Validate_User_Password( _
                False, _
                Absx1.txtFor("USER_EMAIL").Text, _
                Absx1.txtFor("USER_PASSWORD").Text, _
                rowASTPARMP)

                If password_error_checks <> "" Then
                    EMsg &= vbCr & "Password Errors:" & vbCr & vbTab & Replace(password_error_checks, vbCr, vbCr & vbTab)
                End If

        End Select

    End Sub
#End Region

    Private Sub grdASTUSERX_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdASTUSERX.DoubleClickCell
        If grdASTUSERX.ActiveCell Is Nothing Then
            Exit Sub
        End If
        Absx1.txtFor("USER_EMAIL").Text = grdASTUSERX.ActiveCell.Row.Cells("USER_EMAIL").Text
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Call Click_Command("Edit")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdFetchAuthenticationActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchAuthenticationActivity.Click
        EnforceConstraints(False)
        Get_ASTWEBLA()
        EnforceConstraints(True)
    End Sub

    Sub Get_ASTWEBLA()

        Dim caption As String = String.Format _
        ("User Authentication Activity for the {0} days from {1} to {2}", _
         CStr(1 + DateDiff("d", dte1.DateTime, dte2.DateTime)), _
         Format(dte1.Value, "MM/dd/yyyy"), _
         Format(dte2.Value, "MM/dd/yyyy"))

        grdASTWEBLA.Text = caption

        Fill_Records("ASTWEBLA", New Object() {Absx1.txtFor("USER_EMAIL").Text, dte1.Value, dte2.Value})

        Sort_grdColumns(grdASTWEBLA, "EVENT_DATE".ToLower)

    End Sub

    Private Sub cmdFetchOtherActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchOtherActivity.Click
        EnforceConstraints(False)
        Get_ASTWEBL1()
        EnforceConstraints(True)
    End Sub

    Sub Get_ASTWEBL1()

        Dim caption As String = String.Format _
        ("Web Portal Activity for the {0} days from {1} to {2}", _
         CStr(1 + DateDiff("d", dte3.DateTime, dte4.DateTime)), _
         Format(dte3.Value, "MM/dd/yyyy"), _
         Format(dte4.Value, "MM/dd/yyyy"))

        grdASTWEBL1.Text = caption

        Fill_Records("ASTWEBL1", New Object() {Absx1.txtFor("USER_EMAIL").Text, dte3.Value, dte4.Value})

        Sort_grdColumns(grdASTWEBL1, "EVENT_DATE".ToLower)

    End Sub

    Private Sub cmdGeneratePassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGeneratePassword.Click
        Try
            Dim cpossibleChars() As Char
            cpossibleChars = possibleChars.ToCharArray()
            If cpossibleChars.Length < 1 Then
                MessageBox.Show("You must enter one or more possible characters.")
                Return
            End If
            If len < 4 Then
                MessageBox.Show(String.Format("Please choose a password length. That length must be a value between {0} and {1}. Note: values above 1,000 might take a LONG TIME to process on some computers.", 4, Int32.MaxValue))
                Return
            End If

            Dim builder As New StringBuilder()

            For i As Integer = 0 To len - 1
                Dim randInt32 As Integer = GetRandomInt32Value.GetRandomInt()
                Dim r As New Random(randInt32)

                Dim nextInt As Integer = r.[Next](cpossibleChars.Length)
                Dim c As Char = cpossibleChars(nextInt)
                builder.Append(c)
            Next
            Absx1.txtFor("USER_PASSWORD").Text = builder.ToString()
            Absx1.dteFor("USER_PASSWORD_LAST_DATE").Value = DATETIME_STAMP
        Catch ex As Exception
            MessageBox.Show(String.Format("An error has occurred while trying to generate random password! Technical description: {0}", ex.Message.ToString()))
        End Try
    End Sub

    Private Sub optUserType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUserType.ValueChanged
        Absx1.txtFor("USER_ID").Visible = (optUserType.Value = "A")
        Absx1.txtFor("VEND_CODE").Visible = (optUserType.Value <> "A")
    End Sub

End Class

Public Class PasswordGenProfiler
    Public Shared Function GetFrequencyDistributionOfChars(ByVal allowableChars As String, ByVal generatedPass As String) As Dictionary(Of Char, Integer)
        Dim distrib As New Dictionary(Of Char, Integer)()
        ' initialize all values to 0
        For Each c As Char In allowableChars
            ' If character is listed more than once, don't re-add it to our list.
            If Not distrib.ContainsKey(c) Then
                distrib.Add(c, 0)
            End If
        Next
        Dim val As Integer = 0
        For Each passChar As Char In generatedPass
            If distrib.TryGetValue(passChar, val) Then
                distrib(passChar) = System.Threading.Interlocked.Increment(val)
            End If
        Next

        Return distrib
    End Function
End Class

Public Class RandomInt32Value
    Public Function GetRandomInt() As Integer
        Dim randomBytes As Byte() = New Byte(3) {}
        Dim rng As New RNGCryptoServiceProvider()
        rng.GetBytes(randomBytes)
        Dim randomInt As Integer = BitConverter.ToInt32(randomBytes, 0)
        Return randomInt
    End Function
End Class
