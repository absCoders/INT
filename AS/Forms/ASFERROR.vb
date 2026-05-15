Public Class ASFERROR
    Dim tblASTERROR As New DataTable

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Create_Lookup("ASTUSER1")
        dst.Tables.Add(tblASTERROR)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Errors"
                If Absx1.txtFor("USER_ID").Text <> "" Then
                    Validate_Code("USER_ID")
                End If

                If dteERROR_DATE.Value Is Nothing Then
                    EMsg &= vbCr & "No Date Specified"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Errors"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Errors").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Errors from Log")
        Me.Cursor = Cursors.WaitCursor

        Dim sql As String = ""
        If Absx1.txtFor("USER_ID").Text <> "" Then
            sql = " and ASTERROR.INIT_OPER = '" & Absx1.txtFor("USER_ID").Text & "'"
        End If
        If Absx1.txtFor("FORM_NAME").Text <> "" Then
            sql = " and ASTOPST1.MENU_ITEM_OBJECT = '" & Absx1.txtFor("FORM_NAME").Text & "'"
        End If
        If optLoadOption.Value = "0" Then
            sql = sql & " and ASTERROR.INIT_DATE >= '" & Format(dteERROR_DATE.Value, "dd-MMM-yyyy") & "'"
            sql = sql & " and ASTERROR.INIT_DATE < '" & Format(DateValue(dteERROR_DATE.Value.ToString).AddDays(1), "dd-MMM-yyyy") & "'"
        Else
            sql = sql & " and ASTERROR.INIT_DATE >= '" & Format(DateValue(dteERROR_DATE.Value.ToString).AddDays(-100), "dd-MMM-yyyy") & "'"
            sql = sql & " and ASTERROR.INIT_DATE < '" & Format(DateValue(dteERROR_DATE.Value.ToString).AddDays(1), "dd-MMM-yyyy") & "'"
        End If
        ASCMAIN1.sql = "Select ASTERROR.*,ASTOPST1.MENU_ITEM_OBJECT FORM_NAME" _
        & " from ASTERROR,ASTOPST1 " _
        & " where ASTOPST1.SESSION_NO (+) = ASTERROR.SESSION_NO" _
        & "   and ASTOPST1.SELECTION_NO (+) = ASTERROR.SELECTION_NO" _
        & sql

        If optLoadOption.Value = "1" Then
            ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & " order by ASTERROR.INIT_DATE DESC) where ROWNUM < 101"
        End If

        tblASTERROR = ASCDATA1.GetDataTable
        grdASTERROR.DataSource = tblASTERROR

        Sort_grdColumns(grdASTERROR, "INIT_DATE".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        If Absx1.txtFor("USER_ID").Text = "" Then
            Absx1.txtFor("USER_ID").Text = ASCMAIN1.USER_ID
        End If

        If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") Then
            Set_Read_Only(Absx1.txtFor("USER_ID"), True)
        End If
        optLoadOption.CheckedIndex = 0

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "USER_ID" Then
            If ctl.Text <> "" Then
                'Call Click_Command("Load Errors")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "USER_ID" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Call Click_Command("Load Errors", e)
            End If
        End If
    End Sub
#End Region

#Region "grdASTERROR"

#End Region

    Private Sub grdASTERROR_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTERROR.AfterRowActivate
        txtSTACKTRACE.Text = grdASTERROR.ActiveRow.Cells("STACKTRACE").Text
        grpERR_TEXT.Text = grdASTERROR.ActiveRow.Cells("ERR_TEXT").Text
    End Sub
End Class