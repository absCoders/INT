Imports System.Reflection

Public Class ICTCOLL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        grp3PL.Visible = (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT")

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "GLTSEGM1", "*")
            Create_TDA(.Tables.Add, "ICTCOLL0", "*")

            ASCMAIN1.sql = "SELECT SE.COLLECTION_CODE,
                SE.STATE_CODE,TS.STATE_NAME,
                SE.CUST_CODE,C1.CUST_NAME,
                SE.INIT_OPER,SE.INIT_DATE,SE.LAST_OPER,SE.LAST_DATE 
                FROM SATEXCL1 SE 
                JOIN TATSTATE TS ON TS.STATE_CODE=SE.STATE_CODE
                JOIN ARTCUST1 C1 ON C1.CUST_CODE=SE.CUST_CODE
                WHERE COLLECTION_CODE = :PARM1 OR :PARM2='1'"
            Create_TDA(.Tables.Add, "SATEXCL1", "**", 0, True, "VV")
            Fill_Records("SATEXCL1", {"", "1"})
            grdSATEXCL1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdSATEXCL1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            Me.UnsubscribeBeforeRowUpdate(grdSATEXCL1)

        End With


        grdSATEXCL1.DataSource = dst.Tables("SATEXCL1")
        RepositionUpdateExclusionsButton()


        AUDIT.Add("GLTSEGM1", "N")
        AUDIT.Add("ICTCOLL0", "N")

        AUDIT.Add("SATEXCL1", "NED")
    End Sub

  
#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If LookUp("ICTBRAN1", Absx1.txtFor("COLLECTION_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Collection Code that is the same as a Brand Code"
                End If

            Case "Edit"
            Case "Update"

                If Absx1.optFor("COLLECTION_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "Status is Mandatory"
                End If
                If Absx1.optFor("COLLECTION_GENDER").Value & "" = "" Then
                    EMsg &= vbCr & "Gender is Mandatory"
                End If
                If Absx1.txtFor("COLLECTION_NAME").Text & "" = "" Then
                    EMsg &= vbCr & "Collection Name is Mandatory"
                Else
                    If ASCMAIN1.CLIENT = "INT" Then
                        Dim rx As String = "[^a-zA-Z0-9 .$]" ' Allow Upper/Lower case, numbers, space, dot
                        Dim r As New System.Text.RegularExpressions.Regex(rx)
                        If r.IsMatch(Absx1.txtFor("COLLECTION_NAME").Text) Then
                            EMsg &= vbCr & "Collection Name has Special Characters which are not allowed"
                        End If
                    End If
                End If

                If LookUp("ICTBRAN1", Absx1.txtFor("COLLECTION_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Collection Code that is the same as a Brand Code"
                End If

                Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
                If BRAND_CODE = "" Then
                    EMsg &= vbCr & "Brand Code is Mandatory"
                Else
                    If LookUp("ICTBRAN1", BRAND_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Brand"
                    End If
                End If

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Dim BRAND_CODE_3PL As String = Absx1.txtFor("BRAND_CODE_3PL").Text
                    If BRAND_CODE_3PL = "" Then
                        EMsg &= vbCr & "3PL Brand Code is Mandatory"
                    Else
                        If LookUp("ICT3PLB1", BRAND_CODE_3PL) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for 3PL Brand Code"
                        End If
                    End If
                    Dim SUB_BRAND_CODE_3PL As String = Absx1.txtFor("SUB_BRAND_CODE_3PL").Text
                    If SUB_BRAND_CODE_3PL = "" Then
                        'EMsg &= vbCr & "3PL Sub-Brand Code is Mandatory"
                    Else
                        If LookUp("ICT3PLB2", SUB_BRAND_CODE_3PL) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for 3PL Sub-Brand Code"
                        End If
                    End If
                End If

                Dim HC_CODE As String = ""

                If chkNewHighColl.Checked Then
                    HC_CODE = txtHC_CODE.Text
                    If LookUp("ICTCOLL0", HC_CODE) IsNot Nothing Then
                        EMsg &= vbCr & $"Invalid Value Specified for High Collection {HC_CODE} - already exists"
                    End If
                    Dim HC_NAME As String = Absx1.txtFor("HC_NAME").Text
                    If HC_NAME = "" Then
                        EMsg &= vbCr & "HC Name is Mandatory"
                    End If
                    Dim CHECKBOOK As String = Absx1.txtFor("CHECKBOOK").Text
                    If CHECKBOOK = "" Then
                        EMsg &= vbCr & "Checkbook is Mandatory"
                    Else
                        If LookUp("SPTCWRXC", CHECKBOOK) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Checkbook"
                        End If
                    End If
                Else
                    HC_CODE = Absx1.txtFor("HC_CODE").Text
                    If HC_CODE = "" Then
                        EMsg &= vbCr & "High Collection Code is Mandatory"
                    Else

                        If LookUp("ICTCOLL0", HC_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for High Collection Code"
                        Else
                            If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                                'DISABLE THIS CHECK AS PER TJ 09/04/2015
                            Else
                                If cdr.Item("BRAND_CODE") & "" = "" Then
                                    EMsg &= vbCr & "High Collection " & HC_CODE & " is not assigned to a Brand"
                                Else
                                    If cdr.Item("BRAND_CODE") & "" <> BRAND_CODE And BRAND_CODE <> "" Then
                                        EMsg &= vbCr & "High Collection " & HC_CODE & " is already assigned to Brand " & cdr.Item("BRAND_CODE") & ""
                                    End If
                                End If
                            End If

                            If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                                'DISABLE THIS CHECK AS PER TJ 09/04/2015
                            Else
                                ASCMAIN1.sql = "Select Distinct BRAND_CODE from ICTCOLL1 where HC_CODE = '" & HC_CODE & "'"
                                Dim tbl As DataTable = ASCDATA1.GetDataTable
                                If tbl.Rows.Count > 1 Then
                                    EMsg &= vbCr & "High Collection already assigned to Multiple Brands - Please Correct"
                                Else
                                    If tbl.Rows.Count = 1 Then
                                        If tbl.Rows(0).Item("BRAND_CODE") & "" <> BRAND_CODE Then
                                            EMsg &= vbCr & "High Collection " & HC_CODE & " is already assigned to Collections in Brand " & tbl.Rows(0).Item("BRAND_CODE")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
                    OrElse ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG4") & "" <> "1" Then
                    ' NO NEED TO CHECK SEG4
                Else
                    Dim SEG4_CODE As String = Absx1.txtFor("SEG4_CODE").Text
                    If SEG4_CODE <> "" And SEG4_CODE <> Absx1.txtFor("COLLECTION_CODE").Text Then
                        If LookUp("ICTCOLL1", SEG4_CODE) Is Nothing _
                            AndAlso LookUp("ICTBRAN1", SEG4_CODE) Is Nothing _
                            AndAlso LookUp("GLTSEGM1", New String() {"4", SEG4_CODE}) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Segment 4"
                        End If
                    End If
                End If

                ASCMAIN1.sql = "Select ICTCOLL0.HC_CODE, ICTCOLL0.COLLECTION_CODE, ICTCOLL1.HC_CODE HC_CODE2" & vbCrLf _
                    & " from ICTCOLL0, ICTCOLL1" & vbCrLf _
                    & " where ICTCOLL1.COLLECTION_CODE = ICTCOLL0.COLLECTION_CODE" & vbCrLf _
                    & "   and ICTCOLL0.HC_CODE <> ICTCOLL1.HC_CODE"
                Dim MISCORRELATIONS As String = ""
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim mc_HC_CODE As String = row.Item("HC_CODE") & ""
                    Dim mc_COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""
                    Dim mc_HC_CODE2 As String = row.Item("HC_CODE2") & ""
                    MISCORRELATIONS &= vbCr & $"Mis-Correlation - HC {mc_HC_CODE} -> Key Coll {mc_COLLECTION_CODE} -> HC {mc_HC_CODE2}"
                Next
                If MISCORRELATIONS <> "" Then
                    MsgBox("Warning - there are mis-corrleations:" & MISCORRELATIONS, MsgBoxStyle.OkOnly, "Warning: Mis-Correlations")
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        'Dim sqlDelete = "CUST_CODE = '" & CUST_CODE & "'"
        'Update_Record_TDA("SPTDCOM2", sqlDelete)
        If chkNewHighColl.Checked Then
            Dim HC_NAME As String = Absx1.txtFor("HC_NAME").Text
            Dim CHECKBOOK As String = Absx1.txtFor("CHECKBOOK").Text
            rowASFBASE1.Item("HC_CODE") = txtHC_CODE.Text
            Absx1.txtFor("CHECKBOOK").Text = CHECKBOOK
            Absx1.txtFor("HC_NAME").Text = HC_NAME
            'Absx1.txtFor("HC_CODE").Text = txtHC_CODE.Text
        End If

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        Dim SEG4_CODE As String = Absx1.txtFor("SEG4_CODE").Text
        If SEG4_CODE = "" Then
            SEG4_CODE = Absx1.txtFor("COLLECTION_CODE").Text
        End If
        Dim row As DataRow = LookUp("GLTSEGM1", New String() {"4", SEG4_CODE})
        If row Is Nothing Then
            Dim rowGLTSEGM1 As DataRow = dst.Tables("GLTSEGM1").NewRow
            rowGLTSEGM1.Item("ACCT_SEG_ID") = "4"
            rowGLTSEGM1.Item("ACCT_SEG_CODE") = SEG4_CODE
            rowGLTSEGM1.Item("ACCT_SEG_STATUS") = "A"
            rowGLTSEGM1.Item("ACCT_SEG_CLASS") = Absx1.txtFor("BRAND_CODE").Text
            rowGLTSEGM1.Item("ACCT_SEG_DESC") = Absx1.txtFor("COLLECTION_NAME").Text
            dst.Tables("GLTSEGM1").Rows.Add(rowGLTSEGM1)
            Update_Record_TDA("GLTSEGM1")
        End If

        If chkNewHighColl.Checked Then
            Dim rowICTCOLL0 As DataRow = dst.Tables("ICTCOLL0").NewRow
            With rowICTCOLL0
                .Item("HC_CODE") = txtHC_CODE.Text
                .Item("HC_NAME") = Absx1.txtFor("HC_NAME").Text
                .Item("HC_STATUS") = "A"
                .Item("BRAND_CODE") = Absx1.txtFor("BRAND_CODE").Text
                .Item("COLLECTION_CODE") = Absx1.txtFor("COLLECTION_CODE").Text
                .Item("CHECKBOOK") = Absx1.txtFor("CHECKBOOK").Text
            End With
            dst.Tables("ICTCOLL0").Rows.Add(rowICTCOLL0)
            Update_Record_TDA("ICTCOLL0")

        End If

        Update_Record_TDA("SATEXCL1")
    End Sub

    Overrides Sub Show_Record_Special()
        'EnforceConstraints(False)
        'EnforceConstraints(True)

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            'grp3PL.Enabled = (EntryMode = "New")

            Fill_Records("SATEXCL1", {Absx1.txtFor("COLLECTION_CODE").Text, "0"})

            ' LMB 01/04/22 As previously mentioned will you please allow me to EDIT 3pl brand and sub-brand
            grp3PL.Enabled = True
            '    Set_Read_Only(grp3PL, Not (EntryMode = "New"))
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            dst.Tables("SATEXCL1").Rows.Clear()
            Fill_Records("SATEXCL1", {"", "1"})
        End If
        'If ScreenMode Then
        '    EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {""}
        '        dst.Tables(TABLE_NAME).Rows.Clear()
        '    Next
        '    EnforceConstraints(True)
        'End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        'grdSPTDCOM2.Enabled = tf
        grdSATEXCL1.Enabled = (Not tf) Or {"New", "Edit"}.Contains(EntryMode)
        grdSATEXCL1.DisplayLayout.Bands(0).Columns("COLLECTION_CODE").Hidden = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        grdSATEXCL1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        grdSATEXCL1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        btnUpdateExclusions.Visible = Not tf
        btnUpdateExclusions.Enabled = False
        If tf Then
            If EntryMode = "New" Then
                chkNewHighColl.Visible = True

                ' LBM 01/04/2022 They have moved everything to 87. This should be an automatic default going forward.
                Absx1.txtFor("SUB_BRAND_CODE_3PL").Text = "87"
                Absx1.txtFor("BRAND_CODE_3PL").Text = "87"
            End If
        Else
            chkNewHighColl.Checked = False
            chkNewHighColl.Visible = False
            Setup_New_HC()
        End If
    End Sub

    Overrides Sub Load_Popup_Menus()
        MyBase.Load_Popup_Menus()
        Load_Popup_Menu(grdSATEXCL1, "S", "Show Filter")
    End Sub


#End Region

    Private Sub chkNewHighColl_CheckedChanged(sender As Object, e As EventArgs) Handles chkNewHighColl.CheckedChanged
        Setup_New_HC()
    End Sub

    Sub Setup_New_HC()
        txtHC_NAME.ReadOnly = Not chkNewHighColl.Checked
        txtCHECKBOOK.ReadOnly = Not chkNewHighColl.Checked
        Set_Read_Only_for_ctl(txtHC_NAME, Not chkNewHighColl.Checked)
        Set_Read_Only_for_ctl(txtCHECKBOOK, Not chkNewHighColl.Checked)
        If chkNewHighColl.Checked Then
            txtHC_CODE.Text = Absx1.txtFor("HC_CODE").Text
            Absx1.txtFor("HC_CODE").Text = ""
        End If
        txtHC_CODE.Visible = chkNewHighColl.Checked
    End Sub

    Private Sub grdSATEXCL1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATEXCL1.InitializeRow
        If e.Row.IsAddRow Then
            If (ScreenMode) Then
                e.Row.Cells("COLLECTION_CODE").Value = Absx1.txtFor("COLLECTION_CODE").Text
            End If
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP  
        End If
    End Sub


    Private Sub grdSATEXCL1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATEXCL1.ClickCellButton
        Dim codeValue As String = ""
        Dim column = e.Cell.Column.Key
      
        If {"COLLECTION_CODE","STATE_CODE","CUST_CODE"}.Contains(column) Then
            codeValue = Get_Code(column) & ""
        End If
        
        If codeValue <> "" Then e.Cell.Value = codeValue
    End Sub

    Private Sub grdSATEXCL1_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATEXCL1.AfterCellUpdate
        Select Case e.Cell.Column.Key 
            
            Case "STATE_CODE" 
                Dim drState As DataRow = LookUp("TATSTATE", e.Cell.Value)
                e.Cell.Row.Cells("STATE_NAME").Value = drState?.Item("STATE_NAME") & ""

            Case "CUST_CODE"
                 Dim drCust As DataRow = LookUp("ARTCUST1", e.Cell.Value)
                e.Cell.Row.Cells("CUST_NAME").Value = drCust?.Item("CUST_NAME") & ""

        End Select
    End Sub

    Private Sub btnUpdateExclusions_Click(sender As Object, e As EventArgs) Handles btnUpdateExclusions.Click
        grdSATEXCL1.ActiveRow = Nothing 'ActiveRow being the AddRow was causing an error when standards called EndCurrentEdit
        BeginTrans()
        'WriteAuditTrail("SATEXCL1")
        Update_Record_TDA("SATEXCL1")
        CommitTrans("Exclusions updated")
        btnUpdateExclusions.Enabled = False
    End Sub

    Private Sub grdSATEXCL1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSATEXCL1.AfterRowUpdate
        btnUpdateExclusions.Enabled = (Not ScreenMode) And True
    End Sub

    Private Sub grdSATEXCL1_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSATEXCL1.BeforeRowUpdate
        EMsg2 = ""
        Dim coll As DataRow = LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Value & "")
        If coll Is Nothing Then
            EMsg2 = "Invalid collection code"
            e.Cancel = True
        End If

        Dim state As DataRow = LookUp("TATSTATE", e.Row.Cells("STATE_CODE").Value & "")
        If state Is Nothing Then
            e.Cancel = True
            EMsg2 = "Invalid state code"
        End If

        Dim cust As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value & "")
        If cust Is Nothing Then
            e.Cancel = True
            EMsg2 = "Invalid customer code"
        End If

        If Not loading_grd_from_Excel And EMsg2 <> "" Then MsgBox(EMsg2, MsgBoxStyle.OkOnly, "Cannot Update Row")
    End Sub

    Private Sub grdSATEXCL1_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSATEXCL1.AfterRowsDeleted
        btnUpdateExclusions.Enabled = (Not ScreenMode) And True
    End Sub

    Private Sub splControls_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles splControls.SplitterMoved
        RepositionUpdateExclusionsButton()
    End Sub

    Private Sub RepositionUpdateExclusionsButton()
        btnUpdateExclusions.Parent = grdSATEXCL1
        btnUpdateExclusions.Location = New Point(btnUpdateExclusions.Location.X,2)
        btnUpdateExclusions.BringToFront()
        btnUpdateExclusions.Enabled = False
    End Sub
End Class