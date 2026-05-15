Public Class FAFFAMFA

    Public ASSET_NO As String
    Public rowAPTINVHX As DataRow
    Public rowFATFAMF1 As DataRow
    Public ASSET_AMT As Decimal

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Create_TDA(dst.Tables.Add, "FATFAMF1", "*")

        Get_PARM("FATPARM1")

        rowFATFAMF1 = dst.Tables("FATFAMF1").NewRow
        With rowFATFAMF1

            Dim OPS_YYYYPP_AP As String = ASCMAIN1.CYP ' maybe -1
            Dim INIT_DATE As Date = DATETIME_STAMP

            .Item("ASSET_NO") = "......"
            If rowAPTINVHX IsNot Nothing Then
                .Item("ASSET_DESC") = rowAPTINVHX.Item("INV_REF")
                .Item("ASSET_CLASS_CODE") = rowAPTINVHX.Item("ASSET_CLASS_CODE")

                OPS_YYYYPP_AP = rowAPTINVHX.Item("OPS_YYYYPP")
                INIT_DATE = rowAPTINVHX.Item("INIT_DATE")
            End If


            .Item("ASSET_STATUS") = "C"
            .Item("ASSET_DEPR_CODE") = "SL"

            Dim dtes() As Date = ASCMAIN1.Get_Dates(OPS_YYYYPP_AP)
            If Format(INIT_DATE, "yyyyMM") < OPS_YYYYPP_AP Then
                INIT_DATE = dtes(1) ' CDate(Mid(OPS_YYYYPP_AP, 5, 2) & "/01/" & Mid(OPS_YYYYPP_AP, 1, 4))
            ElseIf Format(INIT_DATE, "yyyyMM") > OPS_YYYYPP_AP Then
                INIT_DATE = dtes(dtes.Length - 1)
            End If

            .Item("ASSET_DATE") = INIT_DATE ' rowAPTINVHX.Item("INV_DATE")
            .Item("ASSET_DATE_IN_SERVICE") = INIT_DATE ' rowAPTINVHX.Item("INV_DATE")

            If rowAPTINVHX IsNot Nothing Then
                .Item("VEND_CODE") = rowAPTINVHX.Item("VEND_CODE")
                .Item("VOUCHER_NO") = rowAPTINVHX.Item("VOUCHER_NO")
                .Item("VOUCHER_LNO") = rowAPTINVHX.Item("VOUCHER_LNO")
                .Item("INVOICE_NOTES") = rowAPTINVHX.Item("INV_REF")
                .Item("VEND_CODE") = rowAPTINVHX.Item("VEND_CODE")
                .Item("VEND_NAME") = rowAPTINVHX.Item("VEND_NAME")
            End If

            .Item("INIT_DATE") = Now
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = Now
            .Item("LAST_OPER") = ASCMAIN1.USER_ID

            '.Item("OPS_YYYYPP") = Get_YP_from_Date("OPS_YYYYPP", Format(rowAPTINVHX.Item("INV_DATE"), "MM/dd/yyyy")) ' ASCMAIN1.CYP ' Get_YP_from_Date() '  ASCMAIN1.CYP
            '.Item("OPS_YYYYPP_IN_SERVICE") = Get_YP_from_Date("OPS_YYYYPP_IN_SERVICE", Format(rowAPTINVHX.Item("INV_DATE"), "MM/dd/yyyy")) ' ASCMAIN1.CYP ' Get_YP_from_Date() '  ASCMAIN1.CYP
            .Item("OPS_YYYYPP") = OPS_YYYYPP_AP ' Get_YP_from_Date("OPS_YYYYPP", Format(INIT_DATE, "MM/dd/yyyy")) ' ASCMAIN1.CYP ' Get_YP_from_Date() '  ASCMAIN1.CYP
            .Item("OPS_YYYYPP_IN_SERVICE") = OPS_YYYYPP_AP ' Get_YP_from_Date("OPS_YYYYPP_IN_SERVICE", Format(INIT_DATE, "MM/dd/yyyy")) ' ASCMAIN1.CYP ' Get_YP_from_Date() '  ASCMAIN1.CYP
            .Item("ASSET_AMT") = ASSET_AMT ' rowAPTINVHX.Item("INV_LINE_AMT")
            .Item("ASSET_ACTION") = "D"
        End With
        dst.Tables("FATFAMF1").Rows.Add(rowFATFAMF1)

        If rowAPTINVHX IsNot Nothing Then
            Absx1.txtFor("ASSET_CLASS_CODE").Text = rowAPTINVHX.Item("ASSET_CLASS_CODE")
            rowFATFAMF1.Item("ASSET_LIFE_MOS") = Get_Life_from_Class()
        Else
            Absx1.txtFor("ASSET_CLASS_CODE").Text = ""
        End If

        txtVEND_NAME.Visible = (rowFATFAMF1.Item("VEND_CODE") & "" = "")
        txtVEND_NAME_from_VEND_CODE.Visible = (rowFATFAMF1.Item("VEND_CODE") & "" <> "")

        'If rowAPTINVHX.Item("INV_REF") & "" <> "" Then
        '    Absx1.txtFor("ASSET_DESC").Text = rowAPTINVHX.Item("INV_REF") & ""
        'End If

        'dst.Tables("FATFAMF1").AcceptChanges()
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
        X.EndCurrentEdit()
        'Get_YP_from_Date()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
    (ByVal ctl As Windows.Forms.Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            '        Case "ORDR_FORM_CODE"
            '            sql_where = "ORDR_FORM_STATUS = 'A'"
        End Select
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "ORDR_FORM_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_SOTFORM2()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "ORDR_FORM_CODE"
            '    Prepare_SOTFORM2()
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ASSET_CLASS_CODE"
                Get_Life_from_Class()

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ASSET_DATE_IN_SERVICE"

                'If Absx1.dteFor("ASSET_DATE_IN_SERVICE").IsInEditMode Then
                '    Exit Sub
                'End If

                'Set_Value_for_OPS_YYYYPP_IN_SERVICE()

        End Select
    End Sub

#End Region

    Sub Set_Value_for_OPS_YYYYPP_IN_SERVICE()
        If Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value & "" <> "" Then
            Dim YP As String = Get_YP_from_Date("OPS_YYYYPP_IN_SERVICE")
            'rowFATFAMF1.Item("OPS_YYYYPP_IN_SERVICE") = YP
            Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Value = YP
        End If
    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
        X.EndCurrentEdit()

        Dim EMsg As String = ""
        If Absx1.txtFor("ASSET_DESC").Text = "" Then
            EMsg &= vbCr & "Asset Description is Mandatory"
        End If
        If Absx1.txtFor("ASSET_DEPR_CODE").Text = "" Then
            EMsg &= vbCr & "Asset Depreciation Method is Mandatory"
        Else
            Dim ASSET_DEPR_CODE As String = Absx1.txtFor("ASSET_DEPR_CODE").Text
            Dim rowFATDEPM1 As DataRow = Lookup("FATDEPM1", ASSET_DEPR_CODE)
            If rowFATDEPM1 Is Nothing Then
                EMsg &= vbCr & "Invalid Value specified for Asset Depreciation Method"
            End If
        End If
        If Absx1.txtFor("ASSET_CLASS_CODE").Text = "" Then
            EMsg &= vbCr & "Asset Class is Mandatory"
        Else
            Dim ASSET_CLASS_CODE As String = Absx1.txtFor("ASSET_CLASS_CODE").Text
            Dim rowFATFACL1 As DataRow = Lookup("FATFACL1", ASSET_CLASS_CODE)
            If rowFATFACL1 Is Nothing Then
                EMsg &= vbCr & "Invalid Value specified for Asset Class"
            End If
        End If
        If Val(Absx1.numFor("ASSET_LIFE_MOS").Value & "") <= 0 Then
            EMsg &= vbCr & "Asset Life is Mandatory"
        End If
        If Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value & "" = "" Then
            EMsg &= vbCr & "Asset In Service Date is Mandatory"
        Else
            Dim YP As String = Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Text
            'If YP <= ROWs("FATPARM1").Item("FA_PARM_DEPR_LAST_YP_UPDATED") Then
            '    EMsg &= vbCr & $"Asset In Service Period {YP} must be later than last period Depreciation was updated {ROWs("FATPARM1").Item("FA_PARM_DEPR_LAST_YP_UPDATED") }"
            'End If
            'If YP < Mid(ASCMAIN1.CYP, 1, 4) & "01" Then
            '    EMsg &= vbCr & $"Asset In Service Period {YP} must be later than 1st period of current year {Mid(ASCMAIN1.CYP, 1, 4) }"
            'End If
        End If
        If Val(Absx1.numFor("ASSET_AMT").Value & "") <= 0 Then
            EMsg &= vbCr & "Asset Amount cannot be 0 or Negative"
        End If

        If rowAPTINVHX IsNot Nothing Then
            If Val(Absx1.numFor("ASSET_AMT").Value & "") > ASSET_AMT Then
                EMsg &= vbCr & $"Asset Amount cannot be more than Distribution {Format(ASSET_AMT, "$#,##0.00")}"
            End If
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Add Asset")
            Exit Sub
        End If

        If rowAPTINVHX IsNot Nothing Then
            If Val(Absx1.numFor("ASSET_AMT").Value & "") <> ASSET_AMT Then
                If MsgBox("Asset Amount Capitalized is not the same as the Total Distribution Amount." & vbCrLf & "Continue with Update?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
            End If
        End If


        rowFATFAMF1.Item("ASSET_BAL") = rowFATFAMF1.Item("ASSET_AMT")
        ASSET_NO = ASCMAIN1.Next_Control_No("FATFAMF1.ASSET_NO")
        rowFATFAMF1.Item("ASSET_NO") = ASSET_NO

        BeginTrans()
        Update_Record_TDA("FATFAMF1")
        CommitTrans($"Asset {ASSET_NO} has been added")

        'MsgBox($"Asset {ASSET_NO} has been added", MsgBoxStyle.OkOnly, "Success")
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Function Get_YP_from_Date(COLUMN_NAME As String, Optional DT As String = "") As String
        Dim ASSET_DATE_IN_SERVICE As Date = Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value
        If DT <> "" Then
            ASSET_DATE_IN_SERVICE = CDate(DT)
        End If
        Dim OPS_YYYYPP As String = Format(ASSET_DATE_IN_SERVICE, "yyyyMM")

        ' Absx1.txtFor(COLUMN_NAME).Text = OPS_YYYYPP
        Return OPS_YYYYPP
    End Function

    Function Get_Life_from_Class() As Integer
        Dim ASSET_CLASS_CODE As String = Absx1.txtFor("ASSET_CLASS_CODE").Text
        Dim ASSET_LIFE_MOS As Integer

        Dim rowFATFACL1 As DataRow = Lookup("FATFACL1", ASSET_CLASS_CODE)
        If rowFATFACL1 IsNot Nothing Then
            ASSET_LIFE_MOS = Val(rowFATFACL1.Item("ASSET_LIFE_MOS") & "")
            Absx1.numFor("ASSET_LIFE_MOS").Value = ASSET_LIFE_MOS
        End If

        Return ASSET_LIFE_MOS
    End Function

    Private Sub dteASSET_DATE_IN_SERVICE_ValueChanged(sender As Object, e As EventArgs) Handles dteASSET_DATE_IN_SERVICE.ValueChanged

    End Sub

    Private Sub dteASSET_DATE_IN_SERVICE_LostFocus(sender As Object, e As EventArgs) Handles dteASSET_DATE_IN_SERVICE.LostFocus
        Set_Value_for_OPS_YYYYPP_IN_SERVICE()
    End Sub

    Private Sub dteASSET_DATE_IN_SERVICE_AfterCloseUp(sender As Object, e As EventArgs) Handles dteASSET_DATE_IN_SERVICE.AfterCloseUp
        Set_Value_for_OPS_YYYYPP_IN_SERVICE()
    End Sub
End Class