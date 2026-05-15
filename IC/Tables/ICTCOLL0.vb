Public Class ICTCOLL0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"

                If Absx1.optFor("HC_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "Status is Mandatory"
                End If

                ' CANNOT IMPLEMENT THIS LOGIC BECAUSE OF A CHICKEN/EGG ISSUE - NEED TO CREATE THE HICOLL BEFORE CREATING ANY COLLECTIONS
                'Dim COLLECTION_CODE As String = Absx1.txtFor("COLLECTION_CODE").Text
                'If COLLECTION_CODE = "" Then
                '    EMsg &= vbCr & "Key Collection Code is Mandatory"
                'Else
                '    If LookUp("ICTCOLL1", COLLECTION_CODE) Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value Specified"
                '    Else
                '        If cdr.Item("BRAND_CODE") & "" <> Absx1.txtFor("BRAND_CODE").Text Then
                '            EMsg &= vbCr & "Collection " & COLLECTION_CODE & " belongs to Brand " & cdr.Item("BRAND_CODE")
                '        End If
                '    End If
                'End If


                Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
                If BRAND_CODE = "" Then
                    EMsg &= vbCr & "Brand Code is Mandatory"
                Else
                    If LookUp("ICTBRAN1", BRAND_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Brand Code"
                    Else
                        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                            ' DO NOT CHECK THIS FOR AHA
                        Else
                            ASCMAIN1.sql = "Select COLLECTION_CODE, BRAND_CODE from ICTCOLL1" _
                                & " where HC_CODE = '" & Absx1.txtFor("HC_CODE").Text & "'" _
                                & "   and NVL(BRAND_CODE,'?') <> '" & BRAND_CODE & "'"
                            Dim row As DataRow = ASCDATA1.GetDataRow

                            If row IsNot Nothing Then
                                EMsg &= vbCr & "Collection " & row.Item("COLLECTION_CODE") & " belongs to this High Collection yet belongs to Brand " & row.Item("BRAND_CODE")
                            End If
                        End If
                    End If
                End If

                Dim COLLECTION_CODE As String = Absx1.txtFor("COLLECTION_CODE").Text
                If COLLECTION_CODE = "" Then
                    EMsg &= vbCr & "Collection Code is Mandatory"
                Else
                    If LookUp("ICTCOLL1", COLLECTION_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Collection Code"
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
    End Sub

    Overrides Sub Show_Record_Special()
        'EnforceConstraints(False)
        'EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
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
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
    End Sub
#End Region
End Class