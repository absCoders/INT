Public Class ICTBRAN1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "GLTSEGM1", "*")
        End With

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If LookUp("ICTCOLL1", Absx1.txtFor("BRAND_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Collection Code that is the same as a Brand Code"
                End If

            Case "Edit"
            Case "Update"

                If Absx1.optFor("BRAND_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "Status is Mandatory"
                End If
                If Absx1.txtFor("BRAND_NAME").Text & "" = "" Then
                    EMsg &= vbCr & "Brand Name is Mandatory"
                End If
                If LookUp("ICTCOLL1", Absx1.txtFor("BRAND_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Collection Code that is the same as a Brand Code"
                End If

                If ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
                    OrElse ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG4") & "" <> "1" Then
                    ' NO NEED TO CHECK SEG4
                Else
                    Dim SEG4_CODE As String = Absx1.txtFor("SEG4_CODE").Text
                    If SEG4_CODE <> "" And SEG4_CODE <> Absx1.txtFor("BRAND_CODE").Text Then
                        If LookUp("ICTCOLL1", SEG4_CODE) Is Nothing _
                            AndAlso LookUp("ICTBRAN1", SEG4_CODE) Is Nothing _
                            AndAlso LookUp("GLTSEGM1", New String() {"4", SEG4_CODE}) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Segment 4"
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        'Dim sqlDelete = "CUST_CODE = '" & CUST_CODE & "'"
        'Update_Record_TDA("SPTDCOM2", sqlDelete)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        Dim SEG4_CODE As String = Absx1.txtFor("SEG4_CODE").Text
        If SEG4_CODE = "" Then
            SEG4_CODE = Absx1.txtFor("BRAND_CODE").Text
        End If
        Dim row As DataRow = LookUp("GLTSEGM1", New String() {"4", SEG4_CODE})
        If row Is Nothing Then
            Dim rowGLTSEGM1 As DataRow = dst.Tables("GLTSEGM1").NewRow
            rowGLTSEGM1.Item("ACCT_SEG_ID") = "4"
            rowGLTSEGM1.Item("ACCT_SEG_CODE") = SEG4_CODE
            rowGLTSEGM1.Item("ACCT_SEG_STATUS") = "A"
            rowGLTSEGM1.Item("ACCT_SEG_CLASS") = Absx1.txtFor("BRAND_CODE").Text
            rowGLTSEGM1.Item("ACCT_SEG_DESC") = Absx1.txtFor("BRAND_NAME").Text
            dst.Tables("GLTSEGM1").Rows.Add(rowGLTSEGM1)
            Update_Record_TDA("GLTSEGM1")
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text

        Dim FOLDER_NAME As String = ASCMAIN1.Folders("Images") & "\COLUMN_NAME\BRAND_CODE\"
        Dim IMAGE_NAME As String = BRAND_CODE & ".png"
        If My.Computer.FileSystem.FileExists(FOLDER_NAME & IMAGE_NAME) Then
            img.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
            img.Visible = True
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        btnPickColor.Visible = ScreenMode And (EntryMode = "New" Or EntryMode = "Edit")

        If Not tf Then
            img.Visible = False
        End If

    End Sub


    Private Sub btnPickColor_Click(sender As Object, e As EventArgs) Handles btnPickColor.Click
        Dim cDialog As New ColorDialog()
        cDialog.Color = lblBRAND_COLOR.Appearance.ForeColor  ' initial selection is current color.

        If (cDialog.ShowDialog() = DialogResult.OK) Then
            'lblBRAND_COLOR.Appearance.ForeColor = cDialog.Color 
            'numBRAND_COLOR.Appearance.ForeColor = cDialog.Color
            numBRAND_COLOR.Value = cDialog.Color.ToArgb
        End If

    End Sub

    Private Sub numBRAND_COLOR_ValueChanged(sender As Object, e As EventArgs) Handles numBRAND_COLOR.ValueChanged
        Dim rgb As Int64 = Val(numBRAND_COLOR.Value & "")
        Dim c As System.Drawing.Color = System.Drawing.Color.FromArgb(rgb)
        ' lblBRAND_COLOR.Appearance.ForeColor = c
        numBRAND_COLOR.Appearance.ForeColor = c
        numBRAND_COLOR.Appearance.ForeColorDisabled = c
    End Sub
#End Region
End Class