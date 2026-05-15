Public Class GLRLIST1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i)
            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & "" = "" Then
                Absx1.CtlFor(z & "_CODE").Visible = False
            Else
                Absx1.CtlFor(z & "_CODE").Text = "Print " & ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & " List"
            End If
        Next

    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()

        Dim sql As String

        If Absx1.chkFor("ACCT_CODE").Checked Then
            sql = "Select * from GLTACCT1"
            Dim CODE_VALUES As String = tblASTDSQLA.Rows.Find("ACCT_CODE").Item("CODE_VALUES") & ""
            If CODE_VALUES <> "" Then
                sql = sql & " where ACCT_CODE in ('" & Replace(CODE_VALUES, ",", "','") & "')"
            End If
            dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1"))
        End If

        If Absx1.chkFor("SEG2_CODE").Checked _
        Or Absx1.chkFor("SEG3_CODE").Checked _
        Or Absx1.chkFor("SEG4_CODE").Checked Then

            sql = ""
            For i As Integer = 2 To 4
                Dim z As String = "SEG" & CStr(i) & "_CODE"
                If Absx1.chkFor(z).Checked Then
                    sql = sql & " OR (ACCT_SEG_ID = '" & CStr(i) & "'"
                    Dim CODE_VALUES As String = tblASTDSQLA.Rows.Find(z).Item("CODE_VALUES") & ""
                    If CODE_VALUES <> "" Then
                        sql = sql & " and ACCT_SEG_CODE in ('" & Replace(CODE_VALUES, ",", "','") & "')"
                    End If
                    sql = sql & ")"
                End If
            Next
            If sql <> "" Then
                sql = " where " & Mid$(sql, 4)
            End If
            sql = "Select * from GLTSEGM1 " & sql
            dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTSEGM1"))
        End If


    End Sub

    Public Overrides Sub Print_Report()
        If Absx1.chkFor("ACCT_CODE").Checked Then
            Generate_Report("GLRLIST1")
        End If

        If Absx1.chkFor("SEG2_CODE").Checked _
        Or Absx1.chkFor("SEG3_CODE").Checked _
        Or Absx1.chkFor("SEG4_CODE").Checked Then

            CR_params.Add("SEG2", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
            CR_params.Add("SEG3", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
            CR_params.Add("SEG4", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
            Generate_Report("GLRLIST2")
        End If
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                If Not Absx1.chkFor("ACCT_CODE").Checked _
                And Not Absx1.chkFor("SEG2_CODE").Checked _
                And Not Absx1.chkFor("SEG3_CODE").Checked _
                And Not Absx1.chkFor("SEG4_CODE").Checked Then
                    EMsg = EMsg & vbCr & "How about printing something"
                End If
        End Select
    End Sub
End Class