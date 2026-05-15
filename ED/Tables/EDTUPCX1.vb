Public Class EDTUPCX1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            'ASCMAIN1.sql = "Select * from EDT852T1 where EDI_DOC_SEQ_NO in " _
            '& " (Select Distinct EDI_DOC_SEQ_NO from EDT852T0 " _
            '& "where EDI_ITEM_CODE = :PARM1 and ITEM_CODE is Null)"
            'Call Create_TDA(.Tables.Add, "EDT852T1", "**", 0, False, "V")
        End With

        '  grdEDT852T1.DataSource = dst.Tables("EDT852T1")

        'ASCMAIN1.Add_Value_List(grdEDT852T1, "CUST_COMMENT_KEY")
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        'Call Update_Record_TDA("EDT852T1")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            'rowASFBASE1.Item("TERM_CODE") = "AAA"
            'Absx1.CtlFor("TERM_CODE").Text = "BBB"
        End If

        dst.EnforceConstraints = False
        ' Fill_Records("EDT852T1", New String() {Absx1.txtFor("EDI_ITEM_CODE").Text})

        dst.EnforceConstraints = True
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            ' dst.EnforceConstraints = False
            'dst.Tables("EDT852T1").Rows.Clear()
            ' dst.EnforceConstraints = False
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdEDT852T1.Visible = False ' tf
    End Sub
#End Region
End Class