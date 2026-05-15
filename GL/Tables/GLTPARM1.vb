Public Class GLTPARM1


    Overrides Sub Prepare_for_View_Lookup_Special(
     ByVal ctl As Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "GL_PARM_SOFT_CLOSE_YYYYPP"
                sql_where = "OPS_YYYYPP >= (SELECT PERIOD_CALC(GL_PARM_CURRENT_YYYYPP,-1) FROM GLTPARM1) and OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"


        End Select
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                Dim GL_PARM_SOFT_CLOSE_YYYYPP As String = Absx1.txtFor("GL_PARM_SOFT_CLOSE_YYYYPP").Text
                If GL_PARM_SOFT_CLOSE_YYYYPP = "" Then
                    EMsg &= vbCr & "You must specify a Soft Close Period"
                Else
                    Dim row As DataRow = LookUp("GLTPARM2", GL_PARM_SOFT_CLOSE_YYYYPP)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Soft Close Period specified"
                    Else
                        Dim GL_PARM_CURRENT_YYYYPP As String = Absx1.txtFor("GL_PARM_CURRENT_YYYYPP").Text
                        Dim GL_PARM_CURRENT_YYYYPP_prev As String = ASCMAIN1.Period_Calc(GL_PARM_CURRENT_YYYYPP, -1)
                        If GL_PARM_SOFT_CLOSE_YYYYPP < GL_PARM_CURRENT_YYYYPP_prev Then
                            EMsg &= vbCr & $"Soft Close Period must not be prior to {GL_PARM_CURRENT_YYYYPP_prev}"
                        Else
                            If GL_PARM_SOFT_CLOSE_YYYYPP >= ASCMAIN1.CYP Then
                                EMsg &= vbCr & $"Soft Close Period must be prior to {ASCMAIN1.CYP}"
                            End If
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        ASCMAIN1.sql = "Delete from ASTCODE1 where TABLE_NAME = :PARM1 and COLUMN_NAME = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"GLTSEGM1", "ACCT_SEG_ID"})

        ASCMAIN1.sql = "Insert into ASTCODE1 Values (:PARM1,:PARM2,:PARM3,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "2", Absx1.txtFor("GL_PARM_SEG2_DESC").Text})
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "3", Absx1.txtFor("GL_PARM_SEG3_DESC").Text})
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"GLTSEGM1", "ACCT_SEG_ID", "4", Absx1.txtFor("GL_PARM_SEG4_DESC").Text})
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

    End Sub

    Overrides Sub Clear_Record_Special()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

    End Sub

#End Region

End Class