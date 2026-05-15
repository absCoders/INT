Public Class ASFDEPL1
    Dim F As New List(Of String)

    Private Sub ASFDEPL1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        F.Add("ASTCODE1")
        F.Add("ASTDSQLA")
        F.Add("ASTDSQLB")
        F.Add("ASTDSQLC")
        F.Add("ASTDSQLD")
        F.Add("ASTDSQLE")
        F.Add("ASTDSQLF")
        F.Add("ASTDSQLG")
        F.Add("ASTDSQLH")
        F.Add("ASTDSQLJ")
        F.Add("ASTDSQLK")
        F.Add("ASTDSQLS")
        F.Add("ASTDSQLV")
        F.Add("ASTDSQLW")
        F.Add("ASTDSQLX")
        F.Add("ASTDSQLY")
        F.Add("ASTFFMT1")
        F.Add("ASTFILT1")
        F.Add("ASTMENU1")
        F.Add("ASTMRUL1")
        F.Add("ASTMTKC1")
        F.Add("ASTREQF1")
        F.Add("ASTSECM1")
        F.Add("ASTTABD1")
        F.Add("ASTTTIP1")
        F.Add("ASTVIEW1")
        F.Add("ASTVIEW2")
        F.Add("ASTVIEW3")
        F.Add("ASTVIEW4")

        grd.DataSource = F

    End Sub

    Private Sub cmdDeploy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeploy.Click

        If ASCMAIN1.DBS_SERVER <> "" And chkASTTables.Checked Then
            MsgBox("Deployment must be done while logged into Development Machine, Stupidhead")
            Exit Sub
        End If

        If chkAssemblies.Checked Then
            Call Deploy_Assemblies()
        End If
        If chkASTTables.Checked Then
            Call Deploy_AST_Tables()
        End If

        MsgBox("Deployment Complete")
        Me.Close()

    End Sub

    Sub Deploy_AST_Tables()

        ASCMAIN1.T = ASCMAIN1.oraCon.BeginTransaction

        Dim SFX As String = ASCMAIN1.DBS_COMPANY
        'SFX = "DRC"

        Dim reverse_deploy As Boolean = False

        For Each TABLE_NAME As String In F
            'Stop
            If reverse_deploy Then
                ASCMAIN1.oraCmd.CommandText = "DELETE FROM " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME
                ASCMAIN1.oraCmd.ExecuteNonQuery()
                ASCMAIN1.oraCmd.CommandText = "INSERT INTO " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & " SELECT * FROM " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & "@" & SFX
                ASCMAIN1.oraCmd.ExecuteNonQuery()
            Else
                ASCMAIN1.oraCmd.CommandText = "DELETE FROM " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & "@" & SFX & ""
                ASCMAIN1.oraCmd.ExecuteNonQuery()
                ASCMAIN1.oraCmd.CommandText = "INSERT INTO " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & "@" & SFX & " SELECT * FROM " & TABLE_NAME & ""
                ASCMAIN1.oraCmd.ExecuteNonQuery()
            End If

            'ASCMAIN1.oraCmd.CommandText = "DELETE FROM TST." & TABLE_NAME & "@ODGTST"
            'ASCMAIN1.oraCmd.ExecuteNonQuery()
            'ASCMAIN1.oraCmd.CommandText = "INSERT INTO TST." & TABLE_NAME & "@ODGTST SELECT * FROM " & TABLE_NAME & ""
            'ASCMAIN1.oraCmd.ExecuteNonQuery()

        Next

        'If ASCMAIN1.DBS_COMPANY = "ODG" Then
        '    MsgBox("Now Restoring Vendor Format")
        '    Dim sql As String = "UPDATE ODG.ASTFFMT1@ODG" _
        '    & " SET JUSTIFY='R',FILL_CHAR='0',FIELD_LENGTH='6',ALPHA_NUMERIC='N',FIXED_LENGTH='1'" _
        '    & " WHERE COLUMN_NAME='VEND_CODE'"
        '    ASCDATA1.ExecuteSQL(sql)
        'End If

        ASCMAIN1.T.Commit()

    End Sub

    Sub Deploy_Assemblies()

        Dim p As New System.Diagnostics.ProcessStartInfo()
        'p.Verb = "Deploy"
        p.WindowStyle = ProcessWindowStyle.Normal ' .Hidden

        Dim i As Integer = InStr(My.Application.Info.DirectoryPath, "\ABS\bin")
        p.WorkingDirectory = Mid(My.Application.Info.DirectoryPath, 1, i - 1)

        ' p.WorkingDirectory = "C:\VS\" & ASCMAIN1.SOLUTION
        p.FileName = p.WorkingDirectory & "\deploy.bat"
        If My.Computer.FileSystem.FileExists(p.WorkingDirectory & "\deploy" & ASCMAIN1.DBS_COMPANY & ".bat") Then
            'p.FileName = p.WorkingDirectory & "\deployRGI.bat"
            p.FileName = p.WorkingDirectory & "\deploy" & ASCMAIN1.DBS_COMPANY & ".bat"
        End If
        p.UseShellExecute = True
        System.Diagnostics.Process.Start(p)
    End Sub
End Class
