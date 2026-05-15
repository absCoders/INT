Public Class WHCCWXI1
    ' Pull CoWorx file in from sftp site, archive and load into Oracle

    Inherits WHC000O1

    Dim TMP_COWORX_RSC As String
    Dim COWORX_RSC As String

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHCCWXI1"

        Create_Work_Table()

        With dst
            ASCMAIN1.sql = "Select * from " & TMP_COWORX_RSC
            Create_TDA(.Tables.Add, "COWORX_RSC", "**", 0, False)

        End With

        Main_Process()
    End Sub

    Public Sub Main_Process()

        Fill_Records("CSMST")
        Update_Record()

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Using sr As New System.IO.StreamReader(COWORX_RSC)
            Dim data = sr.ReadToEnd

            For Each line As String In Split(data, vbCrLf)
                R += 1
                Dim RECORD As String = ""

            Next
        End Using
    End Sub

    Overrides Sub Update_Archive()
        MyBase.Update_Archive()

        'My.Computer.FileSystem.CopyFile(CSMST, sftp_folder & CSMST)
        'My.Computer.FileSystem.MoveFile(CSMST, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & CSMST)

        'ASCDATA1.ExecuteSQL("Delete from CONV.CFG_CSMST")
        'ASCDATA1.ExecuteSQL("Insert into CONV.CFG_CSMST Select * from " & TMP_COWORX_RSC)
    End Sub

    Sub Create_Work_Table()

        ASCMAIN1.sql = "Select * from CONV.CFG_CSMST where ROWNUM < 1"
        TMP_COWORX_RSC = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "" _
            & "Select" & vbCrLf
        ASCDATA1.ExecuteSQL("Insert into " & TMP_COWORX_RSC & " " & ASCMAIN1.sql)
    End Sub

    Sub Poll_sftp_site_for_files()
        Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", "COWORX")

    End Sub
End Class