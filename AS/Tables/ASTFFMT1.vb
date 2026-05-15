Public Class ASTFFMT1
    Overrides Sub Proceed_Update_Special_Post()
        ASCMAIN1.tblASTFFMT1.Clear()
        ASCMAIN1.tblASTFFMT1 = ASCDATA1.GetDataTable("*", "ASTFFMT1")
    End Sub
End Class