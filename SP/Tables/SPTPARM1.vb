Public Class SPTPARM1
      
    Private Sub SPTPARM1_Load(sender As Object, e As EventArgs) Handles Me.Load
        grpSPTACOM.Visible = (ASCMAIN1.CLIENT = "AHA")
        grpSPTDCOM.Visible = (ASCMAIN1.CLIENT = "INT")
        grpSPTMODL.Visible = (ASCMAIN1.CLIENT = "INT")
        grpSPTACOM.Top = grpSPTDCOM.Top
    End Sub
End Class