Public Class ASFXHST1

    Public RPT As Boolean

    Private Sub ASFXHST1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ASCMAIN1.sql = "Select * from ASTOPST1 where MENU_ITEM_OBJECT = '" & ASFMAIN1.UltraStatusBar1.Panels("MENU_ITEM_OBJECT").Text & "'"

        Dim tbl As DataTable = ASCDATA1.GetDataTable("", "ASTOPST1")
        grdASTOPST1.DataSource = tbl
        Me.Text = "Execution History for " & ASCMAIN1.ActiveForm.Text

    End Sub

    Private Sub grdASTOPST1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTOPST1.ClickCellButton

        ASCMAIN1.sql = grdASTOPST1.ActiveRow.Cells("SET_ID").Text
        Me.Close()

    End Sub

    Private Sub grdASTOPST1_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTOPST1.InitializeLayout
        Call ASCMAIN1.grdInitializeLayout(grdASTOPST1)
        With grdASTOPST1.DisplayLayout.Bands(0)

            .SortedColumns.Clear()
            .SortedColumns.Add("INIT_DATE", True)

            .Columns("PROCEED_BEGIN").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
            .Columns("PROCEED_END").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
            .Columns("UPDATE_BEGIN").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
            .Columns("UPDATE_END").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
            .Columns("SET_ID").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
            .Columns("SET_DESC").Hidden = (ASCMAIN1.ActiveForm.MENU_ITEM_TYPE <> "R")
        End With

    End Sub

    Private Sub ASFXHST1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If RPT Then
            With grdASTOPST1.DisplayLayout.Bands(0)
                .Columns("SET_ID").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
            End With
            ASCMAIN1.sql = ""
        End If
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
End Class