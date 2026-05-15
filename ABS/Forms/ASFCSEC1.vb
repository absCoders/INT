Public Class ASFCSEC1
    Dim FORM_NAME_current As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.ActiveForm Is Nothing OrElse ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT = "" Then
            Me.Tag = "X"
            Exit Sub
        End If

        With dst

            Create_TDA(.Tables.Add, "ASTCSEC1", "*", 1)

            With .Tables.Add("ASTCSECX")
                .Columns.Add("GROUP_KEY")
                .Columns.Add("ITEM_KEY")
                .Columns.Add("SECURITY_CODES")
            End With

            Create_TDA(.Tables.Add, "ASTSECM1", "*", 0)

            Create_Relation("ASTSECM1", "ASTCSEC1", "SECURITY_CODE")


            .Tables("ASTCSEC1").Columns.Add("SECURITY_DESC", GetType(System.String), "PARENT.SECURITY_DESC")
            .Tables("ASTCSEC1").Columns.Add("SELECTED")
        End With

        grdASTCSECX.DataSource = dst.Tables("ASTCSECX")
        grdASTCSECX.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdASTCSECX.DisplayLayout.Bands(0).SortedColumns.Add("GROUP_KEY", False, True)

        FORM_NAME_current = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT

        Fill_Records("ASTSECM1")
        Dim SECURITY_CODEs As New List(Of String)
        For Each rowASTSECM1 As DataRow In dst.Tables("ASTSECM1").Rows
            SECURITY_CODEs.Add(rowASTSECM1.Item("SECURITY_CODE"))
        Next

        Fill_Records("ASTCSEC1", FORM_NAME_current)
        For Each rowASTCSEC1 As DataRow In dst.Tables("ASTCSEC1").Rows
            rowASTCSEC1.Item("SELECTED") = "1"
        Next

        For Each G As UltraWinExplorerBar.UltraExplorerBarGroup In ASCMAIN1.ActiveForm.UltraExplorerBar1.Groups
            For Each I As UltraWinExplorerBar.UltraExplorerBarItem In G.Items
                Dim rowASTCSECX As DataRow = dst.Tables("ASTCSECX").NewRow
                'rowASTCSECX.Item("FORM_NAME") = FORM_NAME_current
                rowASTCSECX.Item("GROUP_KEY") = G.Key
                rowASTCSECX.Item("ITEM_KEY") = I.Key
                rowASTCSECX.Item("SECURITY_CODES") = Get_SECURITY_CODES(G.Key, I.Key)
                dst.Tables("ASTCSECX").Rows.Add(rowASTCSECX)

                For Each SECURITY_CODE As String In SECURITY_CODEs
                    Dim rowASTCSEC1 As DataRow = dst.Tables("ASTCSEC1").Rows.Find _
                        (New String() {FORM_NAME_current, G.Key, I.Key, SECURITY_CODE})
                    If rowASTCSEC1 Is Nothing Then
                        rowASTCSEC1 = dst.Tables("ASTCSEC1").NewRow
                        rowASTCSEC1.Item("FORM_NAME") = FORM_NAME_current
                        rowASTCSEC1.Item("GROUP_KEY") = G.Key
                        rowASTCSEC1.Item("ITEM_KEY") = I.Key
                        rowASTCSEC1.Item("SECURITY_CODE") = SECURITY_CODE
                        dst.Tables("ASTCSEC1").Rows.Add(rowASTCSEC1)
                    End If
                Next
            Next
        Next

        grdASTCSEC1.DataSource = dst.Tables("ASTCSEC1")
        Sort_grdColumns(grdASTCSEC1, "SECURITY_CODE")

        grdASTCSECX.Rows.ExpandAll(True)

        Me.Text = Me.Text & " - " & ASCMAIN1.ActiveForm.MENU_ITEM_DESC
    End Sub

    Private Sub grdASTCSECX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTCSECX.AfterRowActivate

        If Not grdASTCSECX.ActiveRow.IsDataRow Then
            grdASTCSEC1.Visible = False
        Else
            grdASTCSEC1.Visible = True

            Dim GROUP_KEY As String = grdASTCSECX.ActiveRow.Cells("GROUP_KEY").Text
            Dim ITEM_KEY As String = grdASTCSECX.ActiveRow.Cells("ITEM_KEY").Text

            Dim dvw As DataView = DirectCast(grdASTCSEC1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "GROUP_KEY = '" & GROUP_KEY & "' AND ITEM_KEY = '" & ITEM_KEY & "'"
        End If
    End Sub

    Private Sub grdASTCSECX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTCSECX.InitializeLayout

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Delete_Rows("ASTCSEC1", "ISNULL(SELECTED,'0')<>'1'")
        Dim sql_Delete As String = "FORM_NAME = '" & FORM_NAME_current & "'"
        Update_Record_TDA("ASTCSEC1", sql_Delete)
        Me.Close()
    End Sub

    Function Get_SECURITY_CODES(ByVal GROUP_KEY As String, ByVal ITEM_KEY As String) As String

        Dim SECURITY_CODEs As String = ""

        For Each rowASTCSEC1 As DataRow In dst.Tables("ASTCSEC1") _
        .Select("GROUP_KEY = '" & GROUP_KEY & "' AND ITEM_KEY = '" & ITEM_KEY & "' and SELECTED = '1'", _
                "SECURITY_CODE")

            SECURITY_CODEs &= "," & rowASTCSEC1.Item("SECURITY_CODE")
        Next

        Return Mid(SECURITY_CODEs, 2)
    End Function

    Private Sub grdASTCSEC1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTCSEC1.AfterRowUpdate
        Dim GROUP_KEY As String = grdASTCSECX.ActiveRow.Cells("GROUP_KEY").Text
        Dim ITEM_KEY As String = grdASTCSECX.ActiveRow.Cells("ITEM_KEY").Text
        grdASTCSECX.ActiveRow.Cells("SECURITY_CODES").Value = Get_SECURITY_CODES(GROUP_KEY, ITEM_KEY)
        grdASTCSECX.UpdateData()
    End Sub

    Private Sub grdASTCSEC1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTCSEC1.InitializeLayout

    End Sub

    Private Sub ASFCSEC1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Me.Tag = "X" Then
            Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
        End If
    End Sub
End Class