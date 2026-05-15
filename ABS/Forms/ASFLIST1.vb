Public Class ASFLIST1
    Dim LIST_CODE As String
    Dim COLUMN_NAME_list As String
    Public CODEs As Dictionary(Of String, String)
    Dim rowASTLIST1 As DataRow

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from ASTLIST1 where COLUMN_NAME = :PARM1"
            Create_TDA(.Tables.Add, "ASTLIST1", "**", 0, True, "V", 1)
            .Tables("ASTLIST1").Columns.Add("CODE_VALUES")

            ASCMAIN1.sql = "Select * from ASTLIST2 where LIST_CODE in (Select LIST_CODE from ASTLIST1 where COLUMN_NAME = :PARM1)"
            Create_TDA(.Tables.Add, "ASTLIST2", "**", 0, True, "V", 2)
            .Tables("ASTLIST2").Columns.Add("DESC_VALUE")

            Create_Relation("ASTLIST1", "ASTLIST2", "LIST_CODE")
        End With

        grdASTLIST1.DataSource = dst.Tables("ASTLIST1")
        grdASTLIST2.DataSource = dst.Tables("ASTLIST2")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdASTLIST1.DisplayLayout.Bands(0).Columns
            If gcol.Key = "LIST_DESC" Then
                gcol.Header.Appearance.BackColor = Color.Gold
                'gcol.Header.Appearance.BackColor2 = Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassTop20
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.Header.Appearance.BackColor = Color.Gold
                'gcol.Header.Appearance.BackColor2 = Color.White
                gcol.CellAppearance.BackColor = Color.Beige
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassTop20
            End If
        Next

        For Each gcol As UltraWinGrid.UltraGridColumn In grdASTLIST2.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Color.LightGreen
            'gcol.Header.Appearance.BackColor2 = Color.White
            gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassTop20
        Next


        With grdASTLIST1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.True
        End With

        With grdASTLIST2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            If EntryMode = "S" Then
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        Create_Summary(grdASTLIST1, "LIST_DESC", "Count")
        Create_Summary(grdASTLIST2, "CODE_VALUE", "Count")

        If EntryMode = "S" Then
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
            With rowASTLIST1
                .Item("LIST_CODE") = LIST_CODE
                .Item("LIST_DESC") = txtLIST_DESC.Text
                .Item("COLUMN_NAME") = COLUMN_NAME_list
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)

            For Each CODE_VALUE In CODEs.Keys
                Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
                rowASTLIST2.Item("LIST_CODE") = LIST_CODE
                rowASTLIST2.Item("CODE_VALUE") = CODE_VALUE
                rowASTLIST2.Item("DESC_VALUE") = CODEs(CODE_VALUE)
                dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
            Next
            txtLIST_DESC.Focus()
        Else
            Fill_Records("ASTLIST1", COLUMN_NAME_list)
            Fill_Records("ASTLIST2", COLUMN_NAME_list)

            For Each rowASTLIST1 As DataRow In dst.Tables("ASTLIST1").Rows
                rowASTLIST1.Item("CODE_VALUES") = Get_CODE_VALUES(rowASTLIST1)
            Next

            dst.Tables("ASTLIST1").AcceptChanges()
            Sort_grdColumns(grdASTLIST1, "LIST_DESC")
            Setup_ASTLIST2()
        End If
    End Sub

    Function Get_CODE_VALUES(ByVal rowASTLIST1 As DataRow) As String
        Dim CODE_VALUES As String = ""
        For Each rowASTLIST2 As DataRow In rowASTLIST1.GetChildRows("ASTLIST1_ASTLIST2")
            CODE_VALUES &= "," & rowASTLIST2.Item("CODE_VALUE")
        Next
        Return Mid(CODE_VALUES, 2)
    End Function

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub grdASTLIST1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTLIST1.AfterRowActivate
        Setup_ASTLIST2()
    End Sub

    Sub Setup_ASTLIST2()
        If grdASTLIST1.ActiveRow Is Nothing Then
            GRDASTLIST2.VISIBLE = False
        Else
            Dim LIST_CODE As String = grdASTLIST1.ActiveRow.Cells("LIST_CODE").Value
            Dim dvw As DataView = DirectCast(grdASTLIST2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "LIST_CODE = '" & LIST_CODE & "'"
            Sort_grdColumns(grdASTLIST2, "CODE_VALUE")
            grdASTLIST2.Visible = True
        End If

    End Sub

    Public Sub Maintain_Lists(ByVal COLUMN_NAME As String)

        lblLIST_DESC.Visible = False
        txtLIST_DESC.Visible = False

        COLUMN_NAME_list = COLUMN_NAME

        EntryMode = "M"
        Me.ShowDialog()
    End Sub

    Public Sub Save_List(ByVal COLUMN_NAME As String)
        SplitContainer2.Panel1Collapsed = True

        lblLIST_DESC.Visible = True
        txtLIST_DESC.Visible = True

        COLUMN_NAME_list = COLUMN_NAME

        EntryMode = "S"
        Me.ShowDialog()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        BeginTrans()

        If EntryMode = "S" Then
            rowASTLIST1.Item("LIST_DESC") = txtLIST_DESC.Text
        End If

        Update_Record_TDA("ASTLIST1")
        Update_Record_TDA("ASTLIST2")

        CommitTrans("List has been Saved")

        Me.Close()
    End Sub

    Private Sub grdASTLIST2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTLIST2.AfterRowsDeleted
        Dim LIST_CODE As String = grdASTLIST1.ActiveRow.Cells("LIST_CODE").Value
        Dim rowASTLIST1 As DataRow = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
        rowASTLIST1.Item("CODE_VALUES") = Get_CODE_VALUES(rowASTLIST1)
    End Sub

    Private Sub grdASTLIST1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTLIST1.ClickCellButton
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME_list)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdASTLIST1.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
            Using frmASFCODE1 As New ASFCODE1
                frmASFCODE1.ShowDialog()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    grdASTLIST1.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                    grdASTLIST1.UpdateData()
                End If
            End Using
        End If
    End Sub

End Class