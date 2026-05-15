Imports Infragistics.Win.UltraWinGrid

Public Class WHTTPLP1

    Private isViewMode As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ICTITLP2", "*", 1)
            .Tables("ICTITLP2").Columns.Add("ITEM_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "ICTITLP4", "*", 1)
            .Tables("ICTITLP4").Columns.Add("ITEM_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "ICTITLP3", "*", 1)
            .Tables("ICTITLP3").Columns.Add("ITEM_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "ARTCULP2", "*", 1)
            .Tables("ARTCULP2").Columns.Add("CUST_STORE_NAME", GetType(System.String))

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1", "", True, "SELECT * FROM ICTWHSE1")

        End With

        grdICTITLP2.DataSource = dst.Tables("ICTITLP2")
        Create_Summary(grdICTITLP2, "ITEM_CODE", "Count")

        grdICTITLP3.DataSource = dst.Tables("ICTITLP3")
        Create_Summary(grdICTITLP3, "ITEM_CODE", "Count")

        grdICTITLP4.DataSource = dst.Tables("ICTITLP4")
        Create_Summary(grdICTITLP4, "ITEM_CODE", "Count")

        grdARTCULP2.DataSource = dst.Tables("ARTCULP2")
        Create_Summary(grdARTCULP2, "CUST_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        isViewMode = False

        Select Case eItemKey
            Case "New"

            Case "Edit"

            Case "View"
                isViewMode = True

            Case "Update"
                grdARTCULP2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdICTITLP2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdICTITLP4.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        For Each row As DataRow In dst.Tables("ICTITLP2").Select("", "", DataViewRowState.Added)
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("INIT_DATE") = DATETIME_STAMP
        Next

        For Each row As DataRow In dst.Tables("ICTITLP4").Select("", "", DataViewRowState.Added)
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("INIT_DATE") = DATETIME_STAMP
        Next

        For Each row As DataRow In dst.Tables("ARTCULP2").Select("", "", DataViewRowState.Added)
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("INIT_DATE") = DATETIME_STAMP
        Next

        Update_Record_TDA("ICTITLP2", "LP_CODE = '" & Absx1.txtFor("LP_CODE").Text & "'")
        Update_Record_TDA("ICTITLP4", "LP_CODE = '" & Absx1.txtFor("LP_CODE").Text & "'")
        Update_Record_TDA("ARTCULP2", "LP_CODE = '" & Absx1.txtFor("LP_CODE").Text & "'")

    End Sub

    Overrides Sub Show_Record_Special()
        ASCMAIN1.sql = $"SELECT ICTITLP2.*, ICTITEM1.ITEM_DESC
                            FROM ICTITLP2, ICTITEM1
                            WHERE LP_CODE = '{Absx1.txtFor("LP_CODE").Text}'
                            AND ICTITLP2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
        Fill_Records("ICTITLP2", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = $"SELECT ICTITLP3.*, ICTITEM1.ITEM_DESC
                            FROM ICTITLP3, ICTITEM1
                            WHERE LP_CODE = '{Absx1.txtFor("LP_CODE").Text}'
                            AND ICTITLP3.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
        Fill_Records("ICTITLP3", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = $"SELECT ICTITLP4.*, ICTITEM1.ITEM_DESC
                            FROM ICTITLP4, ICTITEM1
                            WHERE LP_CODE = '{Absx1.txtFor("LP_CODE").Text}'
                            AND ICTITLP4.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
        Fill_Records("ICTITLP4", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = $"SELECT ARTCULP2.*, ARTCUST2.CUST_STORE_NAME
                            FROM ARTCULP2, ARTCUST2
                            WHERE LP_CODE = '{Absx1.txtFor("LP_CODE").Text}'
                            AND ARTCULP2.CUST_CODE = ARTCUST2.CUST_CODE (+)
                            AND ARTCULP2.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
        Fill_Records("ARTCULP2", String.Empty, True, ASCMAIN1.sql)

        UltraTabControl1.Tabs("Transmit All Items").Visible = Absx1.txtFor("LP_CODE").Text <> "CLA"
        btnSendAllItems.Visible = isViewMode

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ICTITLP2").Rows.Clear()
            dst.Tables("ICTITLP3").Rows.Clear()
            dst.Tables("ICTITLP4").Rows.Clear()
            dst.Tables("ARTCULP2").Rows.Clear()
            EnforceConstraints(True)
        End If

        UltraTabControl1.Tabs("Transmit All Items").Visible = False
        btnSendAllItems.Visible = False
    End Sub

    Function Validate_Item(ITEM_CODE_z As String, grd As Infragistics.Win.UltraWinGrid.UltraGrid) As String
        Dim E As String = ""
        ITEM_CODE_z = ASCMAIN1.Format_Field(ITEM_CODE_z, "ITEM_CODE")
        Dim ITEM_CODE As String = ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE_z)

        If rowICTITEM1 Is Nothing Then
            E = "Item is Not on File" & vbCrLf
        Else
            If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                If grd.ActiveRow.IsAddRow Then
                    E = "Item Status is not Active" & vbCrLf
                End If
            End If
            If rowICTITEM1.Item("ITEM_UOM") & "" = "" Then
                E = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E <> "" And grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Item Code Entered is Invalid because ...")
        Else
            If E = "" Then
                ITEM_CODE = rowICTITEM1.Item(0)
            End If
        End If
        Return ITEM_CODE
    End Function

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim rowICTWHSE1 As DataRow

        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE_RFB"
                Absx1.txtFor("WHSE_DESC_RFB").Clear()
                rowICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE_RFB").Text)
                If rowICTWHSE1 IsNot Nothing Then
                    Absx1.txtFor("WHSE_DESC_RFB").Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                End If

            Case "WHSE_CODE_DST"
                Absx1.txtFor("WHSE_DESC_DST").Clear()
                rowICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE_DST").Text)
                If rowICTWHSE1 IsNot Nothing Then
                    Absx1.txtFor("WHSE_DESC_DST").Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                End If

            Case "WHSE_CODE_RTN"
                Absx1.txtFor("WHSE_DESC_RTN").Clear()
                rowICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE_RTN").Text)
                If rowICTWHSE1 IsNot Nothing Then
                    Absx1.txtFor("WHSE_DESC_RTN").Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                End If

            Case "WHSE_CODE_DISC"
                Absx1.txtFor("WHSE_DESC_DISC").Clear()
                rowICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(Absx1.txtFor("WHSE_CODE_DISC").Text)
                If rowICTWHSE1 IsNot Nothing Then
                    Absx1.txtFor("WHSE_DESC_DISC").Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                End If

        End Select

    End Sub

    Private Sub btnSendAllItems_Click(sender As Object, e As EventArgs) Handles btnSendAllItems.Click

        Try
            If Not isViewMode Then
                MessageBox.Show("Transfer only available in view mode.", "Send All Items", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If MessageBox.Show($"Do you want to send all Updated and New Items to 3PL {Absx1.txtFor("LP_CODE").Text}?", "Send All Items", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                Exit Sub
            End If

            Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "All", Absx1.txtFor("LP_CODE").Text)
            MessageBox.Show("Transfer Complete", "Send All Items", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Send All Items", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#Region "grdARTCULP2"

    Private Sub grdARTCULP2_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdARTCULP2.BeforeRowUpdate

        e.Row.Cells("LP_CODE").Value = Absx1.txtFor("LP_CODE").Text

        Dim CUST_CODE As String = (e.Row.Cells("CUST_CODE").Value & String.Empty).TRIM
        If CUST_CODE.Length = 0 Then
            MessageBox.Show("Customer is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim CUST_STORE_NO As String = (e.Row.Cells("CUST_STORE_NO").Value & String.Empty).TRIM
        If CUST_STORE_NO.Length = 0 Then
            MessageBox.Show("Store No is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
        If rowARTCUST2 Is Nothing Then
            MessageBox.Show("Invalid Customer / Ship To.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
        e.Row.Cells("CUST_STORE_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty

        Dim CUST_NO_3PL As String = (e.Row.Cells("CUST_NO_3PL").Value & String.Empty).ToString.Trim.ToUpper
        e.Row.Cells("CUST_NO_3PL").Value = CUST_NO_3PL
        If CUST_NO_3PL.Length = 0 Then
            MessageBox.Show("3PL Customer is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim CUST_STORE_NO_3PL As String = (e.Row.Cells("CUST_STORE_NO_3PL").Value & String.Empty).ToString.Trim.ToUpper
        e.Row.Cells("CUST_STORE_NO_3PL").Value = CUST_STORE_NO_3PL
        If CUST_STORE_NO_3PL.Length = 0 Then
            MessageBox.Show("3PL Store No is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ASCMAIN1.sql = $"LP_CODE = '{Absx1.txtFor("LP_CODE").Text}' 
                            AND CUST_NO_3PL = '{CUST_NO_3PL}' 
                            AND CUST_STORE_NO_3PL = '{CUST_STORE_NO_3PL}' 
                            AND CUST_CODE <> '{CUST_CODE}'
                            AND CUST_STORE_NO <> '{CUST_STORE_NO}'"
        If dst.Tables("ARTCULP2").Select(ASCMAIN1.sql).Length > 0 Then
            Dim dupShipTo As String = dst.Tables("ARTCULP2").Select(ASCMAIN1.sql)(0).Item("CUST_CODE") & String.Empty & "/" & dst.Tables("ARTCULP2").Select(ASCMAIN1.sql)(0).Item("CUST_STORE_NO") & String.Empty
            MessageBox.Show($"Customer / Ship To 3PL Code is already used by Item {dupShipTo}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ' All records are n the grid - do not need this. Grid contents replace what is in Oracle
        'ASCMAIN1.sql = $"Select * from ARTCULP2 
        '                    WHERE LP_CODE = :PARM1 
        '                    AND CUST_NO_3PL = :PARM2 
        '                    AND CUST_STORE_NO_3PL = :PARM3  
        '                    AND CUST_CODE <> :PARM4 
        '                    AND CUST_STORE_NO <> :PARM5"

        'Dim rowDuplicate As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVVVV", New Object() {Absx1.txtFor("LP_CODE").Text, CUST_NO_3PL, CUST_STORE_NO_3PL, CUST_CODE, CUST_STORE_NO})
        'If rowDuplicate IsNot Nothing Then
        '    MessageBox.Show($"3PL Code is already used by Ship To {rowDuplicate.Item("CUST_CODE")} / {rowDuplicate.Item("CUST_STORE_NO")}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    e.Cancel = True
        '    Exit Sub
        'End If

    End Sub

    Private Sub grdARTCULP2_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTCULP2.ClickCellButton
        If grdARTCULP2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key
            Case "CUST_CODE"

            Case "CUST_STORE_NO"
                Dim CUST_CODE As String = (grdARTCULP2.ActiveRow.Cells("CUST_CODE").Value & String.Empty)
                CUST_CODE = CUST_CODE.Replace("'", "")
                grdARTCULP2.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
                sql_where = $"CUST_CODE = '{CUST_CODE}'"

            Case Else
                Exit Sub

        End Select

        grdClickCellButton(grdARTCULP2, sql_where, False)
    End Sub

#End Region

#Region "grdICTITLP2"

    Private Sub grdICTITLP2_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdICTITLP2.BeforeRowUpdate

        e.Row.Cells("LP_CODE").Value = Absx1.txtFor("LP_CODE").Text

        Dim ITEM_CODE As String = (e.Row.Cells("ITEM_CODE").Value & String.Empty).TRIM
        If ITEM_CODE.Length = 0 Then
            MessageBox.Show("Item Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 Is Nothing Then
            MessageBox.Show("Invalid Item Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
        e.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & String.Empty

        Dim ITEM_CODE_3PL As String = (e.Row.Cells("ITEM_CODE_3PL").Value & String.Empty).ToString.Trim.ToUpper
        e.Row.Cells("ITEM_CODE_3PL").Value = ITEM_CODE_3PL

        If ITEM_CODE_3PL.Length = 0 Then
            MessageBox.Show("3PL Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ASCMAIN1.sql = $"LP_CODE = '{Absx1.txtFor("LP_CODE").Text}' AND ITEM_CODE_3PL = '{ITEM_CODE_3PL}' AND ITEM_CODE <> '{ITEM_CODE}'"
        If dst.Tables("ICTITLP2").Select(ASCMAIN1.sql).Length > 0 Then
            Dim ITEM_CODE_DUP As String = dst.Tables("ICTITLP2").Select(ASCMAIN1.sql)(0).Item("ITEM_CODE") & String.Empty
            MessageBox.Show($"3PL Code is already used by Item {ITEM_CODE_DUP}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ' All records are n the grid - do not need this. Grid contents replace what is in Oracle
        'ASCMAIN1.sql = "Select * from ICTITLP2 WHERE LP_CODE = :PARM1 AND ITEM_CODE_3PL = :PARM2 AND ITEM_CODE <> :PARM3"
        'Dim rowDuplicate As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {Absx1.txtFor("LP_CODE").Text, ITEM_CODE_3PL, ITEM_CODE})
        'If rowDuplicate IsNot Nothing Then
        '    MessageBox.Show($"3PL Code is already used by Item {rowDuplicate.Item("ITEM_CODE")}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    e.Cancel = True
        '    Exit Sub
        'End If

    End Sub

    Private Sub grdICTITLP2_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdICTITLP2.ClickCellButton
        If grdICTITLP2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

            Case Else
                Exit Sub
        End Select

        grdClickCellButton(grdICTITLP2, sql_where, False)
    End Sub

    Private Sub grdICTITLP2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITLP2.AfterCellUpdate
        Try
            Select Case e.Cell.Column.Key
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "", grdICTITLP2) ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)
                    If ITEM_CODE <> "" Then
                        e.Cell.Row.Cells("ITEM_CODE_3PL").Value = ITEM_CODE
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub grdICTITLP2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTITLP2.BeforeExitEditMode
        Try
            If grdICTITLP2.ActiveCell IsNot Nothing Then
                With grdICTITLP2.ActiveCell
                    If .EditorResolved.IsValid Then
                        Select Case .Column.Key
                            Case "ITEM_CODE"
                                .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End Select
                    End If

                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

#End Region

#Region "grdICTITLP4"

    Private Sub grdICTITLP4_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdICTITLP4.BeforeRowUpdate

        e.Row.Cells("LP_CODE").Value = Absx1.txtFor("LP_CODE").Text

        Dim ITEM_CODE As String = (e.Row.Cells("ITEM_CODE").Value & String.Empty).TRIM
        If ITEM_CODE.Length = 0 Then
            MessageBox.Show("Item Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 Is Nothing Then
            MessageBox.Show("Invalid Item Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
        e.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & String.Empty

        Dim ITEM_CODE_3PL As String = (e.Row.Cells("ITEM_CODE_3PL").Value & String.Empty).ToString.Trim.ToUpper
        e.Row.Cells("ITEM_CODE_3PL").Value = ITEM_CODE_3PL

        If ITEM_CODE_3PL.Length = 0 Then
            MessageBox.Show("3PL Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ASCMAIN1.sql = $"LP_CODE = '{Absx1.txtFor("LP_CODE").Text}' AND ITEM_CODE_3PL = '{ITEM_CODE_3PL}' AND ITEM_CODE <> '{ITEM_CODE}'"
        If dst.Tables("ICTITLP4").Select(ASCMAIN1.sql).Length > 0 Then
            Dim ITEM_CODE_DUP As String = dst.Tables("ICTITLP4").Select(ASCMAIN1.sql)(0).Item("ITEM_CODE") & String.Empty
            MessageBox.Show($"3PL Code is already used by Item {ITEM_CODE_DUP}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        ' All records are n the grid - do not need this. Grid contents replace what is in Oracle
        'ASCMAIN1.sql = "Select * from ICTITLP4 WHERE LP_CODE = :PARM1 AND ITEM_CODE_3PL = :PARM2 AND ITEM_CODE <> :PARM3"
        'Dim rowDuplicate As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {Absx1.txtFor("LP_CODE").Text, ITEM_CODE_3PL, ITEM_CODE})
        'If rowDuplicate IsNot Nothing Then
        '    MessageBox.Show($"3PL Code is already used by Item {rowDuplicate.Item("ITEM_CODE")}.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    e.Cancel = True
        '    Exit Sub
        'End If

    End Sub

    Private Sub grdICTITLP4_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdICTITLP4.ClickCellButton
        If grdICTITLP4.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

            Case Else
                Exit Sub
        End Select

        grdClickCellButton(grdICTITLP4, sql_where, False)
    End Sub

    Private Sub grdICTITLP4_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITLP4.AfterCellUpdate
        Try
            Select Case e.Cell.Column.Key
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "", grdICTITLP4)
                    If ITEM_CODE <> "" Then
                        e.Cell.Row.Cells("ITEM_CODE_3PL").Value = ITEM_CODE
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub grdICTITLP4_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTITLP4.BeforeExitEditMode
        Try
            If grdICTITLP4.ActiveCell IsNot Nothing Then
                With grdICTITLP4.ActiveCell
                    If .EditorResolved.IsValid Then
                        Select Case .Column.Key
                            Case "ITEM_CODE"
                                .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End Select
                    End If

                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

#End Region

End Class