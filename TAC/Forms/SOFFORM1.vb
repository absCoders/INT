Public Class SOFFORM1
    Public ORDR_FORM_CODE As String
    Public PRICE_BASE_DPCT As String

    Private Sub SOFFORM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTFORM2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_RETAIL_PRICE * (100 - " & CStr(PRICE_BASE_DPCT) & ") / 100 ORDR_UNIT_PRICE" _
                & " from SOTFORM2,ICTITEM1 where ICTITEM1.ITEM_CODE = SOTFORM2.ITEM_CODE" _
                & " and SOTFORM2.ORDR_FORM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTFORM2", "**", 0, False, "V")
            .Tables("SOTFORM2").Columns.Add("ORDR_QTY", GetType(System.Int64))
            .Tables("SOTFORM2").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
        End With

        grdSOTFORM2.DataSource = dst.Tables("SOTFORM2")

        Create_Summary(grdSOTFORM2, "ITEM_CODE", "Count")
        Create_Summary(grdSOTFORM2, New String() {"ORDR_QTY", "ORDR_AMT"})

        With grdSOTFORM2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With
        With grdSOTFORM2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "ORDR_QTY" Or gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        If ORDR_FORM_CODE <> "" Then
            Absx1.txtFor("ORDR_FORM_CODE").Text = ORDR_FORM_CODE
            Prepare_SOTFORM2()
        End If

    End Sub
    Overrides Sub Prepare_for_View_Lookup_Special _
    (ByVal ctl As Windows.Forms.Control, _
     ByVal COLUMN_NAME As String, _
     Optional ByRef sql_where As String = "", _
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "ORDR_FORM_CODE"
                sql_where = "ORDR_FORM_STATUS = 'A'"

        End Select
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ORDR_FORM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Prepare_SOTFORM2()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ORDR_FORM_CODE"
                Prepare_SOTFORM2()
        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        ORDR_FORM_CODE = Absx1.txtFor("ORDR_FORM_CODE").Text
        Dim row As DataRow = LookUp("SOTFORM1", ORDR_FORM_CODE)
        If row Is Nothing Then
            MsgBox("Invalid Form Code Specified", MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        Else
            If Val(dst.Tables("SOTFORM2").Compute("SUM(ORDR_QTY)", "") & "") = 0 Then
                MsgBox("No Qty's Ordered", MsgBoxStyle.OkOnly, "Cannot Update")
                Exit Sub
            End If
        End If

        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        ORDR_FORM_CODE = ""
        Me.Close()
    End Sub

    Sub Prepare_SOTFORM2()
        Dim ORDR_FORM_CODE As String = Absx1.txtFor("ORDR_FORM_CODE").Text
        Dim rowSOTFORM1 As DataRow = Nothing
        If ORDR_FORM_CODE <> "" Then rowSOTFORM1 = LookUp("SOTFORM1", ORDR_FORM_CODE)
        If rowSOTFORM1 IsNot Nothing Then
            Fill_Records("SOTFORM2", ORDR_FORM_CODE)
            Sort_grdColumns(grdSOTFORM2, "ORDR_FORM_LNO", True)
            grdSOTFORM2.Text = rowSOTFORM1.Item("ORDR_FORM_DESC") & ""
        Else
            grdSOTFORM2.Text = ""
            dst.Tables("SOTFORM2").Rows.Clear()
        End If
    End Sub
End Class