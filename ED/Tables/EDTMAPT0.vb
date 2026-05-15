
Public Class EDTMAPT0

    ' COPY FROM FUNCTION
    ' CHANGE KEY OF EDTMAPT1 TO INCLUDE TABLE LEVEL NOT TABLENAME
    ' CHG UPDATE_PRE TO REORDER TABLE LEVELS IN SEQ
    ' IF YOU DELETE A TABLE, DELETE THE MAPT2 RECORDS
    ' NO CHANGE TO TABLE NAME
    ' WHEN ADDING A TABLE CREATE COLUMN LIST
    ' LVL 1 TABLE NEEDS TO HAVE THE FOLLOWING COLUMNS: EDI_TRANSMIT_DATE, EDI_GRP_REF_NO, EDI_INT_REF_NO, EDI_MESSAGE_REF_NO
    ' PERHAPS ALL OF THESE EDI COLUMNS COULD BE IN EDTJRNL1 IF WE AGREE ON WHAT A ROW IN EDTJRNL1 REPRESENTS AND LINK THE 1ST LVL TABLE VIA THE EDI_JRNL_NO
    ' MUST HAVE A LEAD SEGMENT FOR EACH TABLE
    ' MAKE SURE TABLES ARE LVLd IN SEQ AND STARTING AT 1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Call Create_TDA(.Tables.Add, "EDTMAPT1", "*", 2)

            Call Create_TDA(.Tables.Add, "EDTMAPT2", "*", 2)
            .Tables("EDTMAPT2").Columns.Add("COLUMN_ID", GetType(System.Int32))

            Call Create_TDA(.Tables.Add, "EDTMAPTD", "*", 2)
        End With

        grdEDTMAPT1.DataSource = dst.Tables("EDTMAPT1")
        grdEDTMAPT2.DataSource = dst.Tables("EDTMAPT2")

        grdEDTMAPTD.DataSource = dst.Tables("EDTMAPTD")

        Call Create_Summary(grdEDTMAPT1, "TABLE_NAME", "Count")
        Call Create_Summary(grdEDTMAPT2, "COLUMN_NAME", "Count")
        Call Create_Summary(grdEDTMAPTD, "SEG_SEQ_NO", "Count")

        Sort_grdColumns(grdEDTMAPT2, "COLUMN_ID", True)
    End Sub


#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""

        'For i As Integer = dst.Tables("EDTMAPT2").Rows.Count - 1 To 0 Step -1
        '    Dim row As DataRow = dst.Tables("EDTMAPT2").Rows(i)
        'If row.Item("SEGMENT_ID") & "" = "" Then
        '    row.Delete()
        'End If
        'Next

        ASCDATA1.DeleteRows("EDTMAPT2", "SEGMENT_ID IS NULL OR SEGMENT_ID = ''")

        'Dim rows() As DataRow = dst.Tables("EDTMAPT2").Select _
        '("SEGMENT_ID IS NULL OR SEGMENT_ID = ''")
        'If rows.Length > 0 Then
        '    For Each row As DataRow In rows
        '        'dst.Tables("EDTMAPT2").Rows.Remove(row)
        '        row.Delete()
        '    Next
        'End If

        Call Update_Record_TDA("EDTMAPT1")
        Call Update_Record_TDA("EDTMAPT2")
        'Call Update_Record_TDA("EDTMAPTD")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            'rowASFBASE1.Item("TERM_CODE") = "AAA"
            'Absx1.CtlFor("TERM_CODE").Text = "BBB"
        End If

        dst.EnforceConstraints = False
        Call Fill_Records("EDTMAPT1", New String() {Absx1.txtFor("STANDARDS_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text})
        Call Fill_Records("EDTMAPT2", New String() {Absx1.txtFor("STANDARDS_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text})
        Call Fill_Records("EDTMAPTD", New String() {Absx1.txtFor("STANDARDS_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text})
        dst.EnforceConstraints = True

        For Each rowEDT852T1 As DataRow In dst.Tables("EDTMAPT1").Rows
            Get_All_Columns(rowEDT852T1.Item("TABLE_NAME"))
        Next


        ASCMAIN1.Add_Value_List(grdEDTMAPT1, "LEAD_SEGMENT", String.Format("SELECT SEGMENT_ID, SEGMENT_ID FROM EDTMAPTD WHERE STANDARDS_ID = '{0}' AND EDI_DOC_NO = '{1}'", Absx1.txtFor("STANDARDS_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text))
        ASCMAIN1.Add_Value_List(grdEDTMAPT2, "SEGMENT_ID", String.Format("SELECT SEGMENT_ID, SEGMENT_ID FROM EDTMAPTD WHERE STANDARDS_ID = '{0}' AND EDI_DOC_NO = '{1}'", Absx1.txtFor("STANDARDS_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text))

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            dst.EnforceConstraints = False
            dst.Tables("EDTMAPT1").Rows.Clear()
            dst.Tables("EDTMAPT2").Rows.Clear()
            dst.Tables("EDTMAPTD").Rows.Clear()
            dst.EnforceConstraints = False
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdEDTMAPT1.Enabled = tf
        grdEDTMAPT2.Enabled = tf
        grdEDTMAPTD.Enabled = tf
    End Sub

#End Region

    Private Sub grdEDTMAPT1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTMAPT1.AfterRowActivate
        Setup_EDTMAPT2()
    End Sub

    Sub Setup_EDTMAPT2()
        If SELECTION_NO = 0 Then Exit Sub

        If grdEDTMAPT1.ActiveRow Is Nothing OrElse grdEDTMAPT1.ActiveRow.IsAddRow Then
            grdEDTMAPT2.Visible = False
            Exit Sub
        End If

        Dim TABLE_NAME As String = grdEDTMAPT1.ActiveRow.Cells("TABLE_NAME").Text
        Dim dvw As DataView = DirectCast(grdEDTMAPT2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "TABLE_NAME = '" & TABLE_NAME & "'"
        dvw.Sort = "COLUMN_ID"

        If grdEDTMAPT2.Rows.Count > 0 Then
            grdEDTMAPT2.ActiveRow = grdEDTMAPT2.Rows(0)
        End If

        grdEDTMAPT2.Visible = True
        grdEDTMAPT2.Text = "Table Mappings for " & TABLE_NAME
    End Sub

    Private Sub grdEDTMAPT1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDTMAPT1.InitializeLayout

    End Sub

    Sub Get_All_Columns(ByVal TABLE_NAME As String)
        For Each row As DataRow In ASCDATA1.GetDataTable _
        ("Select COLUMN_NAME, COLUMN_ID from USER_TAB_COLUMNS where TABLE_NAME = '" & TABLE_NAME & "'").Select("", "COLUMN_ID")
            Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
            Dim rowEDTMAPT2 As DataRow = dst.Tables("EDTMAPT2").Rows.Find _
            (New String() {Absx1.txtFor("STANDARDS_ID").Text, _
                           Absx1.txtFor("EDI_DOC_NO").Text, _
                           TABLE_NAME, COLUMN_NAME})
            If rowEDTMAPT2 Is Nothing Then
                rowEDTMAPT2 = dst.Tables("EDTMAPT2").NewRow
                rowEDTMAPT2.Item("STANDARDS_ID") = Absx1.txtFor("STANDARDS_ID").Text
                rowEDTMAPT2.Item("EDI_DOC_NO") = Absx1.txtFor("EDI_DOC_NO").Text
                rowEDTMAPT2.Item("TABLE_NAME") = TABLE_NAME
                rowEDTMAPT2.Item("COLUMN_NAME") = COLUMN_NAME
                dst.Tables("EDTMAPT2").Rows.Add(rowEDTMAPT2)
                'ADD THE ROW
            End If
            rowEDTMAPT2.Item("COLUMN_ID") = row.Item("COLUMN_ID")
        Next
    End Sub

    Private Sub grdEDTMAPT2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDTMAPT2.InitializeRow
        If e.Row.Cells("COLUMN_ID").Text = "" Then
            e.Row.Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub
End Class