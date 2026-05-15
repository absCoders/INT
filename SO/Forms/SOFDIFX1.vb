Public Class SOFDIFX1
    Private ediType As String
    Private EDI_DOC_SEQ_NO As String

    Private ediTableName As String

    Public Sub New(byval ediType As String,byval EDI_DOC_SEQ_NO As string)
        MyBase.New()

        Me.ediType = ediType
        Me.EDI_DOC_SEQ_NO = EDI_DOC_SEQ_NO
        Me.ediTableName = $"EDT{ediType}T2"

        AUDIT.Add(Me.ediTableName,"*")
        InitializeComponent()
    End SUb

    Private Sub SOFDIFX1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            Create_TDA(.Tables.Add, ediTableName, "*", 1, True)

        End With

        grdEDTXXXT2.DataSource = dst.Tables(ediTableName)
        Create_Summary(grdEDTXXXT2, "EDI_DTL_SEQ", "Count")

        With grdEDTXXXT2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.True
        End With

        With grdEDTXXXT2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "EDI_EAN" Or gcol.Key = "EDI_UPC" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Sort_grdColumns(grdEDTXXXT2,"EDI_DTL_SEQ")
        Fill_Records(ediTableName,{EDI_DOC_SEQ_NO})
    End Sub



    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click

        Dim eMsg As String = ""
        'Verify the item ean/upcs are all valid
        For Each row As DataRow In dst.Tables(ediTableName).Select()
                eMsg &= ValidateItem(row.Item("EDI_DTL_SEQ"), row.Item("EDI_EAN") & "", row.Item("EDI_UPC") & "")
        Next

        If eMsg = "" Then
            'Record change -- audit or event?
            Update_Record_TDA(ediTableName)
            MsgBox("Update Successful")
            Me.Close()
        Else
            MsgBox(eMsg)
        End If
    
    End Sub

    Private Function ValidateItem(byval EDI_DTL_SEQ As Integer, ByVal ITEM_EAN_CODE As String, ByVal ITEM_UPC_CODE As String) As String
        Dim rowICTITEM1 As DataRow
        If ITEM_EAN_CODE <> "" Then
            ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_EAN_CODE = '" & ITEM_EAN_CODE & "'"
            rowICTITEM1 = ASCDATA1.GetDataRow
        ElseIf ITEM_UPC_CODE <> "" Then
            ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_UPC_CODE = '" & ITEM_UPC_CODE & "'"
            rowICTITEM1 = ASCDATA1.GetDataRow
        End If

        If rowICTITEM1 Is Nothing Then
            Dim errorText As String = ""
            If String.IsNullOrEmpty(ITEM_EAN_CODE) AndAlso String.IsNullOrEmpty(ITEM_UPC_CODE) Then
                errorText = "Missing item EAN/UPC"
            ElseIf Not String.IsNullOrEmpty(ITEM_EAN_CODE) Then
                errorText = $"Invalid item EAN ({ITEM_EAN_CODE})"
            ElseIf Not String.IsNullOrEmpty(ITEM_UPC_CODE) Then
                errorText = $"Invalid item UPC ({ITEM_UPC_CODE})"
            End If

            Return $"{errorText} on EDI Line {EDI_DTL_SEQ}" & vbCrLf
        End If
        Return ""
    End Function

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub grdEDTXXXT2_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDTXXXT2.BeforeRowsDeleted
        e.DisplayPromptMsg = False
    End Sub
End Class