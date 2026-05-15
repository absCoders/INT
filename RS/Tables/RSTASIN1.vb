Public Class RSTASIN1
    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "EAN_UPC"
                If EntryMode <> "" Then
                    Check_Item_Link()
                End If

        End Select
    End Sub
    Overrides Sub Show_Record_Special()
        Check_Item_Link()
    End Sub

    Private Sub Check_Item_Link()
        If Absx1.txtFor("EAN_UPC").Text <> "" Then
            Dim EAN_UPC As String = Absx1.txtFor("EAN_UPC").Text & ""

            Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow($"Select * from ICTITEM1 where ITEM_EAN_CODE = '{EAN_UPC}'")

            If IsNothing(rowICTITEM1) Then
                rowICTITEM1 = ASCDATA1.GetDataRow($"Select * from ICTITEM1 where ITEM_UPC_CODE = '{EAN_UPC}'")
            End If

            If rowICTITEM1 IsNot Nothing Then
                Absx1.txtFor("ITEM_CODE").Text = rowICTITEM1.Item("ITEM_CODE") & ""
                Absx1.txtFor("ITEM_DESC").Text = rowICTITEM1.Item("ITEM_DESC") & ""
            Else
                Absx1.txtFor("ITEM_CODE").Text = ""
                Absx1.txtFor("ITEM_DESC").Text = ""
            End If
        End If

    End Sub

    Private Sub btnSpecial_Click(sender As Object, e As EventArgs) Handles btnSpecial.Click

        If OK_To_Requeue() Then
            Dim sqlWBTRPTS1 As String = "Update WBTRPTS1 SET RUN_NOW = :PARM1, NEXT_RUN_DATE = SYSDATE WHERE RPT_SCHED_NO = :PARM2"
            ASCDATA1.ExecuteSQL(sqlWBTRPTS1, "VV", New String() {"1", "0000000001"})

            MsgBox("Amazon Retail Sales Import has been requeued for execution", MsgBoxStyle.Information & vbOKOnly, "Success!")
        End If

    End Sub

    Function OK_To_Requeue() As Boolean
        Dim sqlCheckForRecentCompleted As String = "Select * from WBTRPTH1 where RPT_SCHED_NO = '0000000001' AND STATUS = 'C' AND STATUS_DATE_C > SYSDATE - 6"
        Dim rowWBTRPTH1_c As DataRow = ASCDATA1.GetDataRow(sqlCheckForRecentCompleted)
        If Not IsNothing(rowWBTRPTH1_c) Then
            Dim STATUS_DATE_C As Date = rowWBTRPTH1_c("STATUS_DATE_C")
            MsgBox($"The last successful import was comnpleted on {STATUS_DATE_C}.", vbOKOnly, "Cannot Proceed")
            Return False
        End If
        Dim sqlCheckForQueuedAndPending As String = "Select * from WBTRPTH1 where RPT_SCHED_NO = '0000000001' AND STATUS in ('Q','P')"
        Dim rowWBTRPTH1_q As DataRow = ASCDATA1.GetDataRow(sqlCheckForQueuedAndPending)
        If Not IsNothing(rowWBTRPTH1_q) Then
            MsgBox($"Report already queued for execution.", vbOKOnly, "Cannot Proceed")
            Return False
        End If
        Return True
    End Function
End Class