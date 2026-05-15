Public Class ICTPRPT1
    Private Sub cmdRegen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegen.Click

        ASCMAIN1.sql = "Select PRICE_POINT_MIN, COUNT (*) from ICTPRPT1" _
        & " GROUP BY PRICE_POINT_MIN having COUNT (*) > 1"
        Dim row As DataRow = ASCDATA1.GetDataRow
        If row IsNot Nothing Then
            MsgBox(String.Format("There are Multiple Price Points codes with a Minimum Price Point of {0}", row.Item(0)), MsgBoxStyle.OkOnly, "Cannot Generate Function unless Price Points are Unique")
        Else
            Dim PLSQL As String = ""
            Dim PRICE_POINT_CODE As String = ""
            ASCMAIN1.sql = "Select * from ICTPRPT1"
            For Each rowICTPRPT1 As DataRow In ASCDATA1.GetDataTable.Select("", "PRICE_POINT_MIN")
                If PRICE_POINT_CODE <> "" Then
                    PLSQL &= " ELSIF NVL(ITEM_RETAIL_PRICE,0) < " _
                    & rowICTPRPT1.Item("PRICE_POINT_MIN") & " THEN" & vbLf _
                    & "  PP := '" & PRICE_POINT_CODE & "';" & vbLf
                End If
                PRICE_POINT_CODE = rowICTPRPT1.Item("PRICE_POINT_CODE")
            Next
            PLSQL &= " ELSE " & vbLf _
            & "  PP := '" & PRICE_POINT_CODE & "';" & vbLf

            PLSQL = "" _
            & "CREATE OR REPLACE FUNCTION PRICE_POINT" & vbLf _
            & "(ITEM_RETAIL_PRICE NUMBER )" & vbLf _
            & "RETURN VARCHAR2 AS" & vbLf _
            & "BEGIN" & vbLf _
            & "DECLARE PP VARCHAR2(10);" & vbLf _
            & "BEGIN" & vbLf _
            & Mid(PLSQL, 5) _
            & " END IF;" & vbLf _
            & "  RETURN PP;" & vbLf _
            & "END;" & vbLf _
            & "END;"

            Try
                ASCDATA1.ExecuteSQL(PLSQL)
                ASCDATA1.ExecuteSQL("Alter Function PRICE_POINT COMPILE")
                MsgBox("Function PRICE_POINT has been sucessfully Re-Created", MsgBoxStyle.OkOnly, "Verification")
            Catch ex As Exception
                ASCDATA1.ExecuteSQL(PLSQL)
                MsgBox("Error in the Creation of Function PRICE_POINT" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error has occurred")
            End Try
        End If
    End Sub
End Class