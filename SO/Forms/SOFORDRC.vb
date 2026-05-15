Public Class SOFORDRC
    Public ORDR_NOs As New List(Of String)
    Public CUST_CODE As String

    Private Sub SOFBINV1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            Create_TDA(.Tables.Add, "SOTORDR1", "*")

            ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")
        End With

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")

        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

        For Each ORDR_NO As String In ORDR_NOs
            Fill_Record("SOTORDR1", ORDR_NO, False, False)
        Next

        Absx1.txtFor("CUST_CODE").Text = Me.CUST_CODE
        Set_Date()

        If dst.Tables("SOTORDR1").Columns.Contains("WHSE_CODE") Then
            dst.Tables("SOTORDR1").Columns("WHSE_CODE").ReadOnly = False
        End If

        With grdSOTORDR1.DisplayLayout
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            .Bands(0).Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
        End With

        With grdSOTORDR1.DisplayLayout.Bands(0).Columns("WHSE_CODE")
            .Hidden = False
            .CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            .CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
            .Nullable = Infragistics.Win.UltraWinGrid.Nullable.Nothing
            .TabStop = True
        End With

        whseFixed.ReadOnly = False
        whseFixed.Enabled = True
        whseFixed.TabStop = True



    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then

                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"

        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        EMsg = ""
 
        If dst.Tables("SOTORDR1").Select("", "", DataViewRowState.ModifiedCurrent).Length = 0 Then
            EMsg &= vbCr & "No Orders were Modified - please Cancel"
        End If

        Dim ORDR_GROUP_NOs As New List(Of String)

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")

            Dim ORDR_GROUP_NO As String = rowSOTORDR1("ORDR_GROUP_NO")
            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)

            Dim ORDR_SHIP_DATE As Date = rowSOTORDR1("ORDR_SHIP_DATE")
            Dim ORDR_CANCEL_DATE As Date = rowSOTORDR1("ORDR_CANCEL_DATE")
            Dim ORDR_ALLO_DATE As Date = IIf(rowSOTORDR1("ORDR_ALLO_DATE") & "" = "", Nothing, rowSOTORDR1("ORDR_ALLO_DATE"))
            Dim ORDR_ARRIVAL_DATE As Date = IIf(rowSOTORDR1("ORDR_ARRIVAL_DATE") & "" = "", Nothing, rowSOTORDR1("ORDR_ARRIVAL_DATE"))
            Dim WHSE_CODE As String = rowSOTORDR1("WHSE_CODE") & ""


            'If Format(ORDR_ARRIVAL_DATE, "MM/dd/yyyy") = "01/01/0001" Then ORDR_ARRIVAL_DATE = Nothing
            'If Format(ORDR_ALLO_DATE, "MM/dd/yyyy") = "01/01/0001" Then ORDR_ALLO_DATE = Nothing

            If Format(ORDR_SHIP_DATE, "yyyyMMdd") > Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                EMsg &= vbCr & "Ship Date Cannot be Later than Cancel Date"
            Else
                If Format(ORDR_ALLO_DATE, "MM/dd/yyyy") <> "01/01/0001" Then ' If ORDR_ALLO_DATE & "" <> "" Then
                    Dim DYS As Int64 = ORDR_SHIP_DATE.Subtract(ORDR_ALLO_DATE).Days
                    If System.Math.Abs(DYS) > 45 Then
                        EMsg &= vbCr & "Allocation Override Date Cannot be > 45 days away from Ship Date"
                    End If
                End If
            End If

            If Format(ORDR_ARRIVAL_DATE, "MM/dd/yyyy") <> "01/01/0001" Then
                If Format(ORDR_SHIP_DATE, "yyyyMMdd") > Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") Then
                    EMsg &= vbCr & "Arrival date must not be prior to Ship-By Date"
                End If
            End If

            If WHSE_CODE = "" Then
                EMsg &= vbCr & $"Order {rowSOTORDR1("ORDR_NO")} has blank WHSE_CODE."
            Else
                Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow($"SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = '{WHSE_CODE}'")
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & $"Order {rowSOTORDR1("ORDR_NO")} has invalid WHSE_CODE '{WHSE_CODE}' (not in ICTWHSE1)."
                End If
            End If

        Next


        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        Dim whseChanges As New List(Of Tuple(Of String, String, String, String))
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & ""
            Dim oldWhse As String = rowSOTORDR1.Item("WHSE_CODE", DataRowVersion.Original) & ""
            Dim newWhse As String = rowSOTORDR1.Item("WHSE_CODE", DataRowVersion.Current) & ""

            If (oldWhse & "") <> "" AndAlso (newWhse & "") <> "" AndAlso
           Not oldWhse.Equals(newWhse, StringComparison.OrdinalIgnoreCase) Then

                Dim ORDR_GROUP_NO_orig As String = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
                whseChanges.Add(Tuple.Create(ORDR_NO, oldWhse, newWhse, ORDR_GROUP_NO_orig))
            End If
        Next

        If whseChanges.Count > 0 Then

            Dim msgMissing As String = ""

            For Each tup As Tuple(Of String, String, String, String) In whseChanges
                Dim ORDR_NO As String = tup.Item1 & ""
                Dim WHSE_NEW As String = tup.Item3 & ""

                Dim missing As List(Of String) = Get_Missing_ICTWHSE2_Items(ORDR_NO, WHSE_NEW)
                If missing IsNot Nothing AndAlso missing.Count > 0 Then

                    Dim showN As Integer = Math.Min(25, missing.Count)
                    Dim showItems As String = String.Join(", ", missing.Take(showN))

                    msgMissing &= vbCr & $"Order {ORDR_NO}: cannot change WHSE to {WHSE_NEW} because these items are not set up in ICTWHSE2 for {WHSE_NEW}: {showItems}"

                    If missing.Count > showN Then
                        msgMissing &= $" ... (+{missing.Count - showN} more)"
                    End If
                End If
            Next

            If msgMissing <> "" Then
                MsgBox(Mid(msgMissing, 2), MsgBoxStyle.OkOnly, "Cannot Update")
                Exit Sub
            End If
        End If

        BeginTrans()

        dst.Tables("SOTORDXR").Rows.Clear()

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "Change", "Multi-Order Edit")

            ASCMAIN1.sql = "Select Max (REV_NO) from SOTORDXR where ORDR_NO = '" & ORDR_NO & "'"
            Dim REV_NO As Int32 = Val(ASCDATA1.GetDataValue & "") + 1
            Dim REV_LNO As Int32 = 0

            For Each COLUMN_NAME As String In New String() _
                {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ALLO_DATE", "ORDR_ARRIVAL_DATE", "WHSE_CODE"}

                If rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" _
                <> rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Current) & "" Then

                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                    With rowSOTORDXR
                        .Item("REV_NO") = REV_NO
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = 0
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                        .Item("NEW_VALUE") = rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Current)
                        .Item("EMODE") = EntryMode
                    End With
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)

                End If
            Next
        Next


        Update_Record_TDA("SOTORDXR")

        Dim grpTotal As New Dictionary(Of String, Int32)
        For Each r As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim g As String = r.Item("ORDR_GROUP_NO") & ""
            If g = "" Then Continue For

            If grpTotal.ContainsKey(g) Then
                grpTotal(g) += 1
            Else
                grpTotal.Add(g, 1)
            End If
        Next

        Dim grpChanged As New Dictionary(Of String, Int32)
        Dim grpNewWhse As New Dictionary(Of String, String)

        For Each TUP As Tuple(Of String, String, String, String) In whseChanges
            Dim ORDR_GROUP_NO_orig As String = TUP.Item4 & ""
            If ORDR_GROUP_NO_orig = "" Then Continue For

            If grpChanged.ContainsKey(ORDR_GROUP_NO_orig) Then
                grpChanged(ORDR_GROUP_NO_orig) += 1
            Else
                grpChanged.Add(ORDR_GROUP_NO_orig, 1)
            End If

            Dim WHSE_CODE_NEW As String = (TUP.Item3 & "").Trim()
            If WHSE_CODE_NEW = "" Then Continue For

            If Not grpNewWhse.ContainsKey(ORDR_GROUP_NO_orig) Then
                grpNewWhse.Add(ORDR_GROUP_NO_orig, WHSE_CODE_NEW)
            Else
                If (grpNewWhse(ORDR_GROUP_NO_orig) & "") <> "" AndAlso
                   Not grpNewWhse(ORDR_GROUP_NO_orig).Equals(WHSE_CODE_NEW, StringComparison.OrdinalIgnoreCase) Then
                    grpNewWhse(ORDR_GROUP_NO_orig) = ""
                End If
            End If
        Next

        Dim moveAllGroups As New HashSet(Of String)
        For Each kvp As KeyValuePair(Of String, Int32) In grpTotal
            Dim g As String = kvp.Key
            Dim total As Int32 = kvp.Value
            Dim changed As Int32 = 0
            If grpChanged.ContainsKey(g) Then changed = grpChanged(g)

            If changed = total Then
                If grpNewWhse.ContainsKey(g) Then
                    If (grpNewWhse(g) & "") <> "" Then
                        moveAllGroups.Add(g)
                    End If
                End If
            End If
        Next

        Dim splitMap As New Dictionary(Of String, String)
        Dim groupsToRecalc As New HashSet(Of String)
        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            If ORDR_GROUP_NO & "" <> "" Then
                groupsToRecalc.Add(ORDR_GROUP_NO)
            End If
        Next

        For Each TUP As Tuple(Of String, String, String, String) In whseChanges
            Dim ORDR_NO As String = TUP.Item1
            Dim WHSE_CODE_OLD As String = TUP.Item2 & ""
            Dim WHSE_CODE_NEW As String = TUP.Item3 & ""
            Dim ORDR_GROUP_NO_orig As String = TUP.Item4 & ""

            If ORDR_GROUP_NO_orig = "" Then Continue For
            If WHSE_CODE_NEW = "" Then Continue For
            If WHSE_CODE_OLD.Equals(WHSE_CODE_NEW) Then Continue For

            If moveAllGroups.Contains(ORDR_GROUP_NO_orig) Then Continue For

            Dim KEY As String = ORDR_GROUP_NO_orig & "|" & WHSE_CODE_NEW.ToUpperInvariant()

            Dim ORDR_GROUP_NO_new As String = ""
            If splitMap.ContainsKey(KEY) Then
                ORDR_GROUP_NO_new = splitMap(KEY)
            Else
                ORDR_GROUP_NO_new = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                splitMap(KEY) = ORDR_GROUP_NO_new
            End If

            If ORDR_GROUP_NO_new <> "" AndAlso Not ORDR_GROUP_NO_new.Equals(ORDR_GROUP_NO_orig) Then
                Dim rows() As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO.Replace("'", "''") & "'")
                If rows IsNot Nothing AndAlso rows.Length > 0 Then
                    rows(0)("ORDR_GROUP_NO") = ORDR_GROUP_NO_new
                    groupsToRecalc.Add(ORDR_GROUP_NO_new)
                End If
            End If
        Next

        Update_Record_TDA("SOTORDR1")
        For Each tup As Tuple(Of String, String, String, String) In whseChanges
            Dim ordr As String = tup.Item1
            Dim oldWhse As String = tup.Item2
            Dim newWhse As String = tup.Item3
            If (oldWhse & "") <> "" AndAlso (newWhse & "") <> "" AndAlso
       Not oldWhse.Equals(newWhse, StringComparison.OrdinalIgnoreCase) Then

                ' subtract from old
                Dependent_Updates(-1, ordr, oldWhse)
                ' add to new
                Dependent_Updates(+1, ordr, newWhse)
                'update SOTORDR2 lines to point at the new warehouse
                Update_SOTORDR2_WHSE_CODE(ordr, oldWhse, newWhse)
            End If
        Next

        For Each ORDR_GROUP_NO As String In groupsToRecalc
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next
        CommitTrans()

        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        CUST_CODE = ""
        Me.Close()
    End Sub

    Private Sub cmdFixDates_Click(sender As Object, e As EventArgs) Handles cmdFixDates.Click
        Dim C As String = optDateType.Value
        If C = "ORDR_SHIP_DATE" Or C = "ORDR_CANCEL_DATE" Then
            If dteFixed.Value & "" = "" Then
                MsgBox("Cannot Set Ship Date or Cancel Date to Nothing", MsgBoxStyle.OkOnly, "Cannot Edit")
                Exit Sub
            End If
        End If

        For Each row As DataRow In dst.Tables("SOTORDR1").Select("")
            If dteFixed.Value & "" = "" Then
                row.Item(C) = DBNull.Value
            Else
                row.Item(C) = dteFixed.Value
            End If

        Next
    End Sub

    Sub Set_Date()

        Dim row As DataRow = dst.Tables("SOTORDR1").Rows(0)
        dteFixed.Value = row.Item(optDateType.Value)
    End Sub

    Private Sub optDateType_ValueChanged(sender As Object, e As EventArgs) Handles optDateType.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Date()
    End Sub
    ' Fires when the Fix WHSE button is clicked
    Private Sub btnFixWhse_Click(sender As Object, e As EventArgs) Handles btnFixWhse.Click
        Dim code As String = (whseFixed.Text & "").Trim().ToUpperInvariant()

        If code = "" Then
            MsgBox("WHSE Code cannot be blank.", MsgBoxStyle.OkOnly, "Invalid WHSE Code")
            Exit Sub
        End If

        Dim exists As Boolean = False
        Dim val As Object = ASCDATA1.GetDataValue($"SELECT 1 FROM ICTWHSE1 WHERE WHSE_CODE = '{code.Replace("'", "''")}' AND ROWNUM = 1")
        exists = (val IsNot Nothing AndAlso (val & "") <> "")

        If Not exists Then
            MsgBox($"'{code}' is not a valid WHSE_CODE (not found in ICTWHSE1).", MsgBoxStyle.OkOnly, "Invalid WHSE Code")
            Exit Sub
        End If

        For Each r As DataRow In dst.Tables("SOTORDR1").Select("")
            r("WHSE_CODE") = code
        Next

    End Sub

    Private Sub Dependent_Updates(ByVal S As Integer, ByVal ORDR_NO As String, ByVal WHSE_CODE As String)
        If (WHSE_CODE & "") = "" Then Exit Sub

        Dim dt As DataTable = ASCDATA1.GetDataTable(
            $"SELECT ITEM_CODE,
                 NVL(ORDR_QTY_OPEN, 0) AS ORDR_QTY_OPEN
            FROM SOTORDR2
           WHERE ORDR_NO = '{ORDR_NO.Replace("'", "''")}'")

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Exit Sub

        For Each rowSOTORDR2 As DataRow In dt.Rows
            Dim item As String = (rowSOTORDR2("ITEM_CODE") & "").Trim()
            If item = "" Then Continue For

            Dim QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT = 0 Then Continue For

            Update_ICTSTAT2(item, WHSE_CODE, S * QTY_TO_COMMIT)
        Next
    End Sub

    Private Sub Update_ICTSTAT2(ByVal ITEM_CODE As String, ByVal WHSE_CODE As String, ByVal QTY As Long)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN",
            New Object() {ITEM_CODE, WHSE_CODE, 0, 0, 0, QTY, 0, 0},
            New String() {"ITEM_CODE_IN", "WHSE_CODE_IN",
                          "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ONPO_in", "WHSE_QTY_PLAN_in",
                          "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_COMM_in"})
    End Sub
    Private Sub Update_SOTORDR2_WHSE_CODE(ByVal ORDR_NO As String,
                                 ByVal oldWhse As String,
                                 ByVal newWhse As String)

        Dim ord As String = ORDR_NO.Replace("'", "''")
        Dim newC As String = (newWhse & "").Trim().Replace("'", "''")
        Dim whereWhse As String

        oldWhse = (oldWhse & "").Trim()
        Dim oldC As String = oldWhse.Replace("'", "''")
        whereWhse = $" WHSE_CODE = '{oldC}'"

        ASCMAIN1.sql =
        "UPDATE SOTORDR2 " &
        $"   SET WHSE_CODE = '{newC}' " &
        $" WHERE ORDR_NO = '{ord}'" &
        $"   AND {whereWhse}"

        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
    End Sub
    Private Function Get_Missing_ICTWHSE2_Items(ByVal ORDR_NO As String, ByVal WHSE_CODE_NEW As String) As List(Of String)

        Dim lst As New List(Of String)

        Dim ord As String = (ORDR_NO & "").Replace("'", "''")
        Dim wh As String = (WHSE_CODE_NEW & "").Trim().ToUpperInvariant().Replace("'", "''")

        If ord = "" OrElse wh = "" Then Return lst

        ASCMAIN1.sql =
        "SELECT DISTINCT d.ITEM_CODE " & vbCrLf &
        "  FROM SOTORDR2 d " & vbCrLf &
        " WHERE d.ORDR_NO = '" & ord & "'" & vbCrLf &
        "   AND (NVL(d.ORDR_QTY_OPEN,0) <> 0 OR NVL(d.ORDR_QTY,0) <> 0) " & vbCrLf &
        "   AND NOT EXISTS (SELECT 1 FROM ICTWHSE2 w " & vbCrLf &
        "                    WHERE w.WHSE_CODE = '" & wh & "'" & vbCrLf &
        "                      AND w.ITEM_CODE = d.ITEM_CODE)"

        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return lst

        For Each r As DataRow In dt.Rows
            Dim it As String = (r.Item("ITEM_CODE") & "").Trim()
            If it <> "" Then lst.Add(it)
        Next

        Return lst

    End Function

End Class