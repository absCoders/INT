Public Class TATIDOC1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "TATIDOC2", "*", 1)
            Create_TDA(.Tables.Add, "TATIDOC3", "*", 1)
        End With

        grdTATIDOC2.DataSource = dst.Tables("TATIDOC2")
        grdTATIDOC3.DataSource = dst.Tables("TATIDOC3")

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        '   Load_Popup_Menu(grdTATIDOC2, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        ASCDATA1.ExecuteSQL("Delete from TATIDOC2")
        ASCDATA1.ExecuteSQL("Delete from TATIDOC3")

        Update_Record_TDA("TATIDOC2")
        Update_Record_TDA("TATIDOC3")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("TATIDOC2", New String() {Absx1.txtFor("IDOC_TABLE").Text})
        Sort_grdColumns(grdTATIDOC2, "IDOC_ID")
        Fill_Records("TATIDOC3", New String() {Absx1.txtFor("IDOC_TABLE").Text})
        Sort_grdColumns(grdTATIDOC3, "IDOC_SEGMENT")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"TATIDOC2", "TATIDOC2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdTATIDOC2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        grdTATIDOC2.Visible = tf
        grdTATIDOC3.Visible = tf
        btnLoadDefinition.Visible = tf
    End Sub

    Public Overrides Function Remote_Control( _
ByVal command As String, _
Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "View"
                If key <> "" Then
                    Absx1.txtFor("IDOC_TABLE").Text = key
                    Click_Command(command)
                End If
        End Select

        Return return_key
    End Function

#End Region

    Sub Load_IDOC_Definition()

        Dim IDOC_TABLE As String = Absx1.txtFor("IDOC_TABLE").Text ' "INVOIC01"
        Dim IDOC_ID As String = ""
        Dim IDOC_SEGMENT As String = ""
        Dim IDOC_DATUM_NO As Integer = 0
        Dim IDOC_DATUM_NAME As String = ""
        Dim IDOC_DATUM_DESC As String = ""
        Dim IDOC_DATUM_TYPE As String = ""
        Dim IDOC_DATUM_LENGTH As Integer = 0

        dst.Tables("TATIDOC2").Rows.Clear()
        dst.Tables("TATIDOC3").Rows.Clear()

        ' USE A FILEDIALOG TO FIND THE FILE HERE
        Dim FILENAME As String = "C:\Users\wjz\Desktop\Interparfums\IPSA\IDOC_Layouts\IDOC.TXT"
        FILENAME = "C:\Users\walter\Desktop\Interparfums\IPSA\IDOC_Layouts\IDOC_REV.TXT"
        FILENAME = "C:\Users\wjz\Desktop\Interparfums\IPSA\IDOC_Layouts\rev2\IDOC_REV2.TXT"

        Using SR As New System.IO.StreamReader(FILENAME)
            Dim TT As String = SR.ReadToEnd
            Dim CC() As String = Split(TT, vbCrLf)

            Dim C As Integer = -1
            Do Until C >= CC.Length - 1
                C += 1
                Dim T As String = CC(C)
                If T <> "" Then
                    If T.StartsWith("Info...") Then
                        Exit Do
                    End If
                    If InStr(T, ":") <> 0 And CC.Length >= C + 3 AndAlso CC(C + 1).StartsWith("internal data type") AndAlso Trim(Split(T, ":")(1)) <> "IDoc" Then
                        IDOC_DATUM_NO += 1
                        IDOC_DATUM_NAME = Trim(Split(T, ":")(0))
                        IDOC_DATUM_DESC = Trim(Split(T, ":")(1))
                        If IDOC_DATUM_DESC = "IDoc" Then
                            C += 4
                        Else
                            C += 1 : T = CC(C)
                            If T = "" Then
                                C += 1
                            Else
                                IDOC_DATUM_TYPE = Trim(Split(T, ":")(1))
                                C += 1 : T = CC(C)
                                IDOC_DATUM_LENGTH = Trim(Split(Trim(Split(T, ":")(1)), " ")(0))
                                C += 1
                                'If CC(C).StartsWith("No decimal places, without sign") Then
                                If CC(C).Contains(" decimal places, without sign") Then
                                    C += 1
                                End If
                                dst.Tables("TATIDOC3").Rows.Add(New Object() {IDOC_TABLE, IDOC_SEGMENT, IDOC_DATUM_NO, IDOC_DATUM_NAME, IDOC_DATUM_DESC, IDOC_DATUM_TYPE, IDOC_DATUM_LENGTH})
                            End If
                        End If
                    Else
                        If Mid(T, 1, 1) <> " " Then
                            If T = IDOC_SEGMENT & " structure" Then
                            Else
                                If T = "Structure of basic type INVOIC01" Then

                                    Do While Trim(CC(C)) <> "Segment structures"
                                        C += 1
                                    Loop
                                    ' C += 1 ' older format had a space line following, new one does not
                                    T = CC(C)
                                Else
                                    Dim IDOC_DESC As String = ""
                                    If InStr(T, ":") <> 0 AndAlso Trim(Split(T, ":")(1)) = "IDoc" Then
                                        IDOC_ID = Trim(Split(T, ":")(0))
                                        IDOC_DESC = Trim(Split(T, ":")(2))
                                        C += 1 ' C += 2 - CHANGE REQUIRED BECAUSE A SPACE LINE IS NO LONGER PRESENT
                                        T = CC(C)
                                        IDOC_SEGMENT = Split(T & "  ", " ")(2)
                                        Dim IDOC_RELEASED_SINCE As String = Split(Split(T & "since Release ", "since Release ")(1), " ")(0) ' IN CASE WE WANT TO ADD THIS IN THE FUTURE
                                        Dim IDOC_SEGMENT_LENGTH As String = Split(T & ", Segment length: ", ", Segment length: ")(1) ' IN CASE WE WANT TO ADD THIS IN THE FUTURE
                                        ' C += 1 ' older format had a space line following, new one does not
                                    Else
                                        IDOC_DESC = T
                                        IDOC_SEGMENT = Split(T, " ")(0)
                                        IDOC_ID = IDOC_SEGMENT
                                        C += 2
                                    End If
                                    If IDOC_SEGMENT.Contains("Position ") Or IDOC_SEGMENT.Contains("Segment ") Then
                                        Stop
                                    End If

                                    IDOC_DATUM_NO = 0

                                    dst.Tables("TATIDOC2").Rows.Add(New Object() {IDOC_TABLE, IDOC_ID, IDOC_SEGMENT, IDOC_DESC})
                                End If
                            End If
                        Else
                            Stop
                        End If
                    End If
                End If
            Loop
        End Using

    End Sub

    Sub Load_IDOC_Definition_Previous()

        Dim IDOC_TABLE As String = Absx1.txtFor("IDOC_TABLE").Text ' "INVOIC01"
        Dim IDOC_ID As String = ""
        Dim IDOC_SEGMENT As String = ""
        Dim IDOC_DATUM_NO As Integer = 0
        Dim IDOC_DATUM_NAME As String = ""
        Dim IDOC_DATUM_DESC As String = ""
        Dim IDOC_DATUM_TYPE As String = ""
        Dim IDOC_DATUM_LENGTH As Integer = 0

        dst.Tables("TATIDOC2").Rows.Clear()
        dst.Tables("TATIDOC3").Rows.Clear()

        ' USE A FILEDIALOG TO FIND THE FILE HERE
        Dim FILENAME As String = "C:\Users\wjz\Desktop\Interparfums\IPSA\IDOC_Layouts\IDOC.TXT"
        FILENAME = "C:\Users\walter\Desktop\Interparfums\IPSA\IDOC_Layouts\IDOC_REV.TXT"

        Using SR As New System.IO.StreamReader(FILENAME)
            Dim TT As String = SR.ReadToEnd
            Dim CC() As String = Split(TT, vbCrLf)

            Dim C As Integer = -1
            Do Until C >= CC.Length - 1
                C += 1
                Dim T As String = CC(C)
                If T <> "" Then
                    If T.StartsWith("Info...") Then
                        Exit Do
                    End If
                    If InStr(T, ":") <> 0 And CC.Length >= C + 3 AndAlso CC(C + 1).StartsWith("internal data type") AndAlso Trim(Split(T, ":")(1)) <> "IDoc" Then
                        IDOC_DATUM_NO += 1
                        IDOC_DATUM_NAME = Trim(Split(T, ":")(0))
                        IDOC_DATUM_DESC = Trim(Split(T, ":")(1))
                        If IDOC_DATUM_DESC = "IDoc" Then
                            C += 4
                        Else
                            C += 1 : T = CC(C)
                            If T = "" Then
                                C += 1
                            Else
                                IDOC_DATUM_TYPE = Trim(Split(T, ":")(1))
                                C += 1 : T = CC(C)
                                IDOC_DATUM_LENGTH = Trim(Split(Trim(Split(T, ":")(1)), " ")(0))
                                C += 1
                                If CC(C).StartsWith("No decimal places, without sign") Then
                                    C += 1
                                End If
                                dst.Tables("TATIDOC3").Rows.Add(New Object() {IDOC_TABLE, IDOC_SEGMENT, IDOC_DATUM_NO, IDOC_DATUM_NAME, IDOC_DATUM_DESC, IDOC_DATUM_TYPE, IDOC_DATUM_LENGTH})
                            End If
                        End If
                    Else
                        If Mid(T, 1, 1) <> " " Then
                            If T = IDOC_SEGMENT & " structure" Then
                            Else
                                If T = "Structure of basic type INVOIC01" Then

                                    Do While Trim(CC(C)) <> "Segment structures"
                                        C += 1
                                    Loop
                                    C += 1
                                    T = CC(C)
                                Else
                                    Dim IDOC_DESC As String = ""
                                    If InStr(T, ":") <> 0 AndAlso Trim(Split(T, ":")(1)) = "IDoc" Then
                                        IDOC_ID = Trim(Split(T, ":")(0))
                                        IDOC_DESC = Trim(Split(T, ":")(2))
                                        C += 2
                                        T = CC(C)
                                        IDOC_SEGMENT = Split(T & "  ", " ")(2)
                                        C += 1
                                    Else
                                        IDOC_DESC = T
                                        IDOC_SEGMENT = Split(T, " ")(0)
                                        IDOC_ID = IDOC_SEGMENT
                                        C += 2
                                    End If
                                    If IDOC_SEGMENT.Contains("Position ") Or IDOC_SEGMENT.Contains("Segment ") Then
                                        Stop
                                    End If

                                    IDOC_DATUM_NO = 0

                                    dst.Tables("TATIDOC2").Rows.Add(New Object() {IDOC_TABLE, IDOC_ID, IDOC_SEGMENT, IDOC_DESC})
                                End If
                            End If
                        Else
                            Stop
                        End If
                    End If
                End If
            Loop
        End Using

    End Sub
    Private Sub btnLoadDefinition_Click(sender As Object, e As EventArgs) Handles btnLoadDefinition.Click
        Load_IDOC_Definition()
    End Sub
End Class