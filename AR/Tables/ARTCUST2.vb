Public Class ARTCUST2

    Private Sub ARTCUST2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.loadFromXls = LoadFromXlsType.Basic
        btnCheckAllAddresses.Visible = (ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz")
    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            rowASFBASE1.Item("CUST_STORE_STATUS") = "A"
            rowASFBASE1.Item("CUST_NO_3PL") = DBNull.Value
            rowASFBASE1.Item("CUST_STORE_NO_3PL") = DBNull.Value
        End If

        MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = True
        MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = True

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            If EntryMode = "Edit" Then
                MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = False
                MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = False
            End If
        Else
            MyBase.Absx1.txtFor("CUST_NO_3PL").Visible = False
            MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").Visible = False
            lbl3PL.Visible = False
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If SELECTION_NO = 0 Then Exit Sub

        MyBase.Absx1.txtFor("CUST_NO_3PL").ReadOnly = True
        MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").ReadOnly = True
    End Sub

    Public Overrides Sub Proceed_PreReq_Special(eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Select Case eItemKey
            Case "New"

                If LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Is Nothing Then
                    EMsg &= vbCr & "Invalid Customer"
                End If


            Case "Edit"
                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.txtFor("CUST_CODE").Text = "IPLBAE" Then
                        EMsg &= vbCr & "Use Sell-Thru Rep Maintenance"
                    End If
                End If

            Case "Update"
                ' validate City, State Zip.

                check_for_special_Characters

                If ASCMAIN1.CLIENT = "INT" Then
                    'Dim CUST_STORE_NAME As String = Absx1.txtFor("CUST_STORE_NAME").Text
                    'Dim CUST_STORE_ADDR1 As String = Absx1.txtFor("CUST_STORE_ADDR1").Text
                    'Dim CUST_STORE_ADDR2 As String = Absx1.txtFor("CUST_STORE_ADDR2").Text
                    'Dim CUST_STORE_CITY As String = Absx1.txtFor("CUST_STORE_CITY").Text
                    'Dim rx As String = "^([A-Z0-9]+)$" ' Allow Upper case, numbers

                    'Dim r As New System.Text.RegularExpressions.Regex(rx)
                    'If Not r.IsMatch(CUST_STORE_NAME) Then
                    '    EMsg &= vbCr & "The Customer Store Name has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(CUST_STORE_ADDR1) Then
                    '    EMsg &= vbCr & "The Customer Store Address 1 has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(CUST_STORE_ADDR2) Then
                    '    EMsg &= vbCr & "The Customer Store Address 2 has Special Characters which are not allowed"
                    'End If
                    'If Not r.IsMatch(CUST_STORE_CITY) Then
                    '    EMsg &= vbCr & "The Store City has Special Characters which are not allowed"
                    'End If
                    'If Absx1.optFor("CUST_STORE_STATUS").Value & "" <> "A" Then
                    '    If Absx1.txtFor("SDS_CODE").Text <> "" Then
                    '        EMsg &= vbCr & "Cannot Change Store to Inactive if an SDS code is defined"
                    '    End If
                    'End If
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
                    Dim SDS_CODE As String = Absx1.txtFor("SDS_CODE").Text
                    Dim SELL_CODE_AC As String = Absx1.txtFor("SELL_CODE_AC").Text

                    If SELL_CODE_AC <> "" Then
                        'If SDS_CODE <> "" Then
                        '    EMsg &= vbCr & "A Store Cannot have both an SDS and an AC defined"
                        'End If
                        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE_AC)
                        If rowSOTSELL1 Is Nothing Then
                            EMsg &= vbCr & "Invalid AC defined"
                        Else
                            Dim SELL_CODE_MGR As String = rowSOTSELL1.Item("SELL_CODE_MGR") & ""
                            Dim SELL_TYPE As String = rowSOTSELL1.Item("SELL_TYPE") & ""
                            If SELL_CODE_MGR = "" Or SELL_TYPE <> "AC" Then
                                EMsg &= vbCr & "Invalid AC defined - Does not appear to be an AC (no Managing AE)"
                            Else
                                If SELL_CODE_MGR <> SELL_CODE Then
                                    EMsg &= vbCr & "Invalid AC defined - Managing AE (" & SELL_CODE_MGR & ") is not defined as the AE for this Store"
                                End If
                            End If
                        End If
                    End If

                    If SELL_CODE <> "" Then
                        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE)
                        If rowSOTSELL1 Is Nothing Then
                            EMsg &= vbCr & "Invalid AE defined"
                        Else
                            Dim SELL_CODE_MGR As String = rowSOTSELL1.Item("SELL_CODE_MGR") & ""
                            Dim SELL_TYPE As String = rowSOTSELL1.Item("SELL_TYPE") & ""
                            If SELL_CODE_MGR <> "" Or SELL_TYPE <> "AE" Then
                                EMsg &= vbCr & "Invalid AE defined - an AE should not have a Manager Defined (like an AC would)"
                            End If
                        End If
                    End If
                End If

                Dim sqlCityState As String = String.Empty

                MyBase.Absx1.txtFor("CUST_STORE_STATE").Text = MyBase.Absx1.txtFor("CUST_STORE_STATE").Text.Trim.ToUpper
                MyBase.Absx1.txtFor("CUST_STORE_CITY").Text = MyBase.Absx1.txtFor("CUST_STORE_CITY").Text.Trim.ToUpper
                MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text = MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text.Trim.ToUpper

                MyBase.Absx1.txtFor("CUST_NO_3PL").Text = MyBase.Absx1.txtFor("CUST_NO_3PL").Text.Trim
                MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").Text = MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").Text.Trim

                If EntryMode = "New" Then
                    If MyBase.Absx1.txtFor("CUST_NO_3PL").TextLength > 0 Then
                        EMsg &= vbCr & "You are not permitted to provide the Customer 3PL value."
                    End If
                    If MyBase.Absx1.txtFor("CUST_STORE_NO_3PL").TextLength > 0 Then
                        EMsg &= vbCr & "You are not permitted to provide the Store 3PL value."
                    End If
                End If

                Dim tblSOTZIPLK As DataTable =
                    ASCDATA1.GetDataTable("Select * from SOTZIPLK Where STATE_CODE = :PARM1 AND CITY = :PARM2 AND ZIP_CODE = :PARM3",
                                          "SOTZIPLK", "VVV",
                                          MyBase.Absx1.txtFor("CUST_STORE_STATE").Text,
                                          MyBase.Absx1.txtFor("CUST_STORE_CITY").Text,
                                          Mid(MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text, 1, 5))


                If tblSOTZIPLK.Rows.Count = 0 Then
                    tblSOTZIPLK =
                        ASCDATA1.GetDataTable("Select * from SOTZIPLK Where STATE_CODE = :PARM1 AND ZIP_CODE = :PARM2",
                                              "SOTZIPLK", "VV",
                                              MyBase.Absx1.txtFor("CUST_STORE_STATE").Text,
                                              Mid(MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text, 1, 5))

                    If tblSOTZIPLK.Rows.Count > 0 Then
                        If MessageBox.Show("There were no entries found for the provided City/State/Zip Code combination. However, there are entries for the" _
                                            & " provided State/Zip Code." & Environment.NewLine & "Would you like to see the City options?", "Address Validation", MessageBoxButtons.YesNo,
                                              MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                            sqlCityState = "Select * from SOTZIPLK Where STATE_CODE = '" & MyBase.Absx1.txtFor("CUST_STORE_STATE").Text & "' AND ZIP_CODE = '" & MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text & "'"
                        End If
                    Else
                        tblSOTZIPLK =
                            ASCDATA1.GetDataTable("Select * from SOTZIPLK Where ZIP_CODE = :PARM1",
                                                  "SOTZIPLK", "V",
                                                  Mid(MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text, 1, 5))
                        If tblSOTZIPLK.Rows.Count > 0 Then
                            If MessageBox.Show("There were no entries found for the provided City/State/Zip Code combination. However, there are entries for the" _
                                                & " provided Zip Code." & Environment.NewLine & "Would you like to see the City/State options?", "Address Validation", MessageBoxButtons.YesNo,
                                                  MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                                sqlCityState = "Select * from SOTZIPLK Where ZIP_CODE = '" & Mid(MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text, 1, 5) & "'"
                            End If
                        Else
                            If MessageBox.Show("There were no City/State/Zip Code entries found for the provided Zip Code. Do you want to continue updating the record?",
                                                 "Address Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                                EMsg = vbCr & "Update cancelled by user."
                            End If
                        End If
                    End If
                End If

                If sqlCityState.Length > 0 AndAlso tblSOTZIPLK.Rows.Count > 0 Then
                    ASCMAIN1.CodeSelector.Get_SQL("")
                    ASCMAIN1.CodeSelector.VIEW_NAME = String.Empty
                    ASCMAIN1.CodeSelector.SQL = sqlCityState
                    ASCMAIN1.CodeSelector.UseDataFromTable = tblSOTZIPLK
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                        MyBase.Absx1.txtFor("CUST_STORE_STATE").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STATE_CODE")
                        MyBase.Absx1.txtFor("CUST_STORE_CITY").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CITY")
                        MyBase.Absx1.txtFor("CUST_STORE_ZIP_CODE").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ZIP_CODE")

                        dst.Tables("ARTCUST2").Rows(0).Item("CUST_STORE_STATE") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STATE_CODE")
                        dst.Tables("ARTCUST2").Rows(0).Item("CUST_STORE_CITY") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CITY")
                        dst.Tables("ARTCUST2").Rows(0).Item("CUST_STORE_ZIP_CODE") = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ZIP_CODE")

                    End If
                End If
        End Select
    End Sub

    Protected Overrides Function Allow_Validate_Lookup(ByVal COLUMN_NAME As String)
        If COLUMN_NAME = "CUST_DC_NO" Then
            Return Not (Absx1.chkFor("CUST_DC_IND").Checked AndAlso Absx1.txtFor("CUST_DC_NO").Text & "" = Absx1.txtFor("CUST_STORE_NO").Text & "")
        Else
            Return True
        End If
    End Function

    Overrides Sub Proceed_Update_Special_Pre()

        If ASCMAIN1.CLIENT = "INT" Then
            ' note - doing it this way does not work
            'If Absx1.optFor("CUST_STORE_STATUS").Value & "" = "I" Then
            '    Absx1.txtFor("SDS_CODE").Text = ""
            '    Absx1.txtFor("SELL_CODE").Text = ""
            'End If

            ' note - LBM wants to disable this for now 05/08/2018
            ' RE-ENABLING SINCE WE ARE SAVING THE LAST AE - SEE SP APPROVAL EMAIL 5/31 RE: removal of AE and RSC when door is closed



            If rowASFBASE1.Item("CUST_STORE_STATUS") & "" = "I" Then
                ' NOTE- BECAUSE THIS PROCEDURE FIRES AFTER THE AUDIT TRAIL IS WRITTEN, THESE CHANGES ARE NOT AUDITED.
                Dim SELL_CODE As String = rowASFBASE1.Item("SELL_CODE", DataRowVersion.Original) & ""
                If SELL_CODE <> "" Then
                    rowASFBASE1.Item("SELL_CODE_LAST") = SELL_CODE
                    rowASFBASE1.Item("SDS_CODE") = DBNull.Value
                    rowASFBASE1.Item("SELL_CODE") = DBNull.Value
                    rowASFBASE1.Item("SELL_CODE_LAST_YP") = ASCMAIN1.CYP
                End If
                Dim SELL_CODE_AC As String = rowASFBASE1.Item("SELL_CODE_AC", DataRowVersion.Original) & ""
                If SELL_CODE_AC <> "" Then
                    rowASFBASE1.Item("SELL_CODE_AC_LAST") = SELL_CODE_AC
                    rowASFBASE1.Item("SELL_CODE_AC") = DBNull.Value
                    rowASFBASE1.Item("SELL_CODE_AC_LAST_YP") = ASCMAIN1.CYP
                End If

            ElseIf rowASFBASE1.Item("CUST_STORE_STATUS") & "" = "A" Then
                Dim SELL_CODE_LAST As String = ""
                Try
                    SELL_CODE_LAST = rowASFBASE1.Item("SELL_CODE_LAST", DataRowVersion.Original) & ""
                Catch ex As Exception

                End Try

                If SELL_CODE_LAST <> "" Then
                    rowASFBASE1.Item("SELL_CODE_LAST") = DBNull.Value
                    rowASFBASE1.Item("SELL_CODE_LAST_YP") = DBNull.Value
                End If
                Dim SELL_CODE_AC_LAST As String = ""
                Try
                    SELL_CODE_AC_LAST = rowASFBASE1.Item("SELL_CODE_AC_LAST", DataRowVersion.Original) & ""
                Catch ex As Exception

                End Try

                If SELL_CODE_AC_LAST <> "" Then
                    rowASFBASE1.Item("SELL_CODE_AC_LAST") = DBNull.Value
                    rowASFBASE1.Item("SELL_CODE_AC_LAST_YP") = DBNull.Value
                End If
            End If
        End If
    End Sub

    Sub Check_for_Special_Characters()

        ' ARFCUSTX HAS SAME CHECKS
        ' this code snippet was copied from ICTITEM1.Proceed_PreReq_Special

        Dim rx As String = "[^a-zA-Z0-9 '#/.&@:;,+_()-]"
        Dim r As New System.Text.RegularExpressions.Regex(rx)

        For Each C As String In New String() {"_NAME", "_ADDR1", "_ADDR2", "_ADDR3", "_CITY", "_STATE", "_ZIP_CODE",
            "_COUNTRY", "_CONTACT", "_EXT", "_EMAIL", "_LOCATION", "_MIL_CODE", "_MARK_FOR"}
            Dim CC As String = "CUST_STORE" & C
            If r.IsMatch(rowASFBASE1.Item(CC) & "") Then ' If r.IsMatch(Absx1.txtFor(CC).Text) Then
                EMsg &= vbCr & $"{CC} has Special Characters which are not allowed"
            End If
        Next

    End Sub

    Private Sub btnCheckAllAddresses_Click(sender As Object, e As EventArgs) Handles btnCheckAllAddresses.Click
        ASCMAIN1.sql = "Select * from ARTCUST2"
        Dim tbl As DataTable = ASCDATA1.GetDataTable

        For Each rowASFBASE1 In tbl.Select("", "CUST_CODE, CUST_STORE_NO")
            Dim CUST_CODE As String = rowASFBASE1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowASFBASE1.Item("CUST_STORE_NO")
            EMsg = ""
            Check_for_Special_Characters()
            If EMsg <> "" Then
                Dim EE As String = Mid(EMsg, 2)
                EE = Mid(EE, 1, InStr(EE, " has ") - 1)
                Dim d As String = rowASFBASE1.Item(EE)
                EE &= d & ":" & ASCMAIN1.AscToHex(d)
                Debug.Print(CUST_CODE & ":" & CUST_STORE_NO & ":" & EE)
            End If
        Next
    End Sub

End Class