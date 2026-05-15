Public Class SOTSELL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2 where CUST_CODE = 'IPLBAE' and CUST_STORE_NO = '000' || :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "V", 2)
        End With
    End Sub
   
#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        If ASCMAIN1.CLIENT = "INT" Then
            Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
            Dim rowARTCUST2 As DataRow = Fill_Record("ARTCUST2", SELL_CODE)
            If rowARTCUST2 Is Nothing Then
                rowARTCUST2 = dst.Tables("ARTCUST2").NewRow
                rowARTCUST2.Item("CUST_CODE") = "IPLBAE"
                rowARTCUST2.Item("CUST_STORE_NO") = "000" & SELL_CODE
                rowARTCUST2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowARTCUST2.Item("INIT_DATE") = DATETIME_STAMP
                dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)
            End If
            Dim SELL_NAME As String = rowASFBASE1.Item("SELL_NAME") & "" ' Absx1.txtFor("SELL_NAME").Text
            rowARTCUST2.Item("CUST_STORE_NAME") = SELL_NAME
            rowARTCUST2.Item("CUST_STORE_ADDR1") = Absx1.txtFor("SELL_ADDR1").Text
            rowARTCUST2.Item("CUST_STORE_ADDR2") = Absx1.txtFor("SELL_ADDR2").Text
            rowARTCUST2.Item("CUST_STORE_ADDR3") = Absx1.txtFor("SELL_ADDR3").Text
            rowARTCUST2.Item("CUST_STORE_CITY") = Absx1.txtFor("SELL_CITY").Text
            rowARTCUST2.Item("CUST_STORE_STATE") = Absx1.txtFor("SELL_STATE").Text
            rowARTCUST2.Item("CUST_STORE_ZIP_CODE") = Absx1.txtFor("SELL_ZIP_CODE").Text
            ' rowARTCUST2.Item("CUST_STORE_COUNTRY") = Absx1.txtFor("SELL_COUNTRY").Text
            rowARTCUST2.Item("CUST_STORE_PHONE") = Absx1.medFor("SELL_PHONE").Text
            rowARTCUST2.Item("CUST_STORE_EXT") = Absx1.txtFor("SELL_EXT").Text
            rowARTCUST2.Item("CUST_STORE_FAX") = Absx1.medFor("SELL_FAX").Text
            rowARTCUST2.Item("CUST_STORE_EMAIL") = Absx1.txtFor("SELL_EMAIL").Text
            rowARTCUST2.Item("CUST_STORE_STATUS") = Absx1.optFor("SELL_STATUS").Value
            rowARTCUST2.Item("SELL_CODE") = SELL_CODE
            rowARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowARTCUST2.Item("LAST_DATE") = DATETIME_STAMP

            Update_Record_TDA("ARTCUST2")
        End If

    End Sub

    Overrides Sub Show_Record_Special()
 
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ARTCUST2").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If ASCMAIN1.CLIENT = "INT" Then
            Absx1.optFor("SELL_TYPE").Visible = True
            If EntryMode = "New" Then
                Set_Read_Only_for_ctl(Absx1.optFor("SELL_TYPE"), False)
            Else
                Set_Read_Only_for_ctl(Absx1.optFor("SELL_TYPE"), True)
            End If
        Else
            Absx1.optFor("SELL_TYPE").Visible = False
        End If
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Update"

                If ASCMAIN1.CLIENT = "INT" Then
                    If txtSELL_NAME.Text = "" Then
                        EMsg &= vbCr & "Name is Mandatory"
                    End If
                    If Absx1.optFor("SELL_STATUS").Value & "" = "" Then
                        EMsg &= vbCr & "Status is Mandatory"
                    End If
                    If Absx1.txtFor("SELL_EMAIL").Text = "" Then
                        EMsg &= vbCr & "email is Mandatory"
                    End If
                    If Absx1.txtFor("REGION_CODE").Text = "" Then
                        EMsg &= vbCr & "Region is Mandatory"
                    End If
                    If Absx1.txtFor("USER_ID").Text = "" Then
                        EMsg &= vbCr & "User ID is Mandatory"
                    End If

                    If EMsg = "" Then
                        Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
                        Dim SELL_CODE_MGR As String = Absx1.txtFor("SELL_CODE_MGR").Text

                        Dim SELL_TYPE As String = Absx1.optFor("SELL_TYPE").Value & ""

                        If SELL_TYPE = "" Then
                            EMsg &= vbCr & "You must choose either Acct Exec or Acct Coord"
                        Else
                            If SELL_TYPE = "AC" Then
                                If SELL_CODE_MGR = "" Then
                                    EMsg &= vbCr & "Acct Coord must have a Managing AE defined"
                                Else
                                    Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE_MGR)
                                    If rowSOTSELL1 Is Nothing Then
                                        EMsg &= vbCr & "Invalid Manager defined"
                                    Else
                                        If rowSOTSELL1.Item("SELL_CODE_MGR") & "" <> "" Or rowSOTSELL1.Item("SELL_TYPE") & "" <> "AE" Then
                                            EMsg &= vbCr & "Invalid Manager defined - appears to be an AC"
                                        End If
                                    End If
                                End If
                            Else
                                If SELL_CODE_MGR <> "" Then
                                    EMsg &= vbCr & "Acct Exec must Not have a Manager defined (this is for ACs only)"
                                End If
                            End If
                        End If
                    End If
                   
                End If
        End Select
    End Sub
#End Region

 
End Class