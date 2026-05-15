Public Class SPFCWRXX

    Dim SPTCWRXY As String
    Dim SPTCWRXZ As String
    Dim rowSPTCWRXX As DataRow
    Dim CWRX_NO As String
    Dim FILENAME As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_Work_Tables(True)

        With dst
            Create_TDA(.Tables.Add, "GLTPARM3", "*", 0)

            ASCMAIN1.sql = "Select * from SPTCWRXX where CWRX_EXPORT_TYPE = 'S'"
            Create_TDA(.Tables.Add, "SPTCWRXX", "**", 0)

            ASCMAIN1.sql = "Select * from " & SPTCWRXY
            Create_TDA(.Tables.Add, "SPTCWRXY", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select * from " & SPTCWRXZ
            Create_TDA(.Tables.Add, "SPTCWRXZ", "**", 0, False, "", 0)
        End With

        grdSPTCWRXX.DataSource = dst.Tables("SPTCWRXX")
        grdSPTCWRXY.DataSource = dst.Tables("SPTCWRXY")
        grdSPTCWRXZ.DataSource = dst.Tables("SPTCWRXZ")

        Create_Summary(grdSPTCWRXX, "CWRX_NO", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSPTCWRXX.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTCWRXX", "SPTCWRXY", "SPTCWRXZ"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("SPTCWRXX")
        Sort_grdColumns(grdSPTCWRXX, "CWRX_NO".ToLower)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        CWRX_NO = ASCMAIN1.Next_Control_No("SPTCWRXX.CWRX_NO")
        dst.Tables("SPTCWRXX").AcceptChanges()
        rowSPTCWRXX = dst.Tables("SPTCWRXX").NewRow
        rowSPTCWRXX.Item("CWRX_NO") = CWRX_NO
        rowSPTCWRXX.Item("CWRX_EXPORT_TYPE") = "S"
        rowSPTCWRXX.Item("CWRX_DATE") = Now.Date
        dst.Tables("SPTCWRXX").Rows.Add(rowSPTCWRXX)

        Create_Work_Tables(False)

        EnforceConstraints(False)
        Fill_Records("SPTCWRXY")
        Fill_Records("SPTCWRXZ")

        Sort_grdColumns(grdSPTCWRXZ, "CUST_CODE,CUST_STORE_NO")
        Sort_grdColumns(grdSPTCWRXY, "CUST_CODE,CUST_STORE_NO")

        Create_Store_File()
        rowSPTCWRXX.Item("FILENAME") = FILENAME

        rowSPTCWRXX.Item("CWRX_RECS_ADD") = dst.Tables("SPTCWRXY").Select("REASON_CODE = 'N'").Length
        rowSPTCWRXX.Item("CWRX_RECS_CHG") = dst.Tables("SPTCWRXY").Select("REASON_CODE = 'C'").Length
        rowSPTCWRXX.Item("CWRX_RECS_REA") = dst.Tables("SPTCWRXY").Select("REASON_CODE = 'R'").Length
        rowSPTCWRXX.Item("CWRX_RECS_TRM") = dst.Tables("SPTCWRXY").Select("REASON_CODE = 'T'").Length

        rowSPTCWRXX.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSPTCWRXX.Item("INIT_DATE") = DATETIME_STAMP

        Fill_Records("GLTPARM3")

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now sftp'ing the file")

        BeginTrans()

        ' Delta - this is what was sent to CoWorx
        ASCMAIN1.sql = "Insert into SPTCWRXY Select * from " & SPTCWRXY
        ASCDATA1.ExecuteSQL()

        ' Image - this is what we will compare to next time
        ASCMAIN1.sql = "Delete from SPTCWRXZ"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into SPTCWRXZ Select * from " & SPTCWRXZ
        ASCDATA1.ExecuteSQL()

        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "\" & FILENAME, ASCMAIN1.Folders("Archive") & "\COWORX\" & FILENAME)
        If dst.Tables("SPTCWRXY").Rows.Count > 0 Then
            TAC.TACSCOM1.sftp_put(Me, "COWORX", True, ASCMAIN1.Folders("Temp") & "\" & FILENAME, FILENAME)
        End If

        Update_Record_TDA("SPTCWRXX")

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTCWRXX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTCWRXZ, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        Select Case e.SourceControl.Name
            Case "grdSATSLSC1"

            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Absx1.txtFor("OPS_YYYYPP").Text <> "" Then
                        Click_Command("Load", e)
                    End If
                End If
        End Select
    End Sub
#End Region

    Sub Create_Work_Tables(initialize As Boolean)
        If initialize Then
            ASCMAIN1.sql = "Select * from SPTCWRXY where ROWNUM < 1"
            SPTCWRXY = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCWRXY & " Add Primary Key (CUST_CODE,CUST_STORE_NO)")

            ASCMAIN1.sql = "Select * from SPTCWRXZ where ROWNUM < 1"
            SPTCWRXZ = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCWRXZ & " Add Primary Key (CUST_CODE,CUST_STORE_NO)")
        Else
            ' Z IS THE CURRENT MASTER (Image)
            ' Y IS THE CHANGES (Delta)

            ASCMAIN1.sql = "Delete from " & SPTCWRXZ
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & SPTCWRXZ & " " & vbCrLf _
                & "Select " & vbCrLf _
                & "ARTCUST1.CUST_NAME," & vbCrLf _
                & "ARTCUST2.CUST_CODE," & vbCrLf _
                & "ARTCUST2.CUST_STORE_NO," & vbCrLf _
                & "ARTCUST2.CUST_STORE_NAME," & vbCrLf _
                & "ARTCUST2.CUST_STORE_ADDR1," & vbCrLf _
                & "ARTCUST2.CUST_STORE_ADDR2," & vbCrLf _
                & "ARTCUST2.CUST_STORE_CITY," & vbCrLf _
                & "ARTCUST2.CUST_STORE_STATE," & vbCrLf _
                & "ARTCUST2.CUST_STORE_ZIP_CODE," & vbCrLf _
                & "ARTCUST2.CUST_STORE_PHONE," & vbCrLf _
                & "ARTCUST2.SELL_CODE," & vbCrLf _
                & "SOTSELL1.SELL_NAME," & vbCrLf _
                & "SOTSELL1.REGION_CODE," & vbCrLf _
                & "SOTSREG1.REGION_DESC," & vbCrLf _
                & "SOTSREG1.VP_CODE," & vbCrLf _
                & "SOTSVPS1.VP_NAME," & vbCrLf _
                & "SPTCWRXZ.CHECK_BOOK," & vbCrLf _
                & "SPTCWRXZ.INIT_DATE," & vbCrLf _
                & "SPTCWRXZ.LAST_DATE," & vbCrLf _
                & "SPTCWRXZ.TERM_DATE," & vbCrLf _
                & "SPTCWRXZ.REAL_DATE," & vbCrLf _
                & "SPTCWRXZ.REASON_CODE" & vbCrLf _
                & " from ARTCUST2,ARTCUST1,SOTSELL1,SOTSREG1,SOTSVPS1,SOTTCLS1,SPTCWRXZ" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
                & "   and SOTSVPS1.VP_CODE (+) = SOTSREG1.VP_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and SPTCWRXZ.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
                & "   and SPTCWRXZ.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & "   and (ARTCUST1.CUST_CODE = 'ULTA' OR ARTCUST1.CUST_CODE = 'SEPHORA' OR ARTCUST1.TRADE_CLASS_CODE IN ('DPT','SPC'))" & vbCrLf _
                & "   and NVL(ARTCUST2.CUST_STORE_STATUS,'?') = 'A'" & vbCrLf _
                & "   and (ARTCUST1.CUST_CODE = 'SEPHORA' OR (NVL(ARTCUST2.CUST_DC_IND,'0') <> '1' OR ARTCUST2.CUST_DC_ALIAS IS NOT NULL))" & vbCrLf _
                & "   and ARTCUST2.SELL_CODE is Not Null"

            ' not sure why this should be temp
            '   Stop ' temp code sell code is not null

            ASCDATA1.ExecuteSQL() ' Create Current Data Set - these are the records that will become the new SPTCWRXZ

            ASCMAIN1.sql = "Delete from " & SPTCWRXY
            ASCDATA1.ExecuteSQL() ' Clear out Diffs

            ASCMAIN1.sql = "Insert into " & SPTCWRXY & " Select * from SPTCWRXZ where (CUST_CODE, CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from SPTCWRXZ minus Select CUST_CODE, CUST_STORE_NO from " & SPTCWRXZ & ")"
            ASCDATA1.ExecuteSQL() ' Put all T's into Diff - these are stores currently in SPTCWRXZ that are not in current data set

            ASCMAIN1.sql = "Update " & SPTCWRXY & " Set TERM_DATE = TRUNC(SYSDATE), REASON_CODE = 'T'"
            ASCDATA1.ExecuteSQL() ' T's are done.  T's will never make it to SPTCWRXZ

            ASCMAIN1.sql = "Insert into " & SPTCWRXY & " (Select * from " & SPTCWRXZ & " minus Select * from SPTCWRXZ)"
            ASCDATA1.ExecuteSQL() ' Insert all records in new dataset which are either new or changed from their versions in current data set

            ASCMAIN1.sql = "Update " & SPTCWRXY & " Set REASON_CODE = '?' where REASON_CODE is Not Null and REASON_CODE <> 'T'"
            ASCDATA1.ExecuteSQL() ' Mark all diff records not previously set to T as ? - as long as their reason code is not null - because

            ASCMAIN1.sql = "Update " & SPTCWRXY & " Set INIT_DATE = TRUNC(SYSDATE), REASON_CODE = 'N' where REASON_CODE is Null"
            ASCDATA1.ExecuteSQL() ' Set all diff records to REASON_CODE = 'N' if they do not have a current value for reason code - this means that they are not in current data set

            ASCMAIN1.sql = "" _
                & "Select CUST_CODE, CUST_STORE_NO, SELL_CODE, REGION_CODE, VP_CODE from " & SPTCWRXZ & " minus " & vbCrLf _
                & "Select CUST_CODE, CUST_STORE_NO, SELL_CODE, REGION_CODE, VP_CODE from SPTCWRXZ"
            ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO from (" & ASCMAIN1.sql & ")"
            ASCMAIN1.sql = "Update " & SPTCWRXY & " Set REAL_DATE = TRUNC(SYSDATE), REASON_CODE = 'R' where (CUST_CODE,CUST_STORE_NO) in (" & ASCMAIN1.sql & ") and REASON_CODE = '?'"
            ASCDATA1.ExecuteSQL() ' Set all diff records to REASON_CODE= 'R' if any of the alignment codes have changed

            ASCMAIN1.sql = "Update " & SPTCWRXY & " Set LAST_DATE = TRUNC(SYSDATE), REASON_CODE = 'C' where REASON_CODE = '?'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is Select * from " & SPTCWRXY & " where REASON_CODE <> 'T';" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SPTCWRXZ & vbCrLf _
                & "    Set INIT_DATE = R1.INIT_DATE, LAST_DATE = R1.LAST_DATE, REAL_DATE = R1.REAL_DATE, TERM_DATE = R1.TERM_DATE, REASON_CODE = R1.REASON_CODE" & vbCrLf _
                & "    where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Sub Create_Store_File()

        FILENAME = "IPLB_STORE_" & Format(Now.Date, "MMddyyyy") & "_" & CWRX_NO & ".CSV"
        Dim R As Int64 = 0
        Dim TABLE_NAME As String = "SPTCWRXY"
        Dim order_by As String = "CUST_CODE,CUST_STORE_NO"
        Dim quo As String = Chr(34)
        Dim sep As String = ","

        Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & "\" & FILENAME)

            Dim sqlw As String = "CUST_CODE = 'BOSCOVS'" ' for testing
            sqlw = ""

            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw, order_by)
                R += 1

                Dim RECORD As String = ""

                For Each dcol As DataColumn In row.Table.Columns
                    Dim DATUM As String = row.Item(dcol.ColumnName) & ""
                    If dcol.DataType.ToString = "System.String" Then
                        If DATUM = "" Then
                            DATUM = " "
                        End If
                        DATUM = quo & DATUM & quo
                    ElseIf dcol.DataType.ToString = "System.DateTime" Then
                        If DATUM <> "" Then
                            DATUM = Format(row.Item(dcol.ColumnName), "yyyyMMdd")
                        Else
                            DATUM = " "
                        End If
                    Else
                        DATUM = CStr(Val(DATUM))
                    End If
                    RECORD &= sep & DATUM
                Next

                sw.WriteLine(Mid(RECORD, 2))
            Next
        End Using
    End Sub

    Private Sub grdSPTCWRXX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTCWRXX.DoubleClickRow
        Dim CWRX_NO As String = e.Row.Cells("CWRX_NO").Value & ""
        Dim FILENAME As String = e.Row.Cells("FILENAME").Value & ""
        Dim FILENAME_temp As String = ASCMAIN1.Folders("Temp") & FILENAME
        Try
            My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Archive") & "COWORX\" & FILENAME, FILENAME_temp, True)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Copy File")
        End Try
        Show_Document(FILENAME_temp)
    End Sub

End Class