Public Class SPFCWRXB

    Dim SPTCWXBD As String
    Dim SPTCWXBI As String
    Dim rowSPTCWRXX As DataRow
    Dim CWRX_NO As String
    Dim FILENAME As String
    Dim Budget_Seasons As String
    Dim SEASON_YEAR As String
    Dim SEASON_TYPE As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_Work_Tables(True)

        With dst
            Create_TDA(.Tables.Add, "GLTPARM3", "*", 0)

            ASCMAIN1.sql = "Select * from SPTCWRXX where CWRX_EXPORT_TYPE = 'B'"
            Create_TDA(.Tables.Add, "SPTCWRXX", "**", 0)

            ASCMAIN1.sql = "Select * from " & SPTCWXBD
            Create_TDA(.Tables.Add, "SPTCWXBD", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select * from " & SPTCWXBI
            Create_TDA(.Tables.Add, "SPTCWXBI", "**", 0, False, "", 0)
        End With

        grdSPTCWXBD.DataSource = dst.Tables("SPTCWXBD")
        grdSPTCWRXX.DataSource = dst.Tables("SPTCWRXX")

        Create_Summary(grdSPTCWRXX, "CWRX_NO", "Count")

        Create_Summary(grdSPTCWXBD, "CUST_CODE", "Count")
        Create_Summary(grdSPTCWXBD, "BUDGET")
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

                If chkInitialize.Checked Then
                    MsgBox("Please be aware that you have created a Budget file with ALL budgets, not just those that were added/changed since the last export." & vbCrLf & vbCrLf & "Please compare totals to another source.", MsgBoxStyle.OkOnly, "Verification")
                End If

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

                ' .Groups("Options").Enabled = Not tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(chkInitialize, ScreenMode)

        grdSPTCWRXX.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTCWRXX", "SPTCWXBD", "SPTCWXBI"}
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
        rowSPTCWRXX.Item("CWRX_EXPORT_TYPE") = "B"
        rowSPTCWRXX.Item("CWRX_DATE") = Now.Date
        dst.Tables("SPTCWRXX").Rows.Add(rowSPTCWRXX)

        EnforceConstraints(False)

        Fill_Records("GLTPARM3")

        ' Create Budget Seasons as Current Season and Next Season
        Dim YYYYWW As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -4)
        Dim row As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
        Dim YYYYMM As String = row.Item("YYYYMM")
        SEASON_YEAR = Mid(YYYYMM, 1, 4)
        SEASON_TYPE = IIf(Mid(YYYYMM, 5, 2) > 6, "F", "S")

        Budget_Seasons = "'" & SEASON_YEAR & SEASON_TYPE & "'"
        If Mid(Budget_Seasons, 6, 1) = "F" Then
            Budget_Seasons &= ",'" & Val(SEASON_YEAR) + 1 & "S'"
        Else
            Budget_Seasons &= ",'" & SEASON_YEAR & "F'"
        End If

        Create_Work_Tables(False)

        Fill_Records("SPTCWXBD")
        Sort_grdColumns(grdSPTCWXBD, "CUST_CODE,CUST_STORE_NO")

        Create_Budget_File()
        rowSPTCWRXX.Item("FILENAME") = FILENAME

        rowSPTCWRXX.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSPTCWRXX.Item("INIT_DATE") = DATETIME_STAMP

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now sftp'ing the file")

        BeginTrans()

        ' Delta - this is what was sent to CoWorx
        ASCMAIN1.sql = "Insert into SPTCWXBD Select * from " & SPTCWXBD
        ASCDATA1.ExecuteSQL()

        ' Image - this is what we will compare to next time
        ASCMAIN1.sql = "Delete from SPTCWXBI"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into SPTCWXBI Select * from " & SPTCWXBI
        ASCDATA1.ExecuteSQL()

        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "\" & FILENAME, ASCMAIN1.Folders("Archive") & "\COWORX\" & FILENAME)
        If dst.Tables("SPTCWXBD").Rows.Count > 0 Then
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
        Load_Popup_Menu(grdSPTCWXBD, "SS", "Show Filter", "Show GroupBox")
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

    Sub Create_Budget_File()

        FILENAME = "IPLB_Budget_" & Format(Now.Date, "MMddyyyy") & "_" & CWRX_NO & ".CSV"
        If chkInitialize.Checked Then
            FILENAME = Replace(FILENAME, "_Budget_", "_Full_Budget_")
        End If

        Dim R As Int64 = 0
        Dim TABLE_NAME As String = "SPTCWXBD"
        Dim order_by As String = "SEASON_CODE,CUST_CODE,CUST_STORE_NO,WEEK_NO"
        Dim quo As String = Chr(34)
        Dim sep As String = ","

        Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & "\" & FILENAME)

            Dim sqlw As String = "CUST_CODE = 'BOSCOVS'" ' for testing
            sqlw = ""

            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw, order_by)
                R += 1

                Dim CUST_CODE As String = row.Item("CUST_CODE")
                Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                Dim SELL_CODE As String = row.Item("SELL_CODE")
                Dim REGION_CODE As String = row.Item("REGION_CODE")
                Dim VP_CODE As String = row.Item("VP_CODE")
                Dim YYYYWW As String = row.Item("YYYYWW")

                Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
                ' Dim YP As String = row.Item("OPS_YYYYPP")
                Dim WEEK_END_DATE As String = Format(rowGLTPARM3.Item("WEEK_END_DATE"), "yyyy-MM-dd")

                Dim CHECKBOOK As String = row.Item("CHECKBOOK")
                Dim BUDGET As Decimal = Val(row.Item("BUDGET") & "")

                Dim RECORD As String = ""

                RECORD &= sep & quo & "IMPORTFORECAST" & quo
                RECORD &= sep & quo & "iplb" & quo
                RECORD &= sep & quo & "SAM04" & quo
                RECORD &= sep & quo & "IPLB-" & VP_CODE & "-" & REGION_CODE & "-" & SELL_CODE & "-" & CUST_CODE & CUST_STORE_NO & quo
                RECORD &= sep & quo & CHECKBOOK & quo
                ' RECORD &= sep & quo & Mid(YP, 1, 4) & "-" & Mid(YP, 5, 2) & "-" & "15" & quo
                RECORD &= sep & quo & WEEK_END_DATE & quo
                RECORD &= sep & BUDGET

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

    Sub Create_Work_Tables(initialize As Boolean)

        If initialize Then
            ASCMAIN1.sql = "Select * from SPTCWXBD where ROWNUM < 1"
            SPTCWXBD = ASCMAIN1.Temp_Table()
            'ASCDATA1.ExecuteSQL("Alter Table " & SPTCWXBD & " Add Primary Key (CUST_CODE,CUST_STORE_NO,CHECKBOOK,SEASON_CODE,WEEK_NO,SELL_CODE,REGION_CODE,VP_CODE)")

            ASCMAIN1.sql = "Select * from SPTCWXBI where ROWNUM < 1"
            SPTCWXBI = ASCMAIN1.Temp_Table()
            'ASCDATA1.ExecuteSQL("Alter Table " & SPTCWXBI & " Add Primary Key (CUST_CODE,CUST_STORE_NO,CHECKBOOK,SEASON_CODE,WEEK_NO,SELL_CODE,REGION_CODE,VP_CODE)")
        Else
            ' I IS THE CURRENT MASTER (Image)
            ' D IS THE CHANGES (Delta)

            ' Clear out the Temp Image
            ASCMAIN1.sql = "Delete from " & SPTCWXBI
            ASCDATA1.ExecuteSQL()

            ' Create Current Data Set in the Temp Image table
            ' these are the records that will become the new Image


            ' MIGHT NEED TO USE 27 INSTEAD OF 26 WHEN FOLLOWING A 53 WEEK YEAR

            ASCMAIN1.sql = "Insert into " & SPTCWXBI & " " & vbCrLf _
                & "Select SPTMXWS2.CUST_CODE, SPTMXWS2.CUST_STORE_NO, SPTMXWS2.CHECKBOOK " & vbCrLf _
                & ", SPTMXWS2.SEASON_CODE, SPTMXWS2.WEEK_NO" & vbCrLf _
                & ", Substr(SPTMXWS2.SEASON_CODE,1,4) || TRIM(TO_CHAR(DECODE(SUBSTR(SEASON_CODE,5,1),'S',0,26) + SPTMXWS2.WEEK_NO,'00')) as YYYYWW" & vbCrLf _
                & ", SYSDATE WEEK_END_DATE" & vbCrLf _
                & ", SOTSELL1.SELL_CODE, SOTSREG1.REGION_CODE, SOTSREG1.VP_CODE" _
                & " , Sum (SPTMXWS2.TY_SPEND) BUDGET " & vbCrLf _
                & " from SPTMXWS2,ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = SPTMXWS2.CUST_CODE " & vbCrLf _
                & " and ARTCUST2.CUST_STORE_NO = SPTMXWS2.CUST_STORE_NO " & vbCrLf _
                & " and SOTSELL1.SELL_CODE = ARTCUST2.SELL_CODE " & vbCrLf _
                & " and SOTSREG1.REGION_CODE = SOTSELL1.REGION_CODE " & vbCrLf _
                & " and SPTMXWS2.WEEK_NO between 1 and 27" & vbCrLf _
                & IIf(Budget_Seasons = "",
                      " and ROWNUM < 1",
                      " and SPTMXWS2.SEASON_CODE in (" & Budget_Seasons & ") and WEEK_NO <> 27") & vbCrLf _
                & " group by SPTMXWS2.CUST_CODE, SPTMXWS2.CUST_STORE_NO, SPTMXWS2.CHECKBOOK" & vbCrLf _
                & ", SPTMXWS2.SEASON_CODE, SPTMXWS2.MONTH_NO, SPTMXWS2.WEEK_NO" & vbCrLf _
                & ", SOTSELL1.SELL_CODE, SOTSREG1.REGION_CODE, SOTSREG1.VP_CODE" & vbCrLf _
                & " having Sum (SPTMXWS2.TY_SPEND) <> 0"
            ASCDATA1.ExecuteSQL()

            'If chkInitialize.Checked Then
            '    ' when level-setting to initialize CoWorx, do not create previous records table so that all budgets are sent over
            'Else
            '    ASCDATA1.ExecuteSQL()
            'End If


            ASCMAIN1.sql = "Update " & SPTCWXBI & " Set WEEK_END_DATE = NULL"
            ASCDATA1.ExecuteSQL()

            ' NOTE THAT YYYYWW IS NOT A RETAIL WEEK
            ' WEEK 01 IN SPTMXWS2 IS 1ST WEEK OF JAN
            ' SO NEED TO SUBTRACT 4 TO GET TO RETAIL WEEK
            ' EXCEPT WHEN THE LAST YEAR IS A 53 WEEK YEAR, IN WHICH CASE YOU NEED TO SUBTRACT 5

            Dim WEEKS_OFFSET As Integer = 4

            If SEASON_TYPE = "S" Then
                ' note that the SEASON_TYPE is inferred from the system date compared to the retail calendar
                ' 2017 was a 53 week year
                ' on 08/03/2018, SEASON_TYPE was S (because the Fall season in the retail calendar began on 08/04)
                ' and the correction below really only pertained to Spring-2018 - and it should not have been applied to Fall
                ' so the Fall 2018 budgets, uploaded prior to 08/04, all were off by 1 week (see CM email 08/07/2018)
                ' all I had her do was to rerun the budget export, and it was corrected since the run date was 08/08 - ie Fall
                ' this confusing mess might also be an issue in the year following the next 53 week year
                ' the problem is that the export file will contain both fall and spring budgets at tomes
                ' so we should probably apply this adjustment to the spring records, and not the fall records, but not sure about that, so I am leaving it be for now

                ASCMAIN1.sql = "Select COUNT (*) from GLTPARM3 where YYYYWW LIKE '" & Format(Val(SEASON_YEAR) - 1, "0000") & "%'"
                Dim WEEKS_LY As Integer = Val(ASCDATA1.GetDataValue)
                If WEEKS_LY = 53 Then
                    WEEKS_OFFSET = 5

                    MsgBox($"Note - {SEASON_YEAR} is a 53 week year - we might need to make an adjustment", MsgBoxStyle.OkOnly, "Please call ABS to review")
                End If
            End If



            ASCMAIN1.sql = "Select Distinct YYYYWW from " & SPTCWXBI
            For Each rowYW As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                Dim YYYYWW As String = rowYW.Item(0)
                Dim YYYYWW_new As String = ASCMAIN1.Week_Calc(YYYYWW, -WEEKS_OFFSET)
                Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW_new)
                Dim WEEK_END_DATE As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                ASCMAIN1.sql = "Update " & SPTCWXBI & vbCrLf _
                    & " Set YYYYWW = '" & YYYYWW_new & "', WEEK_END_DATE = '" & Format(WEEK_END_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
                    & " where YYYYWW = '" & YYYYWW & "'"
                ASCDATA1.ExecuteSQL()
            Next



            ASCMAIN1.sql = "Delete from " & SPTCWXBD
            ASCDATA1.ExecuteSQL() ' Clear out Diffs

            ' Create 0 records for Budgets in old file that are new Diffs

            ASCMAIN1.sql = "Insert into " & SPTCWXBD & vbCrLf _
                & "Select D.*, '" & CWRX_NO & "' CWRX_NO from (" & vbCrLf _
                & "Select * from SPTCWXBI where SEASON_CODE in (" & Budget_Seasons & ") and WEEK_NO <> 27" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select * from " & SPTCWXBI & vbCrLf _
                & ") D"
            If chkInitialize.Checked Then
                ' NOT NEC
            Else
                ASCDATA1.ExecuteSQL()
            End If


            ASCMAIN1.sql = "Delete from " & SPTCWXBD & vbCrLf _
                & " where (CUST_CODE, CUST_STORE_NO, CHECKBOOK, SEASON_CODE, YYYYWW)" & vbCrLf _
                & " in (Select CUST_CODE, CUST_STORE_NO, CHECKBOOK, SEASON_CODE, YYYYWW from " & SPTCWXBI & ")"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SPTCWXBD & " Set BUDGET = 0"
            ASCDATA1.ExecuteSQL()

            ' Create new Diffs

            'ASCMAIN1.sql = "Insert into " & SPTCWXBD & vbCrLf _
            '        & "Select D.*, '" & CWRX_NO & "' CWRX_NO from (" & vbCrLf _
            '        & "Select * from " & SPTCWXBI & vbCrLf _
            '        & " minus " & vbCrLf _
            '        & "Select * from SPTCWXBI where SEASON_CODE in (" & Budget_Seasons & ")" & vbCrLf _
            '        & ") D"

            ASCMAIN1.sql = "Insert into " & SPTCWXBD & vbCrLf _
                & "Select D.*, '" & CWRX_NO & "' CWRX_NO from (" & vbCrLf _
                & "Select * from " & SPTCWXBI & vbCrLf
            If chkInitialize.Checked Then
            Else
                ASCMAIN1.sql &= " minus " & vbCrLf _
                & "Select * from SPTCWXBI where SEASON_CODE in (" & Budget_Seasons & ") and WEEK_NO <> 27"
            End If

            ASCMAIN1.sql &= ") D"

            ASCDATA1.ExecuteSQL()

        End If

    End Sub
End Class