Imports System.Math
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Public Class ASFSRPTM

#Region "Declarations"

    Protected SEQs As Integer       ' Number of Columns Selected to Sort
    Protected FORM_NAME As String   ' Form Object Name for Report (Me.Name)
    Protected REPORT_NO As String   ' Each Report Instance gets its own REPORT_NO
    Protected START_TIME As Date    ' Date/Time Proceed was Clicked

    Protected SET_ID As String      ' Set used during this Execution
    Protected SET_DESC As String    ' Description of SET_ID
    Protected SUBT As String        ' Report Sub-Title - defaulting to SET_DESC

    Protected LIST_CODE As String = ""  ' Working variable for List of Codes
    Protected LIST_DESC As String = ""  ' Corresponding Description

    Protected aRC As String = Chr(142)         ' Token used for PB Report Summaries/Recaps
    Protected G1thru9 As String = "G1,G2,G3,G4,G5,G6,G7,G8,G9"

    Protected RWU As String         ' P = Print Report, R = Report w/Update, U = Update (no report), N = No Update (ie, just like P); appended with 0 = No Eligible Records
    Protected RPT As String         ' Crystal Report (.RPT) File Name
    Protected RPT_TITLE As String   ' Title for RPT (defaults to menu description)

    Protected YYYYPP As String      ' Period selected
    Protected YPD As String         ' Year & Period Description

    Protected Page0 As New ArrayList    ' Page 0 descriptive items
    Protected sql_SELECT_cols As String     ' PB Columns used in the Select List
    Protected sql_GROUP_BY_cols As String   ' PB Columns used in the Group By
    Protected sql_WHERE As String           ' PB where clause
    Protected sql_TABLE_NAMEs As String     ' PB Tables
    Protected sql_JOIN As String            ' PB Join
    Protected sql_TABLE_NAME As String      ' PB Primary Table for Data Source

    Protected COLUMN_NAMEs As New ArrayList         ' PB Column Names
    Protected COLUMN_CAPTIONs As New ArrayList      ' PB Column Captions
    Protected GROUP_ALL_OTHERSs As New ArrayList    ' Whether others should be grouped
    Dim PAGE_BREAKs As String                       ' PB Level Page Breaks
    Protected xErrMsg As String     ' if <> "" (error); exit Main Process in Sub Proceed
    Protected COLUMN_NAME_first As String = ""  ' Forced First Column of a PB Report
    Protected COLUMN_NAME_last As String = ""   ' Forced Last Column of a PB Report
    Protected COLUMN_NAMEs_appended As String = ""
    Protected COLUMN_NAME_RECAP_ROW_NO As String = ""
    Protected tblASTDSQLH As DataTable
    Protected tblASTDSQLS As DataTable
    Protected PB_Report As Boolean = False
    Protected tblASTRECAP As DataTable
    Protected Recap_Report As Boolean = False
    Protected ASTSRPT1 As String
    Protected ASTSRPT1_sum_columns As String
    Protected ASTSRPT1_sql_sum As String
    Protected COLUMN_NAME_sum As New Dictionary(Of String, String)

    Public ExportFormatDefault As String = "RPT"
    Public ExportFilenameDefault As String = ""
    Public ArchiveReportsDefault As Boolean = True
    Public NumExecutions As Integer = 0

    Protected tblASTDSQL1 As DataTable
    Protected tblASTDSQL1_copy As DataTable
    Protected COLUMN_NAME_by_Lvl() As String
    Protected COLUMN_CAPTION_by_Lvl() As String
    Protected G_by_Lvl() As Integer
    Protected COLUMN_NAME_sum_first As String
    Protected DATA_TYPEs() As String
    Protected Totals_Row() As Decimal = Nothing
    Protected Totals_All() As Decimal = Nothing

    Protected sql As String
    Protected RYP As String
    Protected RYPLEGEND As String
    Protected RYP0 As String
    Protected RYPLEGEND0 As String
    Protected RYP1 As String
    Protected RYPLEGEND1 As String

    Protected RYW As String
    Protected RYWLEGEND As String
    Protected RYW0 As String
    Protected RYWLEGEND0 As String
    Protected RYW1 As String
    Protected RYWLEGEND1 As String

    Protected tblASTSPRF1_clone As New DataTable
    Protected tblASTGROUP As DataTable
    Protected tblASTSRPT0 As DataTable
    Protected tblASTSRPT1 As DataTable

    Protected tblASTROPT1 As New DataTable
    Protected tblASTROPT4 As New DataTable

    Public tblASTDSQLA As DataTable
    Protected tblASTDSQLB As New DataTable
    Protected tblASTDSQLC As New DataTable
    Protected tblASTDSQLD As New DataTable
    Protected tblASTDSQLE As New DataTable
    Protected tblASTDSQLF As New DataTable
    Protected tblASTDSQLJ As New DataTable

    Private USER_GROUP_IDs As New List(Of String)
    Private updateSuccessful As Boolean = True

#End Region

#Region "Initialization"

    Private Sub ASFSRPTM_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        If tblASTDSQLH IsNot Nothing Then
            tblASTDSQLH.Dispose()
        End If
        If tblASTDSQLS IsNot Nothing Then
            tblASTDSQLS.Dispose()
        End If
        If tblASTRECAP IsNot Nothing Then
            tblASTRECAP.Dispose()
        End If
        If tblASTDSQL1 IsNot Nothing Then
            tblASTDSQL1.Dispose()
        End If
        If tblASTDSQL1_copy IsNot Nothing Then
            tblASTDSQL1_copy.Dispose()
        End If
        If tblASTSPRF1_clone IsNot Nothing Then
            tblASTSPRF1_clone.Dispose()
        End If
        If tblASTGROUP IsNot Nothing Then
            tblASTGROUP.Dispose()
        End If
        If tblASTSRPT0 IsNot Nothing Then
            tblASTSRPT0.Dispose()
        End If
        If tblASTSRPT1 IsNot Nothing Then
            tblASTSRPT1.Dispose()
        End If
        If tblASTROPT1 IsNot Nothing Then
            tblASTROPT1.Dispose()
        End If
        If tblASTROPT4 IsNot Nothing Then
            tblASTROPT4.Dispose()
        End If
        If tblASTDSQLA IsNot Nothing Then
            tblASTDSQLA.Dispose()
        End If
        If tblASTDSQLB IsNot Nothing Then
            tblASTDSQLB.Dispose()
        End If
        If tblASTDSQLC IsNot Nothing Then
            tblASTDSQLC.Dispose()
        End If
        If tblASTDSQLD IsNot Nothing Then
            tblASTDSQLD.Dispose()
        End If
        If tblASTDSQLE IsNot Nothing Then
            tblASTDSQLE.Dispose()
        End If
        If tblASTDSQLF IsNot Nothing Then
            tblASTDSQLF.Dispose()
        End If
        If tblASTDSQLJ IsNot Nothing Then
            tblASTDSQLJ.Dispose()
        End If

        Page0 = Nothing
        COLUMN_NAMEs = Nothing
        COLUMN_CAPTIONs = Nothing
        GROUP_ALL_OTHERSs = Nothing
        tblASTDSQLH = Nothing
        tblASTDSQLS = Nothing
        tblASTRECAP = Nothing
        COLUMN_NAME_sum = Nothing

        tblASTDSQL1 = Nothing
        tblASTDSQL1_copy = Nothing
        tblASTSPRF1_clone = Nothing
        tblASTGROUP = Nothing
        tblASTSRPT0 = Nothing
        tblASTSRPT1 = Nothing
        tblASTROPT1 = Nothing
        tblASTROPT4 = Nothing
        tblASTDSQLA = Nothing
        tblASTDSQLB = Nothing
        tblASTDSQLC = Nothing
        tblASTDSQLD = Nothing
        tblASTDSQLE = Nothing
        tblASTDSQLF = Nothing
        tblASTDSQLJ = Nothing

        USER_GROUP_IDs = Nothing

    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If
        FORM_NAME = Me.Name

        UltraTabControl1.Tabs("Distribution").Visible = False
        'Me.Load_Popup_Menus() ' THIS WINDS UP CALLING REPORT FORM METHOD
        Load_Popup_Menu(grd, "BB", "Retreive List", "Save As List")
        SplitContainer4.Panel2Collapsed = True

        If ASCMAIN1.Running_in_VS Then
        Else
            cmdOLAP.Visible = False
            cmdPivot.Visible = False
        End If


    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        If Me.Name <> "ASFSRPTM" Then
            Call Setup_grdSetup()
            Call Setup_grdASTRECAP()
            Call Initialize_Form()
            Call Save_Settings("0000000000")
        End If

        grdSetup.DisplayLayout.Bands(0).Columns("GROUP_ALL_OTHERS").Hidden = True

        If ASCMAIN1.JOB_STREAM_XNO.Length > 0 _
                AndAlso ASCMAIN1.JOB_STREAM_CODE.Length > 0 _
                AndAlso ASCMAIN1.JOB_STREAM_FORM_NAME = MENU_ITEM_OBJECT Then
            ' get the settings then run the report
            Dim sql As String = "Select * From ASTJOBM2 Where JOB_STREAM_CODE = :PARM1 and JOB_STREAM_LNO = :PARM2"
            Dim tblASTJOBM2 As DataTable = ASCDATA1.GetDataTable(sql, "", "VN", New Object() {ASCMAIN1.JOB_STREAM_CODE, ASCMAIN1.JOB_STREAM_LNO})
            If tblASTJOBM2.Rows.Count > 0 Then
                Dim rowASTJOBM2 As DataRow = tblASTJOBM2.Rows(0)
                SET_ID = rowASTJOBM2.Item("SET_ID") & String.Empty

                System.Threading.Thread.Sleep(1000)
                sql = "Insert Into ASTJOBM3"
                sql &= " (JOB_STREAM_XNO, JOB_STREAM_CODE, JOB_STREAM_LNO, REPORT_ID, SET_ID, MENU_ID, INIT_DATE, LAST_DATE, SUCCESS)"
                sql &= " Values "
                sql &= "( '" & ASCMAIN1.JOB_STREAM_XNO & "'"
                sql &= ", '" & ASCMAIN1.JOB_STREAM_CODE & "'"
                sql &= ", " & rowASTJOBM2.Item("JOB_STREAM_LNO")
                sql &= ", '" & MENU_ITEM_OBJECT & "'"
                sql &= ", '" & SET_ID & "'"
                sql &= ", '" & MENU_ID & "'"
                sql &= ", SYSDATE, SYSDATE, '0')"
                ASCDATA1.ExecuteSQL(sql)

                If SET_ID.Length > 0 Then
                    sql = "SET_ID = '" & SET_ID & "'"
                    If tblASTROPT1.Select(sql).Length = 0 Then
                        Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
                        Exit Sub
                    End If
                Else
                    SET_ID = "0000000000"
                End If
                Retrieve_Settings()

                Try
                    UltraTabControl1.Tabs("Other Run-Time Options").Selected = True
                    System.Threading.Thread.Sleep(1000)
                Catch ex As Exception

                End Try

                System.Threading.Thread.Sleep(1000)
                Click_Command("Proceed")
                System.Threading.Thread.Sleep(1000)

                ' Perfrom Update
                If rowASTJOBM2.Item("UPDATE_REPORT") & String.Empty = "1" Then
                    With UltraExplorerBar1
                        If .Groups("Update Controls").Items("Update").Settings.Enabled = DefaultableBoolean.True Then
                            Dim key As String = .Groups("Update Controls").Items("Update").Key
                            Click_Command(key)
                            'Else
                            '    updateSuccessful = False
                        End If
                    End With
                End If

                sql = "UPDATE ASTJOBM3 SET SUCCESS = '" & IIf(updateSuccessful, "1", "0") & "'"
                sql &= " , LAST_DATE = SYSDATE"
                sql &= " WHERE JOB_STREAM_XNO = '" & ASCMAIN1.JOB_STREAM_XNO & "'"
                sql &= " AND JOB_STREAM_CODE = '" & ASCMAIN1.JOB_STREAM_CODE & "'"
                sql &= " AND JOB_STREAM_LNO =  " & rowASTJOBM2.Item("JOB_STREAM_LNO")
                ASCDATA1.ExecuteSQL(sql)

                With UltraExplorerBar1
                    If .Groups("Update Controls").Items("Done").Settings.Enabled = DefaultableBoolean.True Then
                        Dim key As String = .Groups("Update Controls").Items("Done").Key
                        Click_Command(key)
                    End If
                End With



                Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
            End If
        End If

    End Sub

    Sub Initialize_Form()
        ASCMAIN1.sql = "Select * from ASTROPT1 where FORM_NAME = '" & FORM_NAME & "'"
        tblASTROPT1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTROPT1", 2)
        grdASTROPT1.DataSource = tblASTROPT1

        'ASCMAIN1.sql = "Select ASTROPT4.*, ASTUSER1.USER_NAME, ASTUSER1.USER_EMAIL, ASTUSER1.USER_STATUS" _
        '& " from ASTROPT4,ASTUSER1 where ASTUSER1.USER_ID = ASTROPT4.USER_ID" _
        '& " and ASTROPT4.FORM_NAME = :PARM1 and ASTROPT4.SET_ID = :PARM2"
        ASCMAIN1.sql = "Select ASTROPT4.*, ASTUSER1.USER_NAME, ASTUSER1.USER_EMAIL, ASTUSER1.USER_STATUS" _
        & " from ASTROPT4,ASTUSER1 where ASTUSER1.USER_ID = ASTROPT4.USER_ID" _
        & " and ASTROPT4.FORM_NAME = '" & FORM_NAME & "' and SET_ID = '0000000000'"

        'ASCMAIN1.sql = "Select * from ASTROPT4 where FORM_NAME = '" & FORM_NAME & "' and SET_ID = '0000000000'"
        tblASTROPT4 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTROPT4", 3)
        grdASTROPT4.DataSource = tblASTROPT4

        Call Show_Settings()
        Call Mode_Settings(False)

        Call Load_Popup_Menu(tvwDQ, "B", "Export to Excel")
    End Sub
#End Region

#Region "ABS Standard Routines"

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        If EntryMode = "Complete" Or EntryMode = "Update" Then
            If EntryMode = "Update" Then
                ScreenMode = tf
            End If
        Else
            Call Set_ScreenMode_Base(tf)
        End If

        With UltraExplorerBar1

            .Groups("Main Controls").Items("Default Settings").Settings.Enabled = not_iScreenMode
            .Groups("Main Controls").Items("Proceed").Settings.Enabled = not_iScreenMode
            ' .Groups("Main Controls").Items("Execution History").Settings.Enabled = not_iScreenMode
            .Groups("Data Query").Visible = ASCMAIN1.USER_ID = "wjzz" AndAlso (EntryMode = "" And PB_Report And Not tf And dst.Tables.Count > 0 And (UltraTabControl1.SelectedTab.Key = "Data Query"))
            .Groups("Data Grid").Visible = ASCMAIN1.USER_ID = "wjzz" AndAlso (EntryMode = "" And PB_Report And Not tf And dst.Tables.Count > 0 And (UltraTabControl1.SelectedTab.Key = "Data Grid"))

            UltraExplorerBar1.Groups("Update Controls").Visible = (EntryMode = "Complete" Or EntryMode = "Update")

            If tf Then
                If EntryMode = "Proceed" Then
                    SplitContainer1.Enabled = False
                    .Groups("Saved Settings").Visible = False

                    .Groups("Update Controls").Items("Update").Settings.AppearancesSmall.Appearance.ForeColorDisabled = Color.Empty
                    .Groups("Update Controls").Items("Update").Text = "Update"

                Else
                    If EntryMode = "Update" Then
                        .Groups("Update Controls").Items("Update").Settings.Enabled = not_iScreenMode
                        .Groups("Update Controls").Items("Cancel").Settings.Enabled = not_iScreenMode
                        .Groups("Update Controls").Items("Done").Settings.Enabled = iScreenMode

                    Else
                        If RWU = "R" And EntryMode <> "Update" Then
                            .Groups("Update Controls").Items("Update").Settings.Enabled = iScreenMode
                            .Groups("Update Controls").Items("Cancel").Settings.Enabled = iScreenMode
                            .Groups("Update Controls").Items("Update").Text = "Update"

                        Else
                            ' if there was a problem and the Update option needs to remain disabled then show it in red

                            If xErrMsg <> "" And RWU <> "N0" Then ' If .Groups("Update Controls").Items("Update").Visible Then
                                .Groups("Update Controls").Items("Update").Settings.AppearancesSmall.Appearance.ForeColorDisabled = Color.Red
                                .Groups("Update Controls").Items("Update").Text = "Update Disabled"
                            End If
                            .Groups("Update Controls").Items("Done").Settings.Enabled = iScreenMode
                        End If
                    End If

                End If

            Else
                SplitContainer1.Enabled = True
                .Groups("Saved Settings").Visible = True
                .Groups("Update Controls").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Update Controls").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Update Controls").Items("Done").Settings.Enabled = iScreenMode
            End If

            If EntryMode = "Proceed" Then
            Else
                ' not sure about these yet
                .Groups("Output Controls").Visible = False
                .Groups("Output Controls").Items("Excel").Settings.Enabled = iScreenMode
                .Groups("Output Controls").Items("email Report").Settings.Enabled = iScreenMode
            End If
        End With

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        If EntryMode = "Complete" Then
        Else
            Call Set_ScreenMode_Special(tf) ' not sure if this goes on the Then or on the Else
        End If


    End Sub

    Overridable Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Call Verify_Special_Pre(eItemKey)

        Select Case eItemKey
            Case "Default Settings"
                '                Call Check_Key_Fields()

            Case "Proceed"
                updateSuccessful = False

                If PB_Report Then
                    If COLUMN_NAME_last = "" And tblASTDSQLH.Rows.Count = 0 And tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                        EMsg &= vbCr & "You must pick at least 1 column to Sort by"
                    End If
                End If

                For Each row As DataRow In tblASTDSQLA.Select("EXCLUDE = '1' and (CODE_VALUES is Null or CODE_VALUES = '')", "")
                    EMsg &= vbCr & "You Cannot Exclude 'All' values (" & row.Item("COLUMN_CAPTION") & ")"
                Next

                If Recap_Report Then
                    tblASTRECAP.AcceptChanges()
                    If grdASTRECAP.Rows.Count = 0 Then
                        EMsg &= vbCr & "You must specify at least 1 Data Row to Show"
                    Else
                        If grdASTRECAP.Rows.AddRowModifiedByUser Then
                            EMsg &= vbCr & "You must either Add or (ESC to) Delete the Data Row Specification Started"
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("Are you Sure that you want to Cancel?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Clicking Cancel will NOT perform the Associated Update") = MsgBoxResult.No Then
                    Exit Sub
                End If

        End Select

        Call Verify_Special(eItemKey)

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Overridable Sub Verify_Special(ByVal eItemKey As String)

    End Sub

    Overridable Sub Verify_Special_Pre(ByVal eItemKey As String)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Default Settings"
                SET_ID = "0000000000"
                Call Retrieve_Settings()
                SET_ID = ""

                '                Call Mode_Settings(False)

            Case "Proceed"
                updateSuccessful = True
                If rowASTOPST1.Item("PROCEED_BEGIN") & "" <> "" Then
                    tblASTOPST1.Rows.Clear()
                    Create_ASTOPST1()
                End If

                EntryMode = "Proceed"
                xErrMsg = ""

                dst.Clear() ' MAYBE WE SHOULD BE USING ASFBASE1.Clear_dst

                HFs.Clear()
                TDAs.Clear()
                TBLs.Clear()
                TBL_SCHEMAs.Clear()
                DVWs.Clear()
                pROWs.Clear()

                'ROWs.Clear()
                'CMDs.Clear()

                If Not dst.Tables.Contains("ASTSQLX1") Then
                    ASCMAIN1.tblASTSQLX1 = Nothing
                    Create_TDA(dst.Tables.Add, "ASTSQLX1", "*")
                    ASCMAIN1.tblASTSQLX1 = dst.Tables("ASTSQLX1")
                End If



                Call Mode_Settings(True)
                Call Main_Process()

                Dim difference As TimeSpan = CDate(rowASTOPST1.Item("PROCEED_END")).Subtract(CDate(rowASTOPST1.Item("PROCEED_BEGIN")))
                Dim ELAPSED_TIME As String = Format(difference.Hours, "00") & ":" & Format(difference.Minutes, "00") & ":" & Format(difference.Seconds, "00")
                Call ASCMAIN1.Progress("Process Complete. Elapsed Time = " & ELAPSED_TIME)

                EntryMode = "Complete"
                Call Mode_Settings(True)
                If RWU = "U" Then
                    Call Click_Command("Update")
                End If

                If xErrMsg <> "" Then
                    Call ASCMAIN1.Progress(xErrMsg, "")
                End If

                If chkDQ_Only.Checked Then
                    Click_Command("Done")
                    UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Data Query")
                End If

            Case "Execution History"
                Dim f As New ASFXHST1
                f.Width = Me.Width * 0.8
                f.Height = Me.Height * 0.8
                f.RPT = True
                f.ShowDialog()
                f.Dispose()
                If ASCMAIN1.sql <> "" Then

                End If

            Case "Update"

                ' Record that Update has been clicked
                rowASTOPST1.Item("UPDATE_BEGIN") = DATETIME_STAMP ' Now + ASCMAIN1.NowTSD
                tdaASTOPST1.Update(tblASTOPST1)

                EntryMode = "Update"
                ASCMAIN1.Progress("Now Updating")

                Try
                    BeginTrans()
                    Update_Record()
                    ASCMAIN1.Progress("")
                    CommitTrans()
                    Update_Record_Post_Commit()

                Catch ex As Exception
                    updateSuccessful = False
                    Rollback("Error Occurred - Please call ABS" & vbCr & vbCr & ex.Message, ex)
                    If ASCMAIN1.Running_in_VS Then
                        Stop
                    End If
                End Try

                ' Record that the Update Process has Ended
                rowASTOPST1.Item("UPDATE_END") = Now + ASCMAIN1.NowTSD
                rowASTOPST1.Item("UPDATED") = "1"
                tdaASTOPST1.Update(tblASTOPST1)
                Dim difference As TimeSpan = CDate(rowASTOPST1.Item("UPDATE_END")).Subtract(CDate(rowASTOPST1.Item("UPDATE_BEGIN")))
                Dim ELAPSED_TIME As String = Format(difference.TotalHours, "00") & ":" & Format(difference.TotalMinutes, "00") & ":" & Format(difference.TotalSeconds, "00")

                Call ASCMAIN1.Progress("Update Complete. Elapsed Time = " & ELAPSED_TIME)
                Call Mode_Settings(True)
                'Call Mode_Settings(False)
                'Me.Close()


            Case "Cancel", "Done"
                Dim exit_when_done As Boolean = False
                If EntryMode = "Update" Then
                    exit_when_done = True
                End If
                EntryMode = ""
                Call ASCMAIN1.Progress("")
                'Call Clear_Record() - THIS IS CALLED WHEN MODES IS FALSE
                Call Mode_Settings(False)
                If exit_when_done Then
                    Me.Close()
                End If

            Case "Excel"
                Call Export_to_Excel(grdASTSRPT1)

            Case "email"
                'Call Export_to_Excel(grdASTAUDT1)
        End Select

    End Sub

    Overridable Sub Update_Record()

    End Sub

    Overridable Sub Update_Record_Post_Commit()

    End Sub

    Overridable Sub Clear_Record()
        ' NO NEED TO CLEAR OUT DST SINCE PROCEED STARTS WITH A NEW ONE - ALSO - FORM CLOSING NEEDS TO DST TO WRITE IT OUT
        'dst.Clear()
        'dst.Tables.Clear()
        If USER_GROUP_IDs IsNot Nothing Then
            USER_GROUP_IDs.Clear()
        End If
    End Sub

    Overridable Sub Proceed_Update_Special_Pre()

    End Sub

    Overridable Sub Proceed_Update_Special_Post()

    End Sub

    Overridable Sub Retrieve_Settings_Post()

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grd, "SBB", "Single Value", "Load List", "Maintain Lists")
        Load_Popup_Menu(grdASTSRPT1, "B", "Remove Captions")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Try
            If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
                e.Cancel = True
                Exit Sub
            End If

            Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
            If grd Is Nothing Then
                e.Cancel = True
                Exit Sub
            End If

            Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
            Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
                'e.Cancel = True
            Else

                Select Case e.SourceControl.Name

                End Select

            End If

            Select Case e.SourceControl.Name
                Case "grd"
                    tlb_btn = tlb_pop.Tools("Save As List")
                    tlb_btn.SharedProps.Visible = Not (grd.Rows.Count = 0)
                Case "grdASTSRPT1"
                    tlb_btn = tlb_pop.Tools("Remove Captions")

                    If grdASTSRPT1.Rows.Count = 0 Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim d As String = grdASTSRPT1.Rows(0).Cells(0).Value & ""
                        tlb_btn.SharedProps.Visible = d.Contains(":")
                    End If
            End Select

        Catch ex As Exception

        End Try
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Export to Excel"
                If tvwDQ.ActiveNode IsNot Nothing Then
                    Dim wb As New Infragistics.Documents.Excel.Workbook
                    Me.Recursive_Export_to_Excel(tvwDQ.ActiveNode, wb)
                    Call Export_to_Excel_Show(wb, Me.Text)

                    Call ASCMAIN1.Progress("")
                    Me.Cursor = Cursors.Default
                End If

            Case "Save As List"
                Dim LIST_DESC_new As String = ASCMAIN1.Get_txt_from_User("List Name", "Enter the name of this list", , 30)
                If LIST_DESC_new <> "" Then SaveAsList(LIST_DESC_new)

            Case "Retreive List"
                Retreive_List()

            Case "Remove Captions"

                Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy
                For Each row As DataRow In tbl.Select("")
                    For i As Integer = 0 To 9
                        If row.Item(i) & "" = "" Then Exit For
                        Dim d As String = Split(row.Item(i) & ":", ":")(1)
                        row.Item(i) = d
                    Next
                Next
                grdASTSRPT1.DataSource = tbl

        End Select

        Select Case grd.Name
            Case "grd"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

    'Private Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)

    'End Sub
#End Region

#Region "grdASTRECAP"
    Private Sub grdASTRECAP_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTRECAP.BeforeRowUpdate
        With grdASTRECAP.ActiveRow
            'For Each Vlist As ValueList In grdRecap.DisplayLayout.ValueLists
            '    If .Bands(0).Cells(Vlist.Key) Then
            'Next
            If "" <> "" Then
                e.Cancel = True
            Else
                'If Val(.Cells("ROW_NO").Value & "") = 0 Then
                '    Dim ROW_NO As Integer = Val(grdRecap.DataSource.Compute("MAX(ROW_NO)", "") & "")
                '    .Cells("ROW_NO").Value = ROW_NO + 1
                'End If
            End If
        End With

    End Sub

    Sub Setup_grdASTRECAP()

        tblASTRECAP = New DataTable("ASTRECAP")
        Dim dcx As DataColumn

        Recap_Report = False

        Dim sql As String
        sql = "Select * from ASTDSQLV where FORM_NAME = '" & FORM_NAME & "'"
        For Each rowASTDSQLV As DataRow In _
        ASCDATA1.GetDataTable(sql, "ASTDSQLV").Select("", "VALUE_LIST_SEQ")

            If Not Recap_Report Then
                dcx = New DataColumn
                dcx.ColumnName = "ASTSRPT1_RECAP_ROW_NO"
                dcx.DataType = GetType(System.Int32)
                dcx.Caption = "No"
                dcx.ReadOnly = False
                dcx.AutoIncrement = True
                dcx.AutoIncrementSeed = 1
                tblASTRECAP.Columns.Add(dcx)

                dcx = New DataColumn
                dcx.ColumnName = "ASTSRPT1_RECAP_ROW_CAPTION"
                dcx.DataType = GetType(System.String)
                dcx.Caption = "Caption"
                dcx.ReadOnly = False
                tblASTRECAP.Columns.Add(dcx)
                Recap_Report = True
            End If

            Dim VALUE_LIST_NAME As String = rowASTDSQLV.Item("VALUE_LIST_NAME") & ""
            Dim VALUE_LIST_CAPTION As String = rowASTDSQLV.Item("VALUE_LIST_CAPTION") & ""
            Dim dc As New DataColumn
            dc.ColumnName = VALUE_LIST_NAME
            dc.DataType = GetType(System.String)
            dc.Caption = VALUE_LIST_CAPTION
            dc.ReadOnly = False
            dc.AllowDBNull = False
            tblASTRECAP.Columns.Add(dc)

            Dim vlist As New ValueList
            vlist.Key = VALUE_LIST_NAME

            sql = "Select * from ASTDSQLW where FORM_NAME = '" & FORM_NAME & "'" _
                & " and VALUE_LIST_NAME = '" & VALUE_LIST_NAME & "'"
            For Each rowASTDSQLW As DataRow In _
            ASCDATA1.GetDataTable(sql, "ASTDSQLW").Select("", "VALUE_LIST_CODE_SEQ")
                Dim VALUE_LIST_CODE As String = rowASTDSQLW.Item("VALUE_LIST_CODE") & ""
                Dim VALUE_LIST_DESC As String = rowASTDSQLW.Item("VALUE_LIST_DESC") & ""
                vlist.ValueListItems.Add(VALUE_LIST_CODE, VALUE_LIST_DESC)
            Next
            grdASTRECAP.DisplayLayout.ValueLists.Add(vlist)
        Next


        dcx = New DataColumn
        dcx.ColumnName = "ASTSRPT1_RECAP_ROW_CALC"
        dcx.DataType = GetType(System.String)
        dcx.Caption = "Calculation"
        dcx.ReadOnly = False
        tblASTRECAP.Columns.Add(dcx)


        Absx1.chkFor("RECAP_LAST_LEVEL").Checked = False
        Absx1.chkFor("RECAP_LAST_LEVEL").Visible = False

        If Recap_Report Then
            grdASTRECAP.Visible = True
            grdASTRECAP.DataSource = tblASTRECAP
            Sort_grdColumns(grdASTRECAP, "ASTSRPT1_RECAP_ROW_NO")
            For Each vlist As ValueList In grdASTRECAP.DisplayLayout.ValueLists
                grdASTRECAP.DisplayLayout.Bands(0).Columns(vlist.Key).ValueList = vlist
                grdASTRECAP.DisplayLayout.Bands(0).Columns(vlist.Key).Style = UltraWinGrid.ColumnStyle.DropDownList
            Next
            tblASTRECAP.PrimaryKey = New DataColumn() {tblASTRECAP.Columns("ASTSRPT1_RECAP_ROW_NO")}
            'grdASTRECAP.DisplayLayout.Bands(0).Columns("ASTSRPT1_RECAP_ROW_NO").CellActivation = UltraWinGrid.Activation.NoEdit
            grdASTRECAP.DisplayLayout.Bands(0).Columns("ASTSRPT1_RECAP_ROW_NO").Width = 50
            grdASTRECAP.DisplayLayout.Bands(0).Columns("ASTSRPT1_RECAP_ROW_CAPTION").Width = 150
            grdASTRECAP.DisplayLayout.Bands(0).Columns("ASTSRPT1_RECAP_ROW_CALC").Width = 150

            For Each gcol As UltraWinGrid.UltraGridColumn In grdASTRECAP.DisplayLayout.Bands(0).Columns
                If New String() {"ASTSRPT1_RECAP_ROW_NO", "ASTSRPT1_RECAP_ROW_CAPTION", "ASTSRPT1_RECAP_ROW_CALC"}.Contains(gcol.Key) Then
                Else
                    gcol.Width = 120
                End If
            Next
           
        Else
            If PB_Report Then
                Absx1.chkFor("RECAP_LAST_LEVEL").Visible = True
            End If

            SplitContainer5.Panel2Collapsed = True
        End If

        chkDQ_Only.Visible = Recap_Report
    End Sub
#End Region

#Region "grdSetup"
    Shadows Sub Clear_grdSetup(ByVal Clear_All As Boolean)
        grdSetup.UpdateData()
        grdSetup.ActiveRow = Nothing
        For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Rows
            dr.Item("SEQUENCE") = DBNull.Value
            dr.Item("PAGE_BREAK") = "0"
            If Clear_All Then
                dr.Item("EXCLUDE") = "0"
                dr.Item("GROUP_ALL_OTHERS") = "0"
                dr.Item("CODE_VALUES") = ""
            End If
        Next
        SEQs = 0
        Call Re_SEQ()

    End Sub

    Shadows Sub Setup_grdSetup()
        tblASTDSQLA = Create_tblASTDSQLA()

        Call Get_PARM("GLTPARM1")

        Dim COLUMN_CAPTION As String = ""
        For Each dr As DataRow In ASCDATA1.GetDataTable("Select ASTDSQLA.COLUMN_NAME, NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION) COLUMN_CAPTION, ASTDSQLA.SORTABLE, ASTDSQLA.COLUMN_LAST from ASTDSQLA,ASTDSQLK WHERE ASTDSQLK.COLUMN_NAME (+) = ASTDSQLA.COLUMN_NAME and ASTDSQLA.FORM_NAME = '" & FORM_NAME & "' ORDER BY NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION)").Rows
            If dr.Item("COLUMN_NAME") = "SEG2_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG3_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG4_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
            Then
                ' SKIP IT
            Else
                COLUMN_CAPTION = dr.Item("COLUMN_CAPTION") & ""
                If dr.Item("COLUMN_NAME") = "SEG2_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG3_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG4_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")
                End If
                If dr.Item("SORTABLE") & "" = "1" Or dr.Item("COLUMN_LAST") & "" = "1" Then
                    PB_Report = True
                End If
                If dr.Item("COLUMN_LAST") & "" = "1" Then
                    COLUMN_NAME_last = dr.Item("COLUMN_NAME")
                    dr.Item("SORTABLE") = "0"
                    'PB_Report = True ?
                End If
                'Call Add_Row(tblASTDSQLA, dr.Item("COLUMN_CAPTION") & "", dr.Item("COLUMN_NAME") & "", dr.Item("SORTABLE") & "")
                Add_Row(tblASTDSQLA, COLUMN_CAPTION, dr.Item("COLUMN_NAME") & "", dr.Item("SORTABLE") & "")
            End If

        Next dr

        If PB_Report Then
            UltraTabControl1.Tabs(0).Text = "Sort && Filter"
        Else
            UltraTabControl1.Tabs(0).Text = "Filter"
        End If

        If PB_Report Then
            tblASTDSQLH = ASCDATA1.GetDataTable("Select * from ASTDSQLH WHERE FORM_NAME = '" & FORM_NAME & "'")
            tblASTDSQLS = ASCDATA1.GetDataTable("Select * from ASTDSQLS WHERE FORM_NAME = '" & FORM_NAME & "'")
            'For Each row As DataRow In ASCDATA1.GetDataTable("Select * from ASTDSQLS WHERE FORM_NAME = '" & FORM_NAME & "' ORDER BY COLUMN_SEQ").Rows
            For Each row As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
                Try
                    COLUMN_NAME_sum.Add(row.Item("COLUMN_NAME"), row.Item("COLUMN_TYPE"))
                    If Val(row.Item("COLUMN_WIDTH") & "") = 0 Then
                        row.Item("COLUMN_WIDTH") = 80
                    End If
                    If row.Item("COLUMN_CAPTION") & "" = "" Then
                        row.Item("COLUMN_CAPTION") = row.Item("COLUMN_NAME")
                    End If

                Catch ex As Exception
                    MsgBox("Cannot Add Report Summary Column for " & row.Item("COLUMN_NAME"))
                End Try
            Next
        End If

        grdSetup.DataSource = tblASTDSQLA
        grdSetup.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus
        If grdSetup.Rows.Count <> 0 Then
            'grdSetup.Rows(0).Height = 25
        End If

        SEQs = 0
        Call Re_SEQ()
        'grdSetup.DisplayLayout.Bands(0).SortedColumns.Add(grdSetup.DisplayLayout.Bands(0).Columns("COLUMN_CAPTION"), False)

        If grdSetup.Rows.Count = 0 Then
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(1)
            UltraTabControl1.TabIndex = 1
            UltraTabControl1.Tabs(0).Visible = False
            UltraTabControl1.Tabs(1).Text = "Run-Time Options"
        Else
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(0)
        End If


        ' GET TO THE TOP
        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

        grdSetup.UpdateData()

        grdSetup.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSetup.DisplayLayout.Bands(0).SortedColumns.Add("COLUMN_CAPTION", False)
    End Sub

    'Sub Create_tblASTDSQLA()
    '    tblASTDSQLA = New DataTable
    '    tblASTDSQLA.Columns.Add("COLUMN_NAME")
    '    tblASTDSQLA.Columns.Add("COLUMN_CAPTION")
    '    tblASTDSQLA.Columns.Add("CODE_VALUES")
    '    tblASTDSQLA.Columns.Add("EXCLUDE")
    '    tblASTDSQLA.Columns.Add("SEQUENCE", GetType(System.Int16))
    '    tblASTDSQLA.Columns.Add("PAGE_BREAK")
    '    tblASTDSQLA.Columns.Add("SORTABLE")
    '    tblASTDSQLA.Columns.Add("GROUP_ALL_OTHERS")
    '    tblASTDSQLA.Columns.Add("COLUMN_LAST")
    '    tblASTDSQLA.PrimaryKey = New DataColumn() {tblASTDSQLA.Columns("COLUMN_NAME")}
    'End Sub

    'Sub Add_Row( _
    'ByVal COLUMN_CAPTION As String, _
    'ByVal COLUMN_NAME As String, _
    'ByVal SORTABLE As String)
    '    Dim dr As DataRow
    '    dr = tblASTDSQLA.NewRow
    '    dr.Item("COLUMN_NAME") = COLUMN_NAME
    '    dr.Item("COLUMN_CAPTION") = COLUMN_CAPTION
    '    dr.Item("EXCLUDE") = "0"
    '    dr.Item("PAGE_BREAK") = "0"
    '    dr.Item("SORTABLE") = SORTABLE
    '    dr.Item("GROUP_ALL_OTHERS") = "0"
    '    tblASTDSQLA.Rows.Add(dr)
    'End Sub

    Shadows Sub Re_SEQ( _
    Optional ByVal COLUMN_NAME As String = "", _
    Optional ByVal add_to_sort As Boolean = False)

        'grdSetup.Update 
        grdSetup.UpdateData()

        Dim tbl As DataTable = DirectCast(grdSetup.DataSource, DataTable)
        Dim row As DataRow

        If COLUMN_NAME <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME)
            If add_to_sort Then
                row.Item("SEQUENCE") = 9
            Else
                row.Item("SEQUENCE") = Null
                row.Item("PAGE_BREAK") = "0"
            End If
        End If

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            row.Item("SEQUENCE") = Null
            row.Item("PAGE_BREAK") = "0"
        End If

        SEQs = 0
        For Each dr As DataRow In tbl.Select _
            ("SEQUENCE IS NOT NULL OR SEQUENCE <> ''", "SEQUENCE")
            SEQs = SEQs + 1
            dr.Item("SEQUENCE") = SEQs
        Next

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            SEQs = SEQs + 1
            row.Item("SEQUENCE") = SEQs
        End If

    End Sub

    Sub Rebuild_Values()
        Dim CODE_VALUES As String = ""
        For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
            CODE_VALUES = CODE_VALUES & "," & gr.Cells(0).Text
        Next
        CODE_VALUES = Mid$(CODE_VALUES, 2)
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = CODE_VALUES
        If CODE_VALUES = "" Then
            grdSetup.ActiveRow.Cells("EXCLUDE").Value = "0"
        End If
        If CODE_VALUES = "" Or grdSetup.ActiveRow.Cells("SEQUENCE").Value & "" = "" Then
            grdSetup.ActiveRow.Cells("GROUP_ALL_OTHERS").Value = "0"
        End If
        grdSetup.UpdateData()

        Dim z As String = ASCMAIN1.CodeSelector.VIEW_DESC
        If grd.Rows.Count <> 0 Then
            z = z & " (" & CStr(grd.Rows.Count) & ")"
        End If
        grd.Text = z

        cmdAll.Visible = (CODE_VALUES <> "")
    End Sub

    Private Sub grdSetup_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.AfterCellUpdate
        'Try
        '    grdSetup.UpdateData()
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub grdSetup_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSetup.AfterRowActivate
        Call Show_grd()
    End Sub

    Private Sub grdSetup_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSetup.AfterRowUpdate
        Call Show_grd()
    End Sub

    Private Sub grdSetup_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSetup.BeforeRowUpdate

        'If Val(grdSetup.ActiveRow.Cells("SEQUENCE").Value & "") = 0 Then
        '    grdSetup.ActiveRow.Cells("PAGE_BREAK").Value = "0"
        'End If
        If Val(e.Row.Cells("SEQUENCE").Value & "") = 0 Then
            grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("PAGE_BREAK").Value = "0"
        End If

        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text ' grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
        If sql <> "" Then
            Dim CODE_VALUES_new As String = ""
            Dim CODE_VALUES As String = e.Row.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("COLUMN_NAME") & ""
            If KEY_EXPRESSION = "" Then
                KEY_EXPRESSION = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")
            End If
            If KEY_EXPRESSION = "T_CODE" And ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_ALIAS") & "" <> "T_CODE" Then
                KEY_EXPRESSION = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_ALIAS") & ""
            End If

            If CODE_VALUES <> "" Then
                Dim CODE_VALUES_old As String = ""
                For Each txt As String In Split(Replace(CODE_VALUES, "'", ""), ",")
                    txt = txt.Trim
                    If ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Count > 0 Then
                        Dim COLUMN_VALUEs() As String = Split(txt, "-")
                        If COLUMN_VALUEs.Length = 2 Then
                            CODE_VALUES_old = CODE_VALUES_old & ",'" & COLUMN_VALUEs(0) & "-" & ASCMAIN1.Format_Field(COLUMN_VALUEs(1), COLUMN_NAME, , True) & "'"
                        End If
                    Else
                        CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                    End If
                    'CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                Next
                CODE_VALUES_old = Mid$(CODE_VALUES_old, 2)
                Dim where_or_and As String = " where "
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    where_or_and = " and "
                End If

                'For Each dr As DataRow In ASCDATA1.GetDataTable _
                '(sql & where_or_and & KEY_EXPRESSION & " IN (" & CODE_VALUES_old & ")" & Get_PreKey_filter(True)).Rows

                Dim COLUMN_NAMEs As String = KEY_EXPRESSION
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("CODE_TABLE") & "" <> "" Then
                    COLUMN_NAMEs = "T_CODE"
                End If
                If ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Count > 0 Then
                    COLUMN_NAMEs = ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Keys(0) & " || '-' || " & KEY_EXPRESSION
                End If
                If CODE_VALUES_old <> "" Then
                    For Each dr As DataRow In ASCDATA1.GetDataTable _
                        (sql & where_or_and & COLUMN_NAMEs & " IN (" & CODE_VALUES_old & ")").Rows
                        If ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Count > 0 Then
                            CODE_VALUES_new &= "," & dr.Item(ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Keys(0)) & "-" & dr.Item(KEY_EXPRESSION)
                        Else
                            If Not dr.Table.Columns.Contains(KEY_EXPRESSION) Then
                                CODE_VALUES_new &= "," & dr.Item(0)
                            Else
                                CODE_VALUES_new &= "," & dr.Item(KEY_EXPRESSION)
                            End If
                        End If
                    Next
                End If
            End If

            CODE_VALUES_new = Mid(CODE_VALUES_new, 2)
            If CODE_VALUES_new <> CODE_VALUES Then
                cmdAll.Visible = (CODE_VALUES <> "")

                grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("CODE_VALUES").Value = CODE_VALUES_new '  .ActiveRow.Cells("CODE_VALUES").Value = CODE_VALUES_new
                Call Show_grd()
            End If
        End If

    End Sub

    Private Sub grdSetup_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.ClickCellButton
        If e.Cell.Column.Key = "COLUMN_CAPTION" Then
            If e.Cell.Row.Cells("SORTABLE").Text = "1" Then
                If e.Cell.Row.Cells("SEQUENCE").Text <> "" Then
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, False)
                Else
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, True)
                End If
            End If
        ElseIf e.Cell.Column.Key = "CODE_VALUES" Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdSetup.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))

                Dim pf As String = Get_PreKey_filter(False)
                pf &= Get_Custom_Filter_for_Codes_Selection(e.Cell.Row.Cells("COLUMN_NAME").Text)
                If pf <> "" Then
                    ASCMAIN1.CodeSelector.SQL &= IIf(InStr(ASCMAIN1.CodeSelector.SQL.ToLower, " where ") = 0, " where ", " and ") & pf
                End If

                'For Each row As DataRow In ASCMAIN1.CodeSelector.grdColumns
                '    If row.Item("COLUMN_PREKEY") & "" = "1" Then
                '        Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
                '        Dim CODE_VALUES As String = SQLA(COLUMN_NAME, , True)
                '        ASCMAIN1.CodeSelector.SQL &= IIf(InStr(" where ", ASCMAIN1.CodeSelector.SQL.ToLower) = 0, " where ", " and ") & COLUMN_NAME & " in (" & CODE_VALUES & ")"
                '    End If
                'Next

                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    If ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Count <> 0 Then
                        Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("COLUMN_NAME")
                        If KEY_EXPRESSION = "" Then
                            KEY_EXPRESSION = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")
                        End If

                        Dim CODES As String = ""
                        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                            CODES &= "," & row.Item(ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Keys(0)) & "-" & row.Item(KEY_EXPRESSION)
                        Next

                        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid(CODES, 2)
                    Else
                        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                    End If
                    grdSetup.UpdateData()
                    Call Show_grd()
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_DoubleClickHeader(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs) Handles grdSetup.DoubleClickHeader
        Call Clear_grdSetup(False)
    End Sub

    Private Sub grdSetup_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSetup.InitializeRow
        If e.Row.Cells("SORTABLE").Text <> "1" Then
            e.Row.Cells("COLUMN_CAPTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit
        End If
    End Sub

    Private Sub grdSetup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdSetup.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Delete Then
            If grdSetup.ActiveCell IsNot Nothing Then
                If grdSetup.ActiveCell.Column.Key = "SEQUENCE" Then
                    If grdSetup.ActiveCell.Text <> "" Then
                        'grdSetup.ActiveCell.Value = DBNull.Value
                        'grdSetup.UpdateData()
                        Call Re_SEQ(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text, False)
                    End If
                End If
            End If
        End If
        If e.Control And e.KeyValue = 86 Then
            If grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
                Dim ED As EmbeddableEditorBase = grdSetup.ActiveCell.EditorResolved

                Dim c As String = Replace(My.Computer.Clipboard.GetText(), vbCrLf, ",")
                Dim CODE_VALUES As String = grdSetup.ActiveCell.Text & ""
                If CODE_VALUES <> "" AndAlso Not c.StartsWith(",") Then c = "," & c
                'c = "," & c
                CODE_VALUES &= c
                'If ED.IsInEditMode Then
                '    c = "," & c
                '    CODE_VALUES &= c
                '    ED.Value = CODE_VALUES
                'Else
                grdSetup.ActiveCell.Value = CODE_VALUES
                'End If
                'grdSetup.ActiveCell.Value = CODE_VALUES
                e.SuppressKeyPress = True
                e.Handled = True
                grdSetup.Update()
            End If
        End If


        If e.KeyValue = Windows.Forms.Keys.Enter Then
            If grdSetup.ActiveCell IsNot Nothing Then
                If grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
                    grdSetup.Update()
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSetup.KeyPress
        If grdSetup.ActiveCell IsNot Nothing Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" And grdSetup.ActiveRow.Cells("SORTABLE").Text = "1" Then
                Dim COLUMN_NAME As String = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
                Dim SEQcur As Integer = Val(grdSetup.ActiveCell.Text)
                Dim SEQnew As Integer = Val(e.KeyChar)
                If SEQnew < 1 Or SEQnew = SEQcur Or (SEQcur = 0 And SEQnew > SEQs + 1) Or (SEQcur <> 0 And SEQnew > SEQs) Then
                    Exit Sub
                End If

                grdSetup.ActiveCell.Value = SEQnew
                grdSetup.UpdateData()

                Dim i As Integer
                Dim z As String
                If SEQnew < SEQcur Or SEQcur = 0 Then
                    z = ">"
                    i = SEQnew
                Else
                    z = "<"
                    i = 0
                End If
                For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Select("SEQUENCE " & z & "= " & CStr(SEQnew), "SEQUENCE")
                    If dr.Item("COLUMN_NAME") <> COLUMN_NAME Then
                        i = i + 1
                        dr.Item("SEQUENCE") = i
                    End If
                Next

                If SEQcur = 0 Then
                    SEQs = SEQs + 1
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSetup.Leave
        grdSetup.UpdateData()
    End Sub
#End Region

#Region "grd"
    Sub Show_grd()
        LIST_CODE = ""
        LIST_DESC = ""

        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)
        If sql = "" Then
            grd.Visible = False
            cmdAll.Visible = False
            SplitContainer4.Panel1.Hide()
            grpCodeLists.Visible = False
        Else
            Dim CODE_VALUES As String = grdSetup.ActiveRow.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")
            Dim sqlx As String = " where "
            If InStr(sql, " where ") <> 0 Then
                sqlx = " and "
            End If
            If CODE_VALUES <> "" Then
                sql = sql & sqlx & KEY_EXPRESSION & " IN ('" & Replace(Replace(CODE_VALUES, "'", ""), ",", "','") & "')"
            Else
                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    sql = sql & sqlx & "1 <> 1"
                Else
                    sql = sql & sqlx & "ROWNUM < 1"
                End If
            End If
            grd.DataSource = Nothing
            grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grd.DataSource = ASCDATA1.GetDataTable(sql)
            Dim z As String = ASCMAIN1.CodeSelector.VIEW_DESC
            If grd.Rows.Count <> 0 Then
                z = z & " (" & CStr(grd.Rows.Count) & ")"
            End If
            grd.Text = z

            For i As Integer = 0 To ASCMAIN1.CodeSelector.grdColumns.Count - 1 ' grd.DisplayLayout.Bands(0).Columns.Count - 1
                grd.DisplayLayout.Bands(0).Columns(i).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_CAPTION")
                If Val(ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_WIDTH") & "") <> 0 Then
                    grd.DisplayLayout.Bands(0).Columns(i).Width = ASCMAIN1.CodeSelector.grdColumns(i).Item("COLUMN_WIDTH")
                End If
            Next i
            grd.Visible = True
            cmdAll.Visible = (grd.Rows.Count <> 0)

            SplitContainer4.Panel1.Show()
            grpCodeLists.Visible = True

            txtList.Text = ""
            chkListShareable.Checked = False
            chkListModifiable.Checked = False
        End If
    End Sub

    Private Sub grd_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.AfterRowsDeleted
        Call Rebuild_Values()
    End Sub
#End Region

#Region "grdASTROPT1"

    Private Sub grdASTROPT1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.AfterRowActivate
        With grdASTROPT1.DisplayLayout.Bands(0)
            If grdASTROPT1.ActiveRow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID Then
                grdASTROPT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdASTROPT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                '.Columns("SET_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("SET_YP_REL").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("SET_ALLOW_OTHERS").CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                grdASTROPT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdASTROPT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                '.Columns("SET_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                '.Columns("SET_YP_REL").CellActivation = UltraWinGrid.Activation.AllowEdit
                '.Columns("SET_ALLOW_OTHERS").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With
    End Sub

    Private Sub grdASTROPT1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.AfterRowsDeleted
        For J As Integer = 1 To ASCMAIN1.grdRows.Count
            Call Delete_Saved_Setting(ASCMAIN1.grdRows(J))
        Next
        tblASTROPT1.AcceptChanges()
    End Sub

    Sub Delete_Saved_Setting(ByVal SET_ID As String)
        Dim Sql As String = ""
        Sql = "Delete from ASTROPT2 where FORM_NAME = '" & FORM_NAME & "' and SET_ID = '" & SET_ID & "'"
        ASCDATA1.ExecuteSQL(Sql)
        Sql = "Delete from ASTROPT1 where FORM_NAME = '" & FORM_NAME & "' and SET_ID = '" & SET_ID & "'"
        ASCDATA1.ExecuteSQL(Sql)
    End Sub

    Private Sub grdASTROPT1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTROPT1.AfterRowUpdate
        Dim TBL As New DataTable
        With ASCDATA1.GetDataAdapter(TBL, "ASTROPT1", "*", True, 2, False, 0)
            For Each rowASTROPT1 As DataRow In DirectCast(grdASTROPT1.DataSource, DataTable) _
                .Select("SET_ID = '" & e.Row.Cells("SET_ID").Text & "'")
                Dim row As DataRow = TBL.NewRow
                row.ItemArray = rowASTROPT1.ItemArray
                TBL.Rows.Add(row)
                row.AcceptChanges()
                row.SetModified()
            Next
            .Update(TBL)
            tblASTROPT1.AcceptChanges()
            .Dispose()
        End With
        If grdASTROPT1.ActiveRow.Cells("SET_ID").Text = SET_ID Then
            txtDescription.Text = grdASTROPT1.ActiveRow.Cells("SET_DESC").Text
        End If
    End Sub

    Private Sub grdASTROPT1_BeforeEnterEditMode(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles grdASTROPT1.BeforeEnterEditMode
        'If chkExecutionHistory.Checked Or (grdASTROPT1.ActiveRow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID And Not grdASTROPT1.ActiveRow.IsAddRow) Then
        '    e.Cancel = True
        'End If
    End Sub

    Private Sub grdASTROPT1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTROPT1.BeforeRowsDeleted
        'If chkExecutionHistory.Checked Then
        '    e.Cancel = True
        'End If

        ASCMAIN1.grdRows.Clear()
        For Each DR As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            ASCMAIN1.grdRows.Add(DR.Cells("SET_ID").Text)
            'If DR.Cells("INIT_OPER").Value <> ASCMAIN1.USER_ID Then
            '    MsgBox("You cannot delete or modify records that you did not create" & vbCr & vbCr & "(" & DR.Cells("SET_DESC").Value & " was created by " & DR.Cells("INIT_OPER").Value & ")", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
            '    e.Cancel = True
            '    ASCMAIN1.grdRows.Clear()
            '    Exit For
            'End If
        Next
    End Sub

    Private Sub grdASTROPT1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT1.ClickCellButton
        SET_ID = grdASTROPT1.ActiveRow.Cells("SET_ID").Text
        Call Retrieve_Settings()
    End Sub

    Private Sub grdASTROPT1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdASTROPT1.KeyDown
        If e.KeyValue = Windows.Forms.Keys.Enter Then
            grdASTROPT1.UpdateData()
        End If
    End Sub

    Private Sub grdASTROPT1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT1.Leave
        grdASTROPT1.UpdateData()
    End Sub
#End Region

#Region "Main - Preparation"

    Protected Friend Sub Main_Process()


        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        START_TIME = DATETIME_STAMP ' Now + ASCMAIN1.NowTSD
        Call ASCMAIN1.Get_Current_YP()
        XNO = ASCMAIN1.Next_Control_No(Me.Name)
        Call Save_Settings(SET_ID, XNO)
        ASCMAIN1.SET_ID = SET_ID

        ' Record that Proceed has been clicked
        rowASTOPST1.Item("PROCEED_BEGIN") = START_TIME
        rowASTOPST1.Item("SET_ID") = SET_ID
        rowASTOPST1.Item("SET_DESC") = txtDescription.Text
        rowASTOPST1.Item("XNO") = XNO
        tdaASTOPST1.Update(tblASTOPST1)

        RPT = FORM_NAME
        RPT_TITLE = Me.Text
        SUBT = txtDescription.Text

        Call ASCMAIN1.TACMAIN1.Site_Specific_Settings()
        Call Build_WorkFile_DB_Init()
        Call ASCMAIN1.Progress("Run-Time Options")
        Call Load_ASTDSQLA() ' Traverse tblASTDSQLA and setup some variables, Page0, etc

        Call Build_Workfile()
        If PB_Report Then
            Build_Report_File()
        End If
        Call Wrap_Up()
        'Application.DoEvents()

        Call Write_DataSet()
        Call Write_DataSet(True)


        If xErrMsg = "" And RWU <> "U" Then
            Call ASCMAIN1.Progress("Now Printing Reports")
            tblASTSPRF1_clone.Clear()
            CR_params.Clear()
            Call Print_Report_Main()
        End If

        ' Record that the Process has Ended
        rowASTOPST1.Item("PROCEED_END") = Now + ASCMAIN1.NowTSD
        tdaASTOPST1.Update(tblASTOPST1)

        If PB_Report Then
            grdASTSRPT1.DataSource = dst.Tables("ASTSRPT1")

            If Not UltraTabControl1.Tabs("Data Query").Visible Then
                UltraTabControl1.Tabs("Data Query").Visible = True
                UltraTabControl1.Tabs("Data Grid").Visible = True
                For i As Integer = 1 To 9
                    With grdASTSRPT1.DisplayLayout.Bands(0).Columns("G" & CStr(i))
                        .CellAppearance.BackColor = Color.LightSteelBlue
                    End With
                Next

                grdASTSRPT1.DisplayLayout.Override.FilterUIType = UltraWinGrid.FilterUIType.FilterRow
                grdASTSRPT1.DisplayLayout.Override.FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
                grdASTSRPT1.DisplayLayout.Override.FilterRowAppearance.BackColor = System.Drawing.Color.AliceBlue
                grdASTSRPT1.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True
                grdASTSRPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

                For Each SCN As String In COLUMN_NAME_sum.Keys
                    With grdASTSRPT1.DisplayLayout.Bands(0).Columns(SCN)
                        Select Case COLUMN_NAME_sum(SCN)
                            Case "QTY"
                                .Format = "#,##0"
                            Case "AMT"
                                .Format = "#,##0.00"
                            Case "DEC"
                                .Format = "#,##0.00"
                        End Select
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                    End With
                    Call Create_Summary(grdASTSRPT1, SCN)
                Next
            End If

            If COLUMN_NAMEs.Count <> 0 Then
                Initialize_Data_Query()
            End If

            Post_Process_Special()

            'If MENU_ITEM_OBJECT = "RSRCOMP1" Then
            '    Try
            '        Prepare_XLS()
            '    Catch ex As Exception
            '        If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
            '    End Try
            'End If

            UltraTabControl1.Tabs("Data Query").Enabled = True
            UltraTabControl1.Tabs("Data Grid").Enabled = True

            For i As Integer = 1 To 9
                With grdASTSRPT1.DisplayLayout.Bands(0).Columns("G" & CStr(i))
                    .Hidden = (i > COLUMN_NAMEs.Count)
                    If i <= COLUMN_NAMEs.Count Then
                        .Header.Caption = COLUMN_CAPTIONs(i - 1)
                    End If
                End With
            Next

        End If

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Overridable Sub Post_Process_Special()

    End Sub

    Public Sub Build_WorkFile_DB_Init()
        Dim sql As String

        Call ASCMAIN1.Track("Initialize Environment", "")
        Page0.Clear()

        Clear_dst()

        ASTSRPT1 = ""

        ' SQL Optimization (Oracle Only)

        sql = "Select * from ASTDSQLF where FORM_NAME = '" & FORM_NAME & "'"
        tblASTDSQLF = ASCDATA1.GetDataTable(sql, "ASTDSQLF")
        If tblASTDSQLF.Rows.Count = 1 Then
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            Else
                Dim OPTIMIZER_MODE As String
                OPTIMIZER_MODE = tblASTDSQLF.Rows(0).Item("OPTIMIZER_MODE") & ""
                If OPTIMIZER_MODE = "C" Then
                    OPTIMIZER_MODE = "CHOOSE"
                ElseIf OPTIMIZER_MODE = "R" Then
                    OPTIMIZER_MODE = "RULE"
                Else
                    OPTIMIZER_MODE = ""
                End If
                If OPTIMIZER_MODE <> "" Then
                    ASCDATA1.ExecuteSQL("ALTER SESSION SET OPTIMIZER_MODE = " & OPTIMIZER_MODE)
                End If
            End If

        End If

        ' Prepare for Dynamic SQL Generation

        tblASTDSQLB = ASCDATA1.GetDataTable("Select * from ASTDSQLB where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLB")

        ASCMAIN1.sql = "SELECT X.FORM_NAME, X.DATA_SOURCE, X.COLUMN_NAME, " _
            & "ASTDSQLC.TABLE_NAME, ASTDSQLC.COLUMN_EXPRESSION, ASTDSQLC.JOIN_SPECIAL, ASTDSQLC.NO_FILTER " _
            & "FROM ASTDSQLC, ( " _
            & "SELECT ASTDSQLF.FORM_NAME, ASTDSQLB.DATA_SOURCE, ASTDSQLA.COLUMN_NAME " _
            & "FROM ASTDSQLF, ASTDSQLA, ASTDSQLB " _
            & "WHERE ASTDSQLF.FORM_NAME = '" & FORM_NAME & "' " _
            & "AND ASTDSQLF.FORM_NAME = ASTDSQLA.FORM_NAME " _
            & "AND ASTDSQLF.FORM_NAME = ASTDSQLB.FORM_NAME " _
            & ") X " _
            & "WHERE ASTDSQLC.FORM_NAME (+) = X.FORM_NAME " _
            & "AND ASTDSQLC.DATA_SOURCE (+) = X.DATA_SOURCE " _
            & "AND ASTDSQLC.COLUMN_NAME (+) = X.COLUMN_NAME "
        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "SELECT     X.FORM_NAME, X.DATA_SOURCE, X.COLUMN_NAME, ASTDSQLC.TABLE_NAME, ASTDSQLC.COLUMN_EXPRESSION, ASTDSQLC.JOIN_SPECIAL, ASTDSQLC.NO_FILTER" _
            & " FROM         ASTDSQLC RIGHT OUTER JOIN" _
            & "                          (SELECT     ASTDSQLF.FORM_NAME, ASTDSQLB.DATA_SOURCE, ASTDSQLA.COLUMN_NAME" _
            & "                            FROM          ASTDSQLF INNER JOIN" _
            & "                                                   ASTDSQLA ON ASTDSQLF.FORM_NAME = ASTDSQLA.FORM_NAME INNER JOIN" _
            & "                                                   ASTDSQLB ON ASTDSQLF.FORM_NAME = ASTDSQLB.FORM_NAME" _
            & "                            WHERE      (ASTDSQLF.FORM_NAME = '" & FORM_NAME & "')) AS X ON ASTDSQLC.FORM_NAME = X.FORM_NAME AND ASTDSQLC.DATA_SOURCE = X.DATA_SOURCE AND " _
            & "                      ASTDSQLC.COLUMN_NAME = X.COLUMN_NAME"
        End If
        tblASTDSQLC = ASCDATA1.GetDataTable("", "ASTDSQLC", 3)
        tblASTDSQLD = ASCDATA1.GetDataTable("Select * from ASTDSQLD where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLD")
        tblASTDSQLJ = ASCDATA1.GetDataTable("Select * from ASTDSQLJ where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLJ")

        ' Get Commonly used RT Options into Variables

        Try
            RYPLEGEND = Absx1.cmbFor("RYP", True).Value
            RYP = Mid(RYPLEGEND, 1, 4) & Mid(RYPLEGEND, 6, 2)
        Catch ex As Exception
        End Try
        Try
            RYPLEGEND0 = Absx1.cmbFor("RYP0", True).Value
            RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        Catch ex As Exception
        End Try
        Try
            RYPLEGEND1 = Absx1.cmbFor("RYP1", True).Value
            RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)
        Catch ex As Exception
        End Try

        Try
            RYWLEGEND = Absx1.cmbFor("RYW", True).Value
            RYW = Mid(RYWLEGEND, 1, 4) & Mid(RYWLEGEND, 6, 2)
        Catch ex As Exception
        End Try
        Try
            RYWLEGEND0 = Absx1.cmbFor("RYW0", True).Value
            If RYWLEGEND0 <> "" AndAlso RYWLEGEND0.Length = 6 Then
                RYWLEGEND0 = Absx1.cmbFor("RYW0", True).Text
            End If
            RYW0 = Mid(RYWLEGEND0, 1, 4) & Mid(RYWLEGEND0, 6, 2)
        Catch ex As Exception
        End Try
        Try
            RYWLEGEND1 = Absx1.cmbFor("RYW1", True).Value
            If RYWLEGEND1 <> "" AndAlso RYWLEGEND1.Length = 6 Then
                RYWLEGEND1 = Absx1.cmbFor("RYW1", True).Text
            End If
            RYW1 = Mid(RYWLEGEND1, 1, 4) & Mid(RYWLEGEND1, 6, 2)
        Catch ex As Exception
        End Try

        ' Clean-Up

        ASCMAIN1.Track("", "")

    End Sub

    Protected Sub Load_ASTDSQLA()

        Dim COLUMN_NAME As String
        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Rows
            Dim z As String = "All"
            If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                z = rowASTDSQLA.Item("CODE_VALUES") & ""
                If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    z = "All Except " & z
                ElseIf rowASTDSQLA.Item("GROUP_ALL_OTHERS") & "" = "1" Then
                    z = "Explicitly Showing " & z
                End If
            End If

            Page0.Add(rowASTDSQLA.Item("COLUMN_CAPTION") & ":" & z)
        Next

        COLUMN_NAMEs.Clear()
        COLUMN_CAPTIONs.Clear()
        GROUP_ALL_OTHERSs.Clear()
        PAGE_BREAKs = ""
        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE")
            COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")
            COLUMN_NAMEs.Add(COLUMN_NAME)
            COLUMN_CAPTIONs.Add(rowASTDSQLA.Item("COLUMN_CAPTION"))
            GROUP_ALL_OTHERSs.Add(rowASTDSQLA.Item("GROUP_ALL_OTHERS"))
            If rowASTDSQLA.Item("PAGE_BREAK") & "" = "1" Then
                PAGE_BREAKs = PAGE_BREAKs & "Y"
            Else
                PAGE_BREAKs = PAGE_BREAKs & "N"
            End If
        Next

        If COLUMN_NAME_last <> "" Then
            If Not COLUMN_NAMEs.Contains(COLUMN_NAME_last) Then
                MsgBox("Please Contact ABS, Report " & Me.Name & ", Column " & COLUMN_NAME_last & vbCr & "Click OK to Proceed with Report")

                COLUMN_NAMEs.Add(COLUMN_NAME_last)
                COLUMN_CAPTIONs.Add(ASCMAIN1.Make_Caption(COLUMN_NAME_last))
                GROUP_ALL_OTHERSs.Add(False)
            End If
        End If

        Page0.Add("Sort & Sub-Total Sequence: " & Join(COLUMN_CAPTIONs.ToArray, ","))

        tblASTGROUP = ASCDATA1.GetDataTable("*", "ASTGROUP")

        For i As Integer = 1 To COLUMN_NAMEs.Count
            Call Write_Group_Record(COLUMN_CAPTIONs(i - 1), "?", "Not Specified")
        Next i

        'COLUMN_CAPTIONs(0) = "Report"
        'For j As Integer = 0 To COLUMN_NAMEs.Count - 1
        '    grp = "Z"
        '    code = aRC & "Recap " & COLUMN_CAPTIONs(j)
        '    codedesc = "Recap " & COLUMN_CAPTIONs(j)
        '    codedesc = codedesc & ", All " & ASCMAIN1.Make_Plural(COLUMN_CAPTIONs(j + 1))
        '    Call Write_Group_Record(grp, code, codedesc)
        '    code = aCC
        '    codedesc = "Consolidated"
        '    Call Write_Group_Record(grp, code, codedesc)
        'Next j

        ' need to do this call at the end of the report, and get only the codes required

        '' Call Get_Group_Desc_All()
    End Sub

    Protected Friend Overridable Sub Build_Workfile()

    End Sub

    Sub Build_Report_File()

        Call ASCMAIN1.Track("Tiering", "")
        Application.DoEvents()

        ' Get Group Descriptions for all Codes Represented in Gx columns

        Call Get_Group_Desc_All()


        ' Update the Group Field Values to contain the Group Column Captions

        Dim Sql As String = ""
        If COLUMN_NAMEs.Count > 0 Then
            For i As Integer = 1 To COLUMN_NAMEs.Count
                Sql &= ", G" & CStr(i) & " = '" & COLUMN_CAPTIONs(i - 1) & ":" & "' || G" & CStr(i)
            Next i
            Sql = "Update " & ASTSRPT1 & " Set " & Mid(Sql, 2)
            ASCDATA1.ExecuteSQL(Sql)
        End If


        ' Create Summarized Result Set in Temporary Table

        Dim TT As String = ASCMAIN1.Temp_Table(ASTSRPT1_sql_sum)


        ' If Recap, then ensure that a RECAP_ROW_NO exists for every key

        If grdASTRECAP.Rows.Count <> 0 Then

            Sql = "Select W_INT ASTSRPT1_RECAP_ROW_NO from TATWORK1 where ROWNUM < 1"
            Dim TTR As String = ASCMAIN1.Temp_Table(Sql)

            For Each row As DataRow In tblASTRECAP.Rows
                Sql = "Insert into " & TTR & " (ASTSRPT1_RECAP_ROW_NO) VALUES (" & row.Item("ASTSRPT1_RECAP_ROW_NO") & ")"
                ASCDATA1.ExecuteSQL(Sql)
            Next

            Sql = "Insert into " & TT _
                & " (" & G1thru9 & COLUMN_NAMEs_appended & COLUMN_NAME_RECAP_ROW_NO & ") " _
                & " Select " & G1thru9 _
                & COLUMN_NAMEs_appended _
                & ", " & TTR & ".ASTSRPT1_RECAP_ROW_NO " _
                & " from " & TT & ", " & TTR _
                & " minus " _
                & " Select " & G1thru9 _
                & COLUMN_NAMEs_appended _
                & COLUMN_NAME_RECAP_ROW_NO _
                & " from " & TT
            ASCDATA1.ExecuteSQL(Sql)
        End If

        ' If Recap or Recapping Last Level, then Summarize the Recap (or Last Level) Rows for each Summary Key Combination in G1-Gx
        ' good example of recapping last level = ODG.SAR12MO2

        If grdASTRECAP.Rows.Count <> 0 Or Absx1.chkFor("RECAP_LAST_LEVEL").Checked Then
            'For i As Integer = COLUMN_NAMEs.Count + IIf(COLUMN_NAMEs_appended <> "", 1, 0) To 1 Step -1
            Dim imax As Int32 = COLUMN_NAMEs.Count
            If grdASTRECAP.Rows.Count = 0 And Absx1.chkFor("RECAP_LAST_LEVEL").Checked Then
                If Me.Name = "GLRPROF1  - STILL DOESN'T WORK WITH THIS CHANGE" Then ' LINE_TAG
                    imax = imax - 1
                Else
                    'imax = imax - 1
                    imax = imax - 2 ' UNREMMED AND ADJ FROM 1 TO 2 BY WJZ 02/28/16 TO GET RECAP TO WORK FOR NON-12MO STYLE REPORTS
                End If
            End If
            If imax >= 0 Then 'If imax >= 1 Then CHG TO 0 BY WJZ ON 02/28 - SORTING RSRCOMP1 BY CUST/STORE-CLASS, RECAP BY STORE-CLASS
                For i As Integer = imax To 0 Step -1 ' CHANGED TO To 0 by WJZ 07/14 to get the grand total
                    Sql = ""
                    Dim sql_group_by As String = ""
                    If i > 1 Then
                        For j As Integer = 1 To i - 1
                            Sql &= ", G" & CStr(j)
                        Next
                        sql_group_by = Sql
                    End If

                    If grdASTRECAP.Rows.Count = 0 And Absx1.chkFor("RECAP_LAST_LEVEL").Checked And i <> COLUMN_NAMEs.Count Then ' And i <> 0 Then
                        sql_group_by &= ", G" & CStr(COLUMN_NAMEs.Count)
                    End If

                    If i > 0 Then Sql &= ", '" & aRC & "' G" & CStr(i)
                    ' Dim sql_where As String = ""
                    If i < 9 Then
                        For j As Integer = i + 1 To 9
                            If grdASTRECAP.Rows.Count = 0 And Absx1.chkFor("RECAP_LAST_LEVEL").Checked And j = COLUMN_NAMEs.Count Then ' And i <> 0 Then
                                Sql = Sql & ", G" & CStr(j)
                            Else
                                Sql = Sql & ", 'x' G" & CStr(j)
                            End If
                        Next
                    End If

                    If i = 0 Then Sql = Replace(Sql, ", 'x'", ", '" & aRC & "'")

                    'Sql = "Insert into " & TT _
                    '    & " Select " & Mid(Sql, 2) _
                    '    & Replace(COLUMN_NAMEs_appended, ",", ", NULL ") _
                    '    & COLUMN_NAME_RECAP_ROW_NO _
                    '    & ASTSRPT1_sum_columns _
                    '    & " from " & TT
                    ' NEED COLUMN_NAMES_appended for UNITS AND SALES RECAP IN SARCOMP1

                    If Me.Name = "SARCSUM1" Then
                        Sql = "Insert into " & TT _
                            & " Select " & Mid(Sql, 2) _
                            & Replace(COLUMN_NAMEs_appended, ",CUST_CODE", ",'X' CUST_CODE") _
                            & COLUMN_NAME_RECAP_ROW_NO _
                            & ASTSRPT1_sum_columns _
                            & " from " & TT

                    Else
                        Dim ADDED_FIELD As String = COLUMN_NAME_RECAP_ROW_NO
                        If grdASTRECAP.Rows.Count = 0 And Absx1.chkFor("RECAP_LAST_LEVEL").Checked Then
                            ADDED_FIELD = ""
                        End If

                        Sql = "Insert into " & TT _
                            & " Select " & Mid(Sql, 2) _
                            & COLUMN_NAMEs_appended _
                            & ADDED_FIELD _
                            & ASTSRPT1_sum_columns _
                            & " from " & TT

                    End If

                    If i <> imax Then
                        If grdASTRECAP.Rows.Count = 0 And Absx1.chkFor("RECAP_LAST_LEVEL").Checked And i <> 0 Then
                            Sql &= " where G" & CStr(imax) & " <> '" & aRC & "'"
                        Else
                            Sql &= " where G" & CStr(i + 1) & " = '" & aRC & "'"
                        End If
                    End If

                    If sql_group_by & COLUMN_NAME_RECAP_ROW_NO <> "" Or i = 0 Then
                        If Me.Name = "SARCSUM1" Then
                            Sql &= " group by " & Mid(sql_group_by _
                                & Replace(COLUMN_NAMEs_appended, ",CUST_CODE", "") _
                                & COLUMN_NAME_RECAP_ROW_NO, 2)
                        Else
                            Sql &= " group by " & Mid(sql_group_by _
                                & COLUMN_NAMEs_appended _
                                & COLUMN_NAME_RECAP_ROW_NO, 2)
                        End If

                        ASCDATA1.ExecuteSQL(Sql)
                    End If
                Next
            End If

            Call Write_Group_Record(aRC, "Recap", "")
        End If


        ' Bring Summarized Result Set in from Oracle

        Build_Report_File_Pre_Ora2ADO(TT)

        Dim sqlx As String = ""
        For i As Integer = 1 To 9
            sqlx &= " and G" & CStr(i) & " = '" & aRC & "'"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & TT & ASCMAIN1.SQL_Add_WHERE(sqlx))
        ' the code above was put in place to avoid duplicate totals when running the 12 month

        Create_TDA(dst.Tables.Add("ASTSRPT1"), ASTSRPT1, "Select * from " & TT)
        Fill_Records("ASTSRPT1")
        tblASTSRPT1 = dst.Tables("ASTSRPT1")
        ' tblASTSRPT1 = ASCDATA1.GetDataTable("Select * from " & TT, "ASTSRPT1")


        ' Add a *- Not on File -* Group Record for all Codes without a Description in ASTGROUP

        For i As Integer = 1 To COLUMN_NAMEs.Count
            Dim tbl As DataTable = ASCMAIN1.Distinct_Values("", tblASTSRPT1, "G" & CStr(i)) ' ASCDATA1.SelectDistinct(tblASTSRPT1, "G" & CStr(i))
            For Each row As DataRow In tbl.Rows
                Dim r As DataRow = tblASTGROUP.Rows.Find(row.Item(0))
                If r Is Nothing Then
                    If row.Item(0) & "" <> "" Then
                    Dim z As String = Mid$(row.Item(0), Len(COLUMN_CAPTIONs(i - 1)) + 2)
                    If InStr(z, "-") <> 0 Then
                        z = Mid$(z, InStr(z, "-") + 1)
                    End If
                    Dim rowASTGROUP As DataRow = tblASTGROUP.NewRow
                    rowASTGROUP.Item("GROUP_KEY") = row.Item(0)
                    rowASTGROUP.Item("GROUP_CODE") = z
                    rowASTGROUP.Item("GROUP_DESC") = "*- Not on File -*"
                    tblASTGROUP.Rows.Add(rowASTGROUP)
                End If
                End If
            Next
        Next i
        dst.Tables.Add(tblASTGROUP)
        '  dst.Tables.Add(tblASTSRPT1)

        ' Create Relationships between Group Descriptions and Group Code Fields

        For i As Integer = 1 To 9
            Dim COLUMN_NAME As String = "G" & Format$(i, "0")
            dst.Relations.Add(COLUMN_NAME, tblASTGROUP.Columns("GROUP_KEY"), tblASTSRPT1.Columns(COLUMN_NAME), False)
        Next

        Build_Report_File_Post_Process()
        Check_if_Empty("ASTSRPT1")

    End Sub

    Overridable Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)

    End Sub

    Overridable Sub Build_Report_File_Post_Process()

    End Sub

    Protected Sub Wrap_Up()
        Call ASCMAIN1.Progress("Wrapping Up")
        Application.DoEvents()

        Dim tblASTPAGE0 As DataTable = ASCDATA1.GetDataTable("*", "ASTPAGE0")
        Dim LINE_NO As Integer = 0
        For i As Integer = 1 To Page0.Count

            Dim LINE_DATA_ALL As String = Page0.Item(i - 1).ToString
            Do
                Dim rowASTPAGE0 As DataRow = tblASTPAGE0.NewRow
                Dim LINE_DATA As String
                If LINE_DATA_ALL.Length <= 255 Then
                    LINE_DATA = LINE_DATA_ALL
                    LINE_DATA_ALL = ""
                Else
                    LINE_DATA = Mid(LINE_DATA_ALL, 1, 255)
                    LINE_DATA_ALL = Mid(LINE_DATA_ALL, 256)
                End If
                LINE_NO += 1
                rowASTPAGE0.Item("LINE_NO") = LINE_NO
                rowASTPAGE0.Item("LINE_DATA") = LINE_DATA
                tblASTPAGE0.Rows.Add(rowASTPAGE0)
            Loop While LINE_DATA_ALL <> ""
        Next
        dst.Tables.Add(tblASTPAGE0)
    End Sub

#End Region

#Region "Main - Reporting"

    Sub Print_Report_Main()
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim XSD_FILENAME As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.ActiveForm.Name & ".XSD"
        If My.Computer.FileSystem.FileExists(XSD_FILENAME) Then
            My.Computer.FileSystem.DeleteFile(XSD_FILENAME)
        End If

        If Not chkDQ_Only.Checked Then
            F = New ASFSRPTV
            Call Print_Report()
            Show_Reports()
        End If

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Overridable Sub Print_Report()
        F.Generate_Report(RPT)
    End Sub

    Public Sub PB_Report_Parameters(Optional ByVal RECAP As String = "N")

        Dim PB_plus As Integer = 0
        If tblASTDSQLH.Rows.Count <> 0 Then
            PB_plus += 1
        End If
        If Recap_Report Then ' MAYBE SHOULD BE CHECKING grdASTRECAP.Rows.Count
            PB_plus += 1
        End If

        CR_params.Add("RECAP", RECAP)
        CR_params.Add("NEWPAGE", PAGE_BREAKs)
        CR_params.Add("RC", aRC)

        Dim HGs(7) As String
        For i As Integer = 1 To 7
            If i <= COLUMN_CAPTIONs.Count Then
                HGs(i) = COLUMN_CAPTIONs(i - 1)
            Else
                HGs(i) = ""
            End If
            CR_params.Add("HG" & CStr(i), HGs(i))
        Next

        CR_params.Add("LVLS", COLUMN_NAMEs.Count + PB_plus)
    End Sub

    Overrides Function Generate_Report( _
    ByVal RPT As String, _
    Optional ByVal RPT_TITLE As String = "", _
    Optional ByVal SUBT As String = "", _
    Optional ByVal RecordSelectionFormula As String = "", _
    Optional ByVal ExportFormat As String = "", _
    Optional ByVal TempExportFilenameBody As String = "", _
    Optional ByVal archive_this_report As Boolean = True)

        If ExportFormat = "" Then
            ExportFormat = ExportFormatDefault
        End If

        If TempExportFilenameBody = "" Then
            TempExportFilenameBody = ExportFilenameDefault
        End If

        If Not ArchiveReportsDefault Then
            archive_this_report = False
        End If

        Call ASCMAIN1.Progress("Now Printing " & IIf(RPT_TITLE <> "", RPT_TITLE, Me.Text))

        If PB_Report Then
            Dim RECAP As String = IIf(Absx1.chkFor("RECAP_LAST_LEVEL").Checked, "Y", "N")
            PB_Report_Parameters(RECAP)
        End If
        Dim REPORT_NO As String = F.Generate_Report(RPT, RPT_TITLE, SUBT, False, PB_Report, _
                                                    RecordSelectionFormula, _
                          ExportFormat, TempExportFilenameBody, archive_this_report)

        If REPORT_NO = "" Then
            If RWU = "R" Then RWU = "N"
        End If

        Return REPORT_NO
    End Function

    Sub Show_Reports()

        If ASCMAIN1.JOB_STREAM_CODE <> "" Then
            ASCMAIN1.JOB_STREAM_REPORT_XNO = XNO
        End If

        If ASCMAIN1.JOB_STREAM_CODE.Length = 0 AndAlso ASCMAIN1.JOB_STREAM_FORM_NAME <> MENU_ITEM_OBJECT Then
            If F.CRs.Count = 0 Then
            Else
                F.Show_Reports()
            End If

        Else

            If ASCMAIN1.JOB_STREAM_EXPORT_FOLDER = "" Then
                ASCMAIN1.JOB_STREAM_EXPORT_FOLDER = ASCMAIN1.Folders("Archive") & "JobStreamReports\"
            End If

            If ASCMAIN1.JOB_STREAM_EXPORT_FOLDER <> "" Then

                F.Dont_Show_Reports()
                For Each ROW As DataRow In F.tblASTSPRF1.Select("")

                    Dim REPORT_NO As String = ROW.Item(0)
                    Dim REPORT_FILENAME As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".RPT"

                    Dim RPT As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    RPT.Load(REPORT_FILENAME)

                    Dim FILENAME As String = Trim(Replace(Replace(Replace(RPT_TITLE, " ", ""), "-", ""), "/", ""))
                    FILENAME &= "" & Format(Now, "MMddyyHHmmss")

                    '    Dim SOMEFOLDER As String = ASCMAIN1.Folders("Temp")

                    For Each SFX As String In New String() {"PDF", "XLS"}

                        If ASCMAIN1.JOB_STREAM_EXPORT_OPTIONS = "*" Or (ASCMAIN1.JOB_STREAM_EXPORT_OPTIONS = SFX) Then
                            Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions
                            DestOpt.DiskFileName = ASCMAIN1.JOB_STREAM_EXPORT_FOLDER & FILENAME & "." & SFX

                            ASCMAIN1.JOB_STREAM_REPORT_FILES.Add(FILENAME & "." & SFX, ASCMAIN1.JOB_STREAM_EXPORT_FOLDER & FILENAME & "." & SFX)

                            With RPT.ExportOptions
                                .DestinationOptions = DestOpt
                                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                                If SFX = "PDF" Then
                                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                ElseIf SFX = "XLS" Then
                                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.Excel
                                End If
                            End With

                            Try
                                RPT.Export()
                            Catch ex As Exception
                                ' NOT IN JOBSTREAM
                                'MsgBox("Problem Report: " & REPORT_FILENAME & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load Report " & REPORT_NO)
                                RPT.Dispose()
                                Exit Sub
                            End Try
                        End If
                    Next
                Next

            End If

            F.Close()
            F = Nothing
        End If
    End Sub

    Sub SetParameterValue(ByVal pfName As String, ByVal pfValue As String)
        Dim Par As CrystalDecisions.Shared.ParameterValues
        Dim ParD As New CrystalDecisions.Shared.ParameterDiscreteValue()
        Par = ASCMAIN1.CR_RPT.DataDefinition.ParameterFields.Item(pfName).CurrentValues
        ParD.Value = pfValue
        Par.Add(ParD)
        ASCMAIN1.CR_RPT.DataDefinition.ParameterFields.Item(pfName).ApplyCurrentValues(Par)
    End Sub

#End Region

#Region "Dynamic SQL - Generation of SQL"

    Sub Get_SQL( _
    ByVal DATA_SOURCE As String, _
    Optional ByVal TABLE_NAME_temp As String = "")

        Call ASCMAIN1.Track("Extracts from Data Sources", DATA_SOURCE)

        Dim TABLE_NAME As String
        Dim COLUMN_NAME As String
        Dim rowASTDSQLC As DataRow

        Dim sql_SELECT_col As String
        Dim sql_Select_col_count As Integer = 0

        DATA_SOURCE = IIf(DATA_SOURCE = "", "*", DATA_SOURCE)

        Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
        sql_TABLE_NAME = rowASTDSQLB.Item("TABLE_NAME") & ""

        'If PB_Report Then
        ' SO WE ALWAYS LOOK FOR FORCED JOINS

            sql_SELECT_cols = ""
            sql_GROUP_BY_cols = ""
            sql_WHERE = ""
            sql_TABLE_NAMEs = ""
            sql_JOIN = ""

            ' Forced Joins - SHOULDN'T WE BE LOOKING AT J FOR THIS?

            For Each rowASTDSQLJ As DataRow In tblASTDSQLJ.Select("FORM_NAME = '" & FORM_NAME & "' and DATA_SOURCE = '" & DATA_SOURCE & "' and ALWAYS_JOIN = '1'")
                TABLE_NAME = rowASTDSQLJ.Item("TABLE_NAME")
                Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
            Next

            ' Sort

            For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE")
                sql_SELECT_col = ""
                'sql_GROUP_BY_col = ""

                COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")

                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})


                If rowASTDSQLC Is Nothing Then
                    rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
                End If

                TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
                If TABLE_NAME <> "" Then
                    If TABLE_NAME <> sql_TABLE_NAME Then
                        Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                    End If
                Else
                    TABLE_NAME = sql_TABLE_NAME
                End If

                'If rowASTDSQLC.Item("EXPRESSION_IND") & "" = "1" Then
                If rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                    sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
                    'sql_GROUP_BY_col = rowASTDSQLC.Item("COLUMN_EXPRESSION_Y") & ""
                Else
                    sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
                End If

                If rowASTDSQLA.Item("GROUP_ALL_OTHERS") & "" = "1" And rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                    sql_SELECT_col = "Case When " & sql_SELECT_col & " in (" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "') Then " & sql_SELECT_col & " else '*' End"
                End If

                ASCMAIN1.TACMAIN1.Get_Column_Expression_Exceptions(FORM_NAME, DATA_SOURCE, COLUMN_NAME, sql_SELECT_col) ' , sql_GROUP_BY_col)

                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col & " AS " & COLUMN_NAME
                'If sql_GROUP_BY_col = "" Then
                '    sql_GROUP_BY_col = sql_SELECT_col
                'End If
                'sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_GROUP_BY_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
                sql_Select_col_count = sql_Select_col_count + 1
            Next

            If COLUMN_NAMEs.Count > sql_Select_col_count Then
                sql_SELECT_col = COLUMN_NAMEs(COLUMN_NAMEs.Count - 1)
                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
            End If

            sql_SELECT_cols = Mid$(sql_SELECT_cols, 3)
            sql_GROUP_BY_cols = Mid$(sql_GROUP_BY_cols, 3)
        'End If

        ' Filter

        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("CODE_VALUES is Not Null AND CODE_VALUES <> ''")
            COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")

            rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})
            If rowASTDSQLC Is Nothing Then
                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
            End If

            If rowASTDSQLC.Item("NO_FILTER") & "" <> "1" Then
            TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
            If TABLE_NAME <> "" Then
                If TABLE_NAME <> sql_TABLE_NAME Then
                    Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                End If
            Else
                TABLE_NAME = sql_TABLE_NAME
            End If

            If rowASTDSQLC.Item("JOIN_SPECIAL") & "" = "1" Then
                sql_SELECT_col = GetSpecialSelectedJoin(COLUMN_NAME, DATA_SOURCE)
            ElseIf rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
            Else
                sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
            End If

            Dim in_or_equal As String
            Dim not_in_or_not_equal As String

            Dim CODE_VALUES_sql As String = "'" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "'"
            If InStr(CODE_VALUES_sql, ",") = 0 Then
                in_or_equal = "="
                not_in_or_not_equal = "<>"
            Else
                in_or_equal = "IN"
                not_in_or_not_equal = "NOT IN"
            End If

            If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                sql_WHERE = sql_WHERE & " AND (" & sql_SELECT_col & " IS NULL OR " & sql_SELECT_col & " " & not_in_or_not_equal & " (" & CODE_VALUES_sql & "))"
            Else
                sql_WHERE = sql_WHERE & " AND " & sql_SELECT_col & " " & in_or_equal & " (" & CODE_VALUES_sql & ")"
            End If
            End If
        Next

        'tblASTDSQLC = Nothing

        If PB_Report Then
            ' Pad sql_SELECT_cols for unused Group By's

            If COLUMN_NAMEs.Count < 9 Then
                For i As Integer = COLUMN_NAMEs.Count + 1 To 9
                    sql_SELECT_cols = sql_SELECT_cols & ", 'x' as G" & CStr(i)
                Next
                If COLUMN_NAMEs.Count = 0 Then
                    sql_SELECT_cols = Mid(sql_SELECT_cols, 3)
                End If
            End If

            ' Create Report Work File

            If ASTSRPT1 = "" Then
                COLUMN_NAMEs_appended = ""
                Dim sql_sum As String = ""
                Dim sql_sum_group_by As String = ""
                Dim sql As String = ""
                For i As Integer = 1 To 9
                    sql &= ",ASTSRPT1.G" & CStr(i)
                    sql_sum &= ",G" & CStr(i)
                Next
                sql = "Select " & Mid(sql, 2)
                sql_sum_group_by = Mid(sql_sum, 2)
                sql_sum = "Select " & sql_sum_group_by
                Dim ZTBL As String = ""
                If tblASTDSQLH.Rows.Count <> 0 Then
                    For Each ROW As DataRow In tblASTDSQLH.Select("", "COLUMN_SEQ")
                        Dim TABLE_NAME_appended_column As String = ROW.Item("TABLE_NAME") & ""
                        If TABLE_NAME_appended_column = "" Then
                            If TABLE_NAME_temp <> "" Then
                                TABLE_NAME_appended_column = TABLE_NAME_temp
                            Else
                                ' Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
                                TABLE_NAME_appended_column = rowASTDSQLB("TABLE_NAME") & ""
                            End If
                        End If
                        sql = sql & "," & TABLE_NAME_appended_column & "." & ROW.Item("COLUMN_NAME") & " " & ROW.Item("COLUMN_ALIAS")
                        If InStr(ZTBL, "," & TABLE_NAME_appended_column) = 0 Then
                            ZTBL &= "," & TABLE_NAME_appended_column
                        End If
                        Dim COLUMN_ALIAS As String = ROW.Item("COLUMN_ALIAS") & ""
                        If COLUMN_ALIAS = "" Then
                            COLUMN_ALIAS = ROW.Item("COLUMN_NAME")
                        End If
                        sql_sum &= "," & COLUMN_ALIAS
                        COLUMN_NAMEs_appended &= "," & COLUMN_ALIAS
                        sql_sum_group_by &= "," & COLUMN_ALIAS
                    Next
                End If
                If tblASTRECAP.Rows.Count <> 0 Then
                    COLUMN_NAME_RECAP_ROW_NO = ", ASTSRPT1_RECAP_ROW_NO"
                    sql = sql & ", TATWORK1.W_INT " & Mid(COLUMN_NAME_RECAP_ROW_NO, 3)
                    If InStr(ZTBL, "," & "TATWORK1") = 0 Then
                        ZTBL &= "," & "TATWORK1"
                    End If
                    sql_sum &= COLUMN_NAME_RECAP_ROW_NO
                    sql_sum_group_by &= COLUMN_NAME_RECAP_ROW_NO

                    dst.Tables.Add(tblASTRECAP.Copy)
                Else
                    ' CHECKBOX SAYS RECAP LAST LEVEL
                    Dim LVLs As Integer = tblASTDSQLA.Select("ISNULL(SEQUENCE,0) <> 0").Length
                    COLUMN_NAME_RECAP_ROW_NO = ", G" & CStr(LVLs)
                End If
                ASTSRPT1_sum_columns = ""
                For Each KEY As String In COLUMN_NAME_sum.Keys
                    Select Case COLUMN_NAME_sum(KEY)
                        Case "QTY"
                            sql = sql & ",ASTSRPT1.W_QTY " & KEY
                        Case "AMT"
                            sql = sql & ",ASTSRPT1.W_AMT " & KEY
                        Case "DEC"
                            sql = sql & ",ASTSRPT1.W_DEC " & KEY
                        Case Else
                            MsgBox("Invalid Data Type")
                            Stop
                    End Select
                    sql_sum = sql_sum & ",SUM(" & KEY & ") " & KEY
                    ASTSRPT1_sum_columns &= ",SUM(" & KEY & ") " & KEY
                Next

                sql = sql & " from ASTSRPT1" & ZTBL & " where ROWNUM < 1"
                ASTSRPT1 = ASCMAIN1.Temp_Table(sql)
                sql_sum = sql_sum & " from " & ASTSRPT1 & " group by " & sql_sum_group_by
                ASTSRPT1_sql_sum = sql_sum
            End If
        End If

    End Sub

    Private Sub Get_SQL_Join_Criteria(ByVal TABLE_NAME As String, ByVal DATA_SOURCE As String)

        Dim rowASTDSQLJ As DataRow = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, TABLE_NAME})
        If rowASTDSQLJ Is Nothing Then
            rowASTDSQLJ = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, "*", TABLE_NAME})
        End If
        Dim TABLE_NAME_IS_ALIAS_FOR As String = rowASTDSQLJ.Item("TABLE_NAME_IS_ALIAS_FOR") & ""

        Dim TABLE_NAME_WITH_ALIAS As String = TABLE_NAME
        If TABLE_NAME_IS_ALIAS_FOR <> "" Then
            TABLE_NAME_WITH_ALIAS = TABLE_NAME_IS_ALIAS_FOR & " " & TABLE_NAME
        End If

        If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME_WITH_ALIAS) <> 0 Then
            Exit Sub
        Else
            sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME_WITH_ALIAS
        End If

        Dim COLUMN_NAME As String
        Dim sql As String

        Dim JOIN_TYPE As String = ""
        If rowASTDSQLJ IsNot Nothing AndAlso rowASTDSQLJ.Item("OUTER_JOIN") & "" = "1" Then
            JOIN_TYPE = "(+)"
        End If

        Dim drsASTDSQLD() As DataRow
        sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
            & "' and DATA_SOURCE = '" & DATA_SOURCE & "'"
        drsASTDSQLD = tblASTDSQLD.Select(sql)
        If drsASTDSQLD.Length = 0 Then
            sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
                & "' and DATA_SOURCE = '" & "*" & "'"
            drsASTDSQLD = tblASTDSQLD.Select(sql)
        End If

        If drsASTDSQLD.Length = 0 Then
            Dim tbl As DataTable = ASCDATA1.GetDataTable("*", TABLE_NAME, -1, False)
            ReDim drsASTDSQLD(tbl.PrimaryKey.Length - 1)
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                Dim row As DataRow = tblASTDSQLD.NewRow
                Dim dc As DataColumn = tbl.PrimaryKey(i)
                row.Item("FORM_NAME") = FORM_NAME
                row.Item("DATA_SOURCE") = DATA_SOURCE
                row.Item("TABLE_NAME") = TABLE_NAME
                row.Item("COLUMN_NAME") = dc.ColumnName
                drsASTDSQLD(i) = row
            Next
        End If

        Dim rowASTDSQLD As DataRow
        For i As Integer = 0 To UBound(drsASTDSQLD)
            rowASTDSQLD = drsASTDSQLD(i)

            If rowASTDSQLD.Item("TABLE_NAME_JOIN") & "" <> "" Then
                Call Get_SQL_Join_Criteria(rowASTDSQLD.Item("TABLE_NAME_JOIN"), DATA_SOURCE)
            End If

            COLUMN_NAME = rowASTDSQLD.Item("COLUMN_NAME")
            Dim TABLE_NAME_JOIN As String = rowASTDSQLD.Item("TABLE_NAME_JOIN") & ""
            If TABLE_NAME_JOIN = "" Then
                TABLE_NAME_JOIN = sql_TABLE_NAME
            End If

            sql = TABLE_NAME & "." & COLUMN_NAME & JOIN_TYPE & " = "

            'sql = JOIN_TYPE & TABLE_NAME & " ON " & COLUMN_NAME & " = "
            ''If rowASTDSQLD.Item("EXPRESSION_IND") & "" = "1" Then
            '' sql = sql & rowASTDSQLD.Item("COLUMN_NAME_JOIN")
            '' Else
            Dim COLUMN_NAME_JOIN As String = rowASTDSQLD.Item("COLUMN_NAME_JOIN") & ""
            If COLUMN_NAME_JOIN.Contains(",") Or COLUMN_NAME_JOIN.Contains(".") Or COLUMN_NAME_JOIN.Contains("(") Or COLUMN_NAME_JOIN.Contains(")") Or COLUMN_NAME_JOIN.Contains("'") Then
                ' COLUMN_NAME_JOIN IS AN EXPRESSION, LIKE AHA.SAR12MO1.SOTSREP1
                ' LOOKS LIKE THE CODE ABOVE USED TO RELY ON A FIELD EXPRESSION_IND TO ACCOMPLISH THIS - WHICH SOUNDS LIKE A BETTER IDEA
                ' PERHAPS WILL IMPLEMENT THIS LATER
                sql = sql & COLUMN_NAME_JOIN
                sql_JOIN = sql_JOIN & " AND " & sql
            Else
                If COLUMN_NAME_JOIN = "" Then
                    COLUMN_NAME_JOIN = COLUMN_NAME
                End If
                sql = sql & TABLE_NAME_JOIN & "." & COLUMN_NAME_JOIN
                'End If

                sql_JOIN = sql_JOIN & " AND " & sql
            End If
        Next

    End Sub

    Private Sub Get_SQL_Join_Criteria_Special(ByVal FORM_NAME As String, ByVal TABLE_NAME As String, ByVal DATA_SOURCE As String)

        'If TABLE_NAME = "" Then Exit Sub

        'jz = ""
        '' Special Conditions
        'Select Case FORM_NAME
        '    Case "SOFWHOD1"
        '        Select Case TABLE_NAME
        '            Case "ARTCUST1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ARTCUST1.CUST_CODE (+) = Y2.CUST_CODE"
        '                Else
        '                    jz = "ARTCUST1.CUST_CODE (+) = X.CUST_CODE"
        '                End If

        '            Case "ICTITEM1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ICTITEM1.ITEM_CODE (+) = Y1.ITEM_CODE"
        '                Else
        '                    jz = "ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        '                End If

        '        End Select

        'End Select

        'If jz <> "" And InStr(1, sqljoin, jz) = 0 Then
        '    sqljoin = sqljoin & " AND " & jz
        '    If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME) = 0 Then 'Exit Sub
        '        sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME
        '    End If
        'End If

    End Sub

    Private Function GetSpecialSelectedJoin(ByVal COLUMN_NAME As String, ByVal DATA_SOURCE As String) As String
        Dim z As String = ""

        'Select Case FORM_NAME
        '    Case "SOFSLSF1"
        '        Select Case COLUMN_NAME
        '            Case "ITEM_BRAND_CODE"
        '        End Select
        '    Case "SOFWHOD1"
        '        Select Case COLUMN_NAME
        '            Case "CUST_CODE"
        '                If DATA_SOURCE = "A" Then
        '                    z = "Y2.CUST_CODE"
        '                Else
        '                    z = "X.CUST_CODE"
        '                End If

        '            Case "ITEM_CODE"
        '                If DATA_SOURCE = "C" Or DATA_SOURCE = "B" Then
        '                    z = ""
        '                ElseIf DATA_SOURCE = "A" Then
        '                    z = "Y1.ITEM_CODE"
        '                Else
        '                    z = "X.ITEM_CODE"
        '                End If
        '        End Select
        'End Select

        GetSpecialSelectedJoin = z

    End Function
#End Region

#Region "ASTGROUP"

    Sub Write_Group_Record(ByVal GROUP_KEY As String, ByVal GROUP_CODE As String, ByVal GROUP_DESC As String)
        Dim rowASTGROUP As DataRow = tblASTGROUP.NewRow
        If GROUP_KEY = aRC Then
            rowASTGROUP.Item("GROUP_KEY") = GROUP_KEY
        Else
            If InStr(GROUP_KEY, ":") = 0 Then
                rowASTGROUP.Item("GROUP_KEY") = GROUP_KEY & ":" & GROUP_CODE
            Else
                rowASTGROUP.Item("GROUP_KEY") = GROUP_KEY
            End If
        End If
        rowASTGROUP.Item("GROUP_CODE") = GROUP_CODE
        rowASTGROUP.Item("GROUP_DESC") = GROUP_DESC
        tblASTGROUP.Rows.Add(rowASTGROUP)
    End Sub

    Sub Get_Group_Desc_All()
        Dim i As Integer
        Dim j As Integer
        Dim z As String
        Dim sql As String
        Dim COLUMN_NAME As String
        Dim GROUP_KEY As String = ""

        For i = 1 To COLUMN_NAMEs.Count
            COLUMN_NAME = COLUMN_NAMEs(i - 1)
            sql = ASCMAIN1.TACMAIN1.Get_Code_SQL_X(FORM_NAME, COLUMN_NAME, GROUP_KEY)

            ' If no match then look for : separating field
            If sql = "" And InStr(1, COLUMN_NAME, ":") > 0 Then
                j = InStr(1, COLUMN_NAME, ":")
                z = Mid(COLUMN_NAME, 1, j - 1)
                sql = ASCMAIN1.TACMAIN1.Get_Code_SQL_X(FORM_NAME, z, GROUP_KEY)
            End If

            If sql = "" Then
                Dim row As DataRow
                row = ASCDATA1.GetDataRow("SELECT * FROM ASTDSQLG WHERE TABLE_NAME LIKE '" & Mid(FORM_NAME, 1, 2) & "%' AND COLUMN_NAME_CODE = '" & COLUMN_NAME & "'")
                If row Is Nothing Then
                    row = ASCDATA1.GetDataRow("SELECT * FROM ASTDSQLG WHERE COLUMN_NAME_CODE = '" & COLUMN_NAME & "'")
                End If
                If row IsNot Nothing Then
                    sql = "Select " & row.Item("COLUMN_NAME_CODE") & ", " & row.Item("COLUMN_NAME_DESC") & " from " & row.Item("TABLE_NAME")
                    If row.Item("COLUMN_NAME_KEY") & "" <> "" Then
                        GROUP_KEY = row.Item("COLUMN_NAME_KEY")
                    End If
                End If
            End If

            If sql <> "" Then
                sql = "Select Distinct '" & COLUMN_CAPTIONs(i - 1) & ":' " & ASCMAIN1.DBS_CONCAT & " " & GROUP_KEY & ", " & Mid$(sql, 8)
                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    For Each row As DataRow In ASCDATA1.GetDataTable(sql).Rows
                        Dim rowASTGROUP As DataRow = tblASTGROUP.NewRow
                        For j = 0 To tblASTGROUP.Columns.Count - 1
                            rowASTGROUP.Item(j) = row.Item(j)
                        Next
                        tblASTGROUP.Rows.Add(rowASTGROUP)
                    Next
                Else
                    If InStr(sql.ToUpper, " WHERE ") = 0 Then
                        sql = sql & " where " & GROUP_KEY & " in (Select Distinct G" & CStr(i) & " from " & ASTSRPT1 & ")"
                    Else
                        sql = sql & " and " & GROUP_KEY & " in (Select Distinct G" & CStr(i) & " from " & ASTSRPT1 & ")"
                        'sql = "Select * from (" & sql & ") where " & GROUP_KEY & " in (Select Distinct G" & CStr(i) & " from " & ASTSRPT1 & ")"
                    End If
                    Dim tbl1 As DataTable = ASCDATA1.GetDataTable(sql)
                    tbl1.Columns(0).ColumnName = "GROUP_KEY"
                    tbl1.Columns(1).ColumnName = "GROUP_CODE"
                    tbl1.Columns(2).ColumnName = "GROUP_DESC"
                    tbl1.PrimaryKey = New DataColumn() {tbl1.Columns("GROUP_KEY")}
                    tblASTGROUP.Merge(tbl1)
                End If
            Else

                Dim GROUP_CODEs As New ArrayList
                Dim GROUP_DESCs As New ArrayList
                ASCMAIN1.TACMAIN1.Write_Group_Record_X(COLUMN_CAPTIONs(i - 1), COLUMN_NAME, GROUP_CODEs, GROUP_DESCs)
                If GROUP_CODEs.Count <> 0 Then
                    For j = 1 To GROUP_CODEs.Count
                        Write_Group_Record(COLUMN_CAPTIONs(i - 1), GROUP_CODEs(j - 1), GROUP_DESCs(j - 1))
                    Next
                End If
            End If
        Next i
    End Sub
#End Region

#Region "Lists"

    Private Sub cmdListRetrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListRetrieve.Click
        Retreive_List()
    End Sub

    Sub Retreive_List()
        COLUMN_NAME = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LIST_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.SQL &= " where COLUMN_NAME = '" & COLUMN_NAME & "'"
            ASCMAIN1.CodeSelector.SQL &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or NVL(LIST_SHAREABLE,'0') = '1')"
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections = 1 Then
                LIST_CODE = ASCMAIN1.CodeSelector.SelectedRows(0).Item("LIST_CODE")
                Dim i As Integer
                If grdSetup.ActiveRow.Cells("CODE_VALUES").Text <> "" Then
                    Dim frmASFMSGBF As New ASFMSGBF
                    i = frmASFMSGBF.Get_opt_from_User("Load this List of Codes", New String() {"By Replacing the Existing List of Codes", "By Appending to the Existing List of Codes"}, 0, "Retrieve Code List Option")
                    frmASFMSGBF.Dispose()
                End If

                Call Load_Code_List(i = 0)
            End If
        End If
    End Sub

    Sub Load_Code_List(ByVal replace_codes As Boolean)
        Dim tblASTLIST1 As DataTable = ASCDATA1.GetDataTable("Select * from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
        Dim tblASTLIST2 As DataTable = ASCDATA1.GetDataTable("Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "' order by CODE_VALUE")

        Dim CODE_VALUES As String
        If replace_codes Then
            CODE_VALUES = ""
        Else
            CODE_VALUES = grdSetup.ActiveRow.Cells("CODE_VALUES").Value
        End If

        For Each dr As DataRow In tblASTLIST2.Rows
            If Not InStr("," & CODE_VALUES & ",", "," & dr.Item("CODE_VALUE") & ",") Then
                CODE_VALUES &= "," & dr.Item("CODE_VALUE")
            End If
        Next
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid(CODE_VALUES, 2)
        grdSetup.UpdateData()


        Call Show_grd()

        LIST_CODE = tblASTLIST1.Rows(0).Item("LIST_CODE")
        LIST_DESC = tblASTLIST1.Rows(0).Item("LIST_DESC")
        txtList.Text = LIST_DESC
        chkListShareable.Checked = (tblASTLIST1.Rows(0).Item("LIST_SHAREABLE") & "" = "1")
        chkListModifiable.Checked = (tblASTLIST1.Rows(0).Item("LIST_MODIFIABLE") & "" = "1")
        chkListShareable.Enabled = (tblASTLIST1.Rows(0).Item("INIT_OPER") = ASCMAIN1.USER_ID)
        chkListModifiable.Enabled = (tblASTLIST1.Rows(0).Item("INIT_OPER") = ASCMAIN1.USER_ID)
    End Sub

    Private Sub cmdListSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListSaveAs.Click
        If Trim(txtList.Text) = "" Then
            MsgBox("You Must Enter a List Description", MsgBoxStyle.OkOnly, "Cannot Save List")
            Exit Sub
        End If

        SaveAsList(txtList.Text)
    End Sub

    Sub SaveAsList(LIST_DESC_new As String)

        COLUMN_NAME = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        Dim CODE_VALUES As String = grdSetup.ActiveRow.Cells("CODE_VALUES").Value

        If CODE_VALUES = "" Then
            MsgBox("No Code Values in the List")
            'Stop
            'ABS.UI.MessageBox.Show("No Code Values in the List", ABS.UI.Types.MessageBoxButton.OKOnly, "Cannot Save List")
            Exit Sub
        End If

        Dim i As Integer = 0
        If LIST_CODE = "" Then
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
        Else
            Dim frmASFMSGBF As New ASFMSGBF
            If Not chkListModifiable.Enabled And Not chkListModifiable.Checked Then
                If LIST_DESC_new = LIST_DESC Then
                    MsgBox("You must change the Description of this List")
                    'Stop
                    'ABS.UI.MessageBox.Show("You must change the Description of this List" & vbCr & " in order to Save it (as one of your own Lists)", ABS.UI.Types.MessageBoxButton.OKOnly, "Cannot Save List")
                    Exit Sub
                End If
                i = 0
            Else
                i = frmASFMSGBF.Get_opt_from_User("Save this List of Codes", New String() {"As a New List", "By Replacing Existing List"}, 0, "Save Code List Option")
            End If
            If i = -1 Then
                Exit Sub
            ElseIf i = 0 Then
                LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            Else
                ASCMAIN1.sql = "Delete from ASTLIST2 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If
        End If

        Dim tblASTLIST1 As New DataTable
        ASCMAIN1.sql = "Select * from ASTLIST1 where COLUMN_NAME = '" & COLUMN_NAME & "' and LIST_CODE = '" & LIST_CODE & "'"
        With ASCDATA1.GetDataAdapter(tblASTLIST1, "ASTLIST1", "", True)
            If i = 1 Then
                tblASTLIST1.Rows(0).Item("LIST_DESC") = LIST_DESC_new
                tblASTLIST1.Rows(0).Item("LIST_SHAREABLE") = CStr(Abs(Val(chkListShareable.Checked)))
                tblASTLIST1.Rows(0).Item("LIST_MODIFIABLE") = CStr(Abs(Val(chkListModifiable.Checked)))
                tblASTLIST1.Rows(0).Item("LAST_OPER") = ASCMAIN1.USER_ID
                tblASTLIST1.Rows(0).Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            Else
                Dim rowASTLIST1 As DataRow = tblASTLIST1.NewRow
                rowASTLIST1.Item("COLUMN_NAME") = COLUMN_NAME
                rowASTLIST1.Item("LIST_CODE") = LIST_CODE
                rowASTLIST1.Item("LIST_DESC") = LIST_DESC_new
                rowASTLIST1.Item("LIST_SHAREABLE") = CStr(Abs(Val(chkListShareable.Checked)))
                rowASTLIST1.Item("LIST_MODIFIABLE") = CStr(Abs(Val(chkListModifiable.Checked)))
                rowASTLIST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowASTLIST1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                rowASTLIST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowASTLIST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                tblASTLIST1.Rows.Add(rowASTLIST1)
                .Update(tblASTLIST1)
                .Dispose()
            End If
        End With

        Dim tblASTLIST2 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTLIST2, "ASTLIST2", "*", True, -1, False)
            For Each CODE_VALUE As String In Split(CODE_VALUES, ",")
                Dim rowASTLIST2 As DataRow = tblASTLIST2.NewRow
                rowASTLIST2.Item("LIST_CODE") = LIST_CODE
                rowASTLIST2.Item("CODE_VALUE") = CODE_VALUE
                tblASTLIST2.Rows.Add(rowASTLIST2)
                .Update(tblASTLIST2)
            Next
            .Dispose()
        End With

        MsgBox("Code List '" & LIST_DESC_new & "' has been Saved", MsgBoxStyle.OkOnly, "Success")
        Call Load_Code_List(True)

    End Sub
#End Region

#Region "Settings"

    Private Sub chkMySettingsOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMySettingsOnly.CheckedChanged
        If chkMySettingsOnly.CheckState Then
            grdASTROPT1.Rows.ColumnFilters("INIT_OPER").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, ASCMAIN1.USER_ID)
        Else
            grdASTROPT1.Rows.ColumnFilters("INIT_OPER").FilterConditions.Clear()
            grdASTROPT1.Rows.Refresh(Infragistics.Win.UltraWinGrid.RefreshRow.ReloadData)
        End If
    End Sub

    Private Sub cmdSaveSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveSettings.Click

        If SET_ID <> "" Then
            Dim rowASTROPT1 As DataRow = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})
            If rowASTROPT1.Item("INIT_OPER") <> ASCMAIN1.USER_ID Then
                SET_ID = ""
            Else

                Select Case MsgBox("Update the Current Setting (Y) or Create a New One (N)?", MsgBoxStyle.YesNoCancel, "Save Setting Option")
                    Case MsgBoxResult.Yes
                    Case MsgBoxResult.No
                        SET_ID = ""
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If
        End If

        Call Save_Settings(SET_ID)
        grdASTROPT1.ActiveRow = grdASTROPT1.Rows.GetRowWithListIndex(tblASTROPT1.Rows.IndexOf(tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})))
    End Sub

    Sub Retrieve_Settings()

        Dim SET_CTL_NAME As String
        Dim SET_CTL_TYPE As String
        Dim SET_CTL_TAG As String
        Dim SET_CTL_DATA As String

        Dim rowASTROPT1 As DataRow = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})

        Dim sql As String

        If SET_ID = "0000000000" Then
            txtDescription.Text = ""
        Else
            txtDescription.Text = grdASTROPT1.ActiveRow.Cells("SET_DESC").Text
        End If

        Call Clear_grdSetup(True)
        tblASTRECAP.Rows.Clear()

        ASCMAIN1.sql = "Select * from ASTROPT4 where FORM_NAME = '" & FORM_NAME & "' and SET_ID = '" & SET_ID & "'"
        tblASTROPT4 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTROPT4", 3)

        sql = "Select * from ASTROPT2 where FORM_NAME = '" & FORM_NAME & "'"
        sql = sql & " and SET_ID = '" & SET_ID & "'"
        sql = sql & " and XNO is Null"
        For Each rowASTROPT2 As DataRow In ASCDATA1.GetDataTable(sql).Select("", "SET_CTL_TAG")
            SET_CTL_NAME = rowASTROPT2.Item("SET_CTL_NAME") & ""
            SET_CTL_TYPE = rowASTROPT2.Item("SET_CTL_TYPE") & ""
            SET_CTL_TAG = rowASTROPT2.Item("SET_CTL_TAG") & ""
            SET_CTL_DATA = rowASTROPT2.Item("SET_CTL_DATA") & ""

            Dim gDR As DataRow

            If SET_CTL_NAME = "grdSetup" Then
                If SET_CTL_TAG = "" Then
                    Dim GRDCOLS() As String = Split(SET_CTL_DATA, vbTab)
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(GRDCOLS(0))
                    If gDR IsNot Nothing Then
                        If Val(GRDCOLS(1) & "") <> 0 Then
                            gDR.Item("SEQUENCE") = Val(GRDCOLS(1) & "")
                        End If
                        gDR.Item("PAGE_BREAK") = GRDCOLS(2)
                        gDR.Item("EXCLUDE") = GRDCOLS(3)
                        gDR.Item("GROUP_ALL_OTHERS") = GRDCOLS(4)
                    End If
                Else
                    Dim COLUMN_NAME As String = SET_CTL_TAG
                    gDR = DirectCast(grdSetup.DataSource, DataTable).Rows.Find(COLUMN_NAME)
                    If gDR.Item("CODE_VALUES") & "" = "" Then
                        gDR.Item("CODE_VALUES") = SET_CTL_DATA
                    Else
                        gDR.Item("CODE_VALUES") &= "," & SET_CTL_DATA
                    End If
                End If

            ElseIf SET_CTL_NAME = "grdASTRECAP" Then
                Dim R() As String = Split(SET_CTL_DATA, vbTab)
                If tblASTRECAP.Columns.Count > R.Length Then
                    ReDim Preserve R(tblASTRECAP.Columns.Count - 1)
                End If
                For i As Int16 = 0 To R.Length - 1
                    If R(i) = "" Then
                        If tblASTRECAP.Columns(i).ColumnName = "ASTSRPT1_RECAP_ROW_CAPTION" Or tblASTRECAP.Columns(i).ColumnName = "ASTSRPT1_RECAP_ROW_CALC" Then
                            'R(i) = "" ' grdASTRECAP.DisplayLayout.ValueLists(tblASTRECAP.Columns(i).ColumnName).ValueListItems(0).DataValue & ""
                        Else
                            R(i) = grdASTRECAP.DisplayLayout.ValueLists(tblASTRECAP.Columns(i).ColumnName).ValueListItems(0).DataValue & ""
                        End If
                    End If
                Next
                tblASTRECAP.Rows.Add(R)
            Else
                Dim C As Control = Absx1.CtlFor(SET_CTL_TAG, True)
                If C IsNot Nothing Then
                    Select Case SET_CTL_TYPE
                        Case "UltraCheckEditor"
                            Absx1.chkFor(SET_CTL_TAG).Checked = (SET_CTL_DATA = "True")
                        Case "UltraOptionSet"
                            Absx1.optFor(SET_CTL_TAG).Value = SET_CTL_DATA
                        Case "UltraTrackBar"
                            If SET_CTL_DATA <> "" Then
                                DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraTrackBar).Value = SET_CTL_DATA
                            End If
                        Case "ABSCheckBox"
                            DirectCast(Absx1.CtlFor(SET_CTL_TAG), ABSCS.ABSCheckBox).ABSChecked = SET_CTL_DATA
                        Case "UltraCombo"
                            Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinGrid.UltraCombo)
                            'cmbctl.Text = SET_CTL_DATA
                            cmbctl.Value = SET_CTL_DATA
                            If SET_CTL_TAG = "RYP" Or SET_CTL_TAG = "RYP0" Or SET_CTL_TAG = "RYP1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" AndAlso rowASTROPT1.Item("SET_YP_BASE") & "" <> "" Then
                                    Dim RYP As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NP As Integer = ASCMAIN1.Period_Diff(rowASTROPT1.Item("SET_YP_BASE") & "", RYP)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, NP)), 1, 16)
                                End If
                            End If
                            If SET_CTL_TAG = "RYW" Or SET_CTL_TAG = "RYW0" Or SET_CTL_TAG = "RYW1" Then
                                If SET_ID <> "0000000000" AndAlso rowASTROPT1.Item("SET_YP_REL") & "" = "1" AndAlso rowASTROPT1.Item("SET_YW_BASE") & "" <> "" Then
                                    Dim RYW As String = Mid(SET_CTL_DATA, 1, 4) & Mid(SET_CTL_DATA, 6, 2)
                                    Dim NW As Integer = ASCMAIN1.Week_Diff(rowASTROPT1.Item("SET_YW_BASE") & "", RYW)
                                    cmbctl.Text = Mid(ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(ASCMAIN1.CYW, NW)), 1, 17)
                                End If
                            End If
                        Case "UltraComboEditor"
                            Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(Absx1.CtlFor(SET_CTL_TAG), UltraWinEditors.UltraComboEditor)
                            cbectl.Value = SET_CTL_DATA
                        Case Else
                            Absx1.CtlFor(SET_CTL_TAG).Text = SET_CTL_DATA
                    End Select

                End If
            End If
        Next

        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If
        If grdASTRECAP.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

    End Sub

    Sub Save_Settings(ByRef SET_ID As String, Optional ByVal XNO As String = "")

        Call BeginTrans()

        Dim rowASTROPT1 As DataRow

        If XNO <> "" Then
            ' It is ok that ASTROPT1 does not get recorded 
            ' (although ASTROPT2 does get recorded) here.
            ' When XNO <> "", ASTOPST1 may serve as a "header" for ASTROPT2, 
            ' and in fact does, when we view Execution History
        Else
            Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD
            If SET_ID = "" Then
                rowASTROPT1 = tblASTROPT1.NewRow()
                SET_ID = ASCMAIN1.Next_Control_No("ASTROPT1.SET_ID")
                rowASTROPT1.Item("FORM_NAME") = FORM_NAME
                rowASTROPT1.Item("SET_ID") = SET_ID
                rowASTROPT1.Item("SET_YP_BASE") = ASCMAIN1.CYP
                rowASTROPT1.Item("SET_YP_REL") = "1"
                rowASTROPT1.Item("SET_ALLOW_OTHERS") = "0"
                rowASTROPT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowASTROPT1.Item("INIT_DATE") = LAST_DATE
                tblASTROPT1.Rows.Add(rowASTROPT1)
            Else
                rowASTROPT1 = tblASTROPT1.Rows.Find(New Object() {FORM_NAME, SET_ID})
                ASCMAIN1.sql = "Delete from ASTROPT2 " _
                    & " where FORM_NAME = '" & FORM_NAME & "'" _
                    & " and SET_ID = '" & SET_ID & "'" _
                    & " and XNO is Null"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If

            If SET_ID = "0000000000" Then
                'rowASTROPT1.Item("SET_DESC") = "{Defaults}"
            Else

                If txtDescription.Text = "" Then
                    rowASTROPT1.Item("SET_DESC") = "{Enter a Description for these Settings}"
                Else
                    rowASTROPT1.Item("SET_DESC") = txtDescription.Text
                End If
                rowASTROPT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowASTROPT1.Item("LAST_DATE") = LAST_DATE
                'ASCMAIN1.sql = "Select * from ASTROPT1 " _
                '    & " where FORM_NAME = '" & FORM_NAME & "'"
                With ASCDATA1.GetDataAdapter(New DataTable, "ASTROPT1", "*", True, 2, False)
                    .Update(tblASTROPT1)
                    .Dispose()
                End With
            End If

        End If

        If SET_ID Is Nothing Then
            SET_ID = ""
        End If
        If SET_ID & "" <> "" Then
            Dim tblASTROPT4_proxy As New DataTable
            ASCMAIN1.sql = "Delete from ASTROPT4 where FORM_NAME = :PARM1 and SET_ID = :PARM2"

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {FORM_NAME, SET_ID})
            With ASCDATA1.GetDataAdapter(tblASTROPT4_proxy, "ASTROPT4", "*", True, 0, False)
                For Each rowASTROPT4 As DataRow In tblASTROPT4.Rows
                    Dim rowASTROPT4_proxy As DataRow = tblASTROPT4_proxy.NewRow
                    'rowASTROPT4_proxy.ItemArray = rowASTROPT4.ItemArray
                    rowASTROPT4_proxy.Item("FORM_NAME") = FORM_NAME
                    rowASTROPT4_proxy.Item("SET_ID") = SET_ID
                    rowASTROPT4_proxy.Item("USER_ID") = ASCMAIN1.USER_ID
                    tblASTROPT4_proxy.Rows.Add(rowASTROPT4_proxy)
                Next
                .Update(tblASTROPT4_proxy)
                .Dispose()
            End With

        End If

        Dim rowASTROPT2 As DataRow
        Dim tblASTROPT2 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTROPT2, "ASTROPT2", "*", True, 0, False)

            Call Save_Settings_ctls(UltraTabPageControl2, FORM_NAME, SET_ID, XNO, tblASTROPT2)

            For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSetup.Rows
                rowASTROPT2 = tblASTROPT2.NewRow()
                With rowASTROPT2
                    .Item("FORM_NAME") = FORM_NAME
                    .Item("SET_ID") = SET_ID
                    .Item("SET_CTL_NAME") = grdSetup.Name
                    .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                    .Item("SET_CTL_DATA") = gr.Cells("COLUMN_NAME").Text _
                                & vbTab & gr.Cells("SEQUENCE").Text _
                                & vbTab & gr.Cells("PAGE_BREAK").Value _
                                & vbTab & gr.Cells("EXCLUDE").Value _
                                & vbTab & gr.Cells("GROUP_ALL_OTHERS").Value
                    .Item("SET_CTL_TAG") = ""
                    .Item("XNO") = XNO
                End With
                tblASTROPT2.Rows.Add(rowASTROPT2)

                If gr.Cells("CODE_VALUES").Text <> "" Then
                    Dim CODE_VALUES() As String = Split(gr.Cells("CODE_VALUES").Text, ",")
                    For Each CODE_VALUE As String In CODE_VALUES
                        rowASTROPT2 = tblASTROPT2.NewRow()
                        With rowASTROPT2
                            .Item("FORM_NAME") = FORM_NAME
                            .Item("SET_ID") = SET_ID
                            .Item("SET_CTL_NAME") = grdSetup.Name
                            .Item("SET_CTL_TYPE") = grdSetup.GetType.Name
                            .Item("SET_CTL_DATA") = CODE_VALUE
                            .Item("SET_CTL_TAG") = gr.Cells("COLUMN_NAME").Text
                            .Item("XNO") = XNO
                        End With
                        tblASTROPT2.Rows.Add(rowASTROPT2)
                    Next
                End If
            Next


            For Each gr As Infragistics.Win.UltraWinGrid.UltraGridRow In grdASTRECAP.Rows
                rowASTROPT2 = tblASTROPT2.NewRow()
                With rowASTROPT2
                    .Item("FORM_NAME") = FORM_NAME
                    .Item("SET_ID") = SET_ID
                    .Item("SET_CTL_NAME") = grdASTRECAP.Name
                    .Item("SET_CTL_TYPE") = grdASTRECAP.GetType.Name
                    Dim ASTRECAP_row As String = ""
                    For i As Integer = 0 To grdASTRECAP.DisplayLayout.Bands(0).Columns.Count - 1
                        ASTRECAP_row &= vbTab & gr.Cells(i).Value
                    Next
                    .Item("SET_CTL_DATA") = Mid(ASTRECAP_row, 2)
                    .Item("SET_CTL_TAG") = ""
                    .Item("XNO") = XNO
                End With
                tblASTROPT2.Rows.Add(rowASTROPT2)
            Next

            .Update(tblASTROPT2)

            tblASTROPT2.Dispose()
            .Dispose()
        End With

        Call CommitTrans()

        Sort_grdColumns(grdASTROPT1, "LAST_DATE".ToLower)

        If XNO = "" And SET_ID <> "0000000000" Then
            MsgBox("Settings have been Saved", MsgBoxStyle.OkOnly, "Verification")
        End If

    End Sub

    Sub Save_Settings_ctls( _
    ByRef cc As Control, _
    ByRef FORM_NAME As String, _
    ByRef SET_ID As String, _
    ByRef XNO As String, _
    ByRef tblASTROPT2 As DataTable)

        Dim rowASTROPT2 As DataRow
        For Each ctl As Control In cc.Controls
            If ctl.Controls.Count > 0 Then
                Call Save_Settings_ctls(ctl, FORM_NAME, SET_ID, XNO, tblASTROPT2)
            End If
            Dim ABSCOLUMN_NAME As String = Absx1.GetABSColumnName(ctl)
            If ABSCOLUMN_NAME <> "" Then
                rowASTROPT2 = tblASTROPT2.NewRow()
                With rowASTROPT2
                    .Item("FORM_NAME") = FORM_NAME
                    .Item("SET_ID") = SET_ID
                    .Item("SET_CTL_NAME") = ctl.Name
                    .Item("SET_CTL_TYPE") = ctl.GetType.Name
                    Select Case ctl.GetType.Name
                        Case "UltraCheckEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraCheckEditor).Checked
                        Case "UltraOptionSet"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraOptionSet).Value
                        Case "ABSCheckBox"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, ABSCS.ABSCheckBox).ABSChecked
                        Case "UltraTrackBar"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraTrackBar).Value
                        Case "UltraCombo"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinGrid.UltraCombo).Value
                        Case "UltraComboEditor"
                            .Item("SET_CTL_DATA") = DirectCast(ctl, Infragistics.Win.UltraWinEditors.UltraComboEditor).Value
                        Case Else
                            .Item("SET_CTL_DATA") = ctl.Text
                    End Select
                    .Item("SET_CTL_TAG") = ABSCOLUMN_NAME
                    .Item("XNO") = XNO
                End With

                tblASTROPT2.Rows.Add(rowASTROPT2)
            End If
        Next
    End Sub

    Sub Show_Settings()
        ASCMAIN1.sql = "Select * from ASTROPT1 where FORM_NAME = '" & FORM_NAME & "'"
        tblASTROPT1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTROPT1", 2)
        grdASTROPT1.DataSource = tblASTROPT1
        Sort_grdColumns(grdASTROPT1, "LAST_DATE".ToLower)
    End Sub
#End Region

#Region "Supporting Routines"

    Function SQL_in(ByVal COLUMN_NAME As String, _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        Dim CODE_VALUES = SQLA(COLUMN_NAME, "CODE_VALUES", True)
        Dim sql As String = ""

        If CODE_VALUES <> "" Then
            Dim single_code_value As Boolean = (InStr(CODE_VALUES, "','") = 0)

            If single_code_value Then
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " <> ", " = ") & CODE_VALUES
            Else
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " in (" & CODE_VALUES & ")"
            End If

            If DB_COLUMN_NAME <> "" Then
                sql = " and " & DB_COLUMN_NAME & sql
            Else
                sql = " and " & COLUMN_NAME & sql
            End If
        End If

        Return sql
    End Function

    Function SQLA( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal COLUMN_NAME As String = "CODE_VALUES", _
    Optional ByVal SQL_List As Boolean = False) As String
        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        If rowASTDSQLA Is Nothing Then
            SQLA = ""
        Else
            SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
            If SQL_List And SQLA <> "" Then
                SQLA = "'" & Replace(SQLA, ",", "','") & "'"
            End If
        End If
        Return SQLA
    End Function

    Function SQLA_filter( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal DB_TABLE_NAME As String = "", _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        If DB_TABLE_NAME <> "" Then
            If DB_COLUMN_NAME = "" Then
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & PB_COLUMN_NAME
            Else
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & DB_COLUMN_NAME
            End If
        End If

        If DB_COLUMN_NAME = "" Then
            DB_COLUMN_NAME = PB_COLUMN_NAME
        End If

        Dim z As String
        z = SQLA(PB_COLUMN_NAME, "CODE_VALUES", True)
        If z <> "" Then
            SQLA_filter = " AND " & DB_COLUMN_NAME & IIf(SQLA(PB_COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " IN (" & z & ")" & vbCr
        Else
            SQLA_filter = ""
        End If
        Return SQLA_filter
    End Function

    Function Get_Filter(ByVal COLUMN_NAME As String, ByVal SQL_ELEMENT_TO_COMPARE_TO As String) As String
        Dim sqlw As String = ""
        If SQLA(COLUMN_NAME, "CODE_VALUES") <> "" Then
            sqlw = " and " & SQL_ELEMENT_TO_COMPARE_TO & " " & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", "Not ", "") & "in (" & SQLA(COLUMN_NAME, "CODE_VALUES", True) & ")"
        End If
        Return sqlw

    End Function

    ''' <summary>
    ''' Used to pull data from Oracle into a DataTable based on the Code Values in use in the DataTableName_with_CodeValues specified
    ''' </summary>
    ''' <param name="DataTableName_with_CodeValues"></param>
    ''' <param name="COLUMN_NAMEs"></param>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="SelectList"></param>
    ''' <remarks></remarks>
    Sub Get_WKCodes( _
    ByVal DataTableName_with_CodeValues As String, _
    ByVal COLUMN_NAMEs As String, _
    ByVal TABLE_NAME As String, _
    Optional ByVal SelectList As String = "*")

        Dim COLUMNS() As String = Split(COLUMN_NAMEs, ",")
        Dim Number_of_Key_Fields As Integer = COLUMNS.Length

        ' Get Distinct Values from DataTable

        Dim SQLX As String = ""
        For Each row As DataRow In _
        ASCMAIN1.Distinct_Values("", dst.Tables(DataTableName_with_CodeValues), COLUMNS).Rows ' ASCDATA1.SelectDistinct(dst.Tables(DataTableName_with_CodeValues), COLUMNS).Rows
            Dim z As String = ""
            For i As Integer = 0 To Number_of_Key_Fields - 1
                z = z & ",'" & row.Item(i) & "'"
            Next i
            If Number_of_Key_Fields = 1 Then
                SQLX = SQLX & z
            Else
                SQLX = SQLX & ",(" & Mid$(z, 2) & ")"
            End If
        Next

        ' SELECT * FROM ARTCUST2 WHERE (CUST_CODE,CUST_SHIP_TO_NO) IN 
        ' (('025000','010115'),('025000','010120'),('025000','010125'))

        If SQLX <> "" Then
            SQLX = " where (" & COLUMN_NAMEs & ") in (" & Mid$(SQLX, 2) & ")"
        Else
            SQLX = " where ROWNUM < 1"
        End If

        If Not dst.Tables.Contains(TABLE_NAME) Then
            ASCMAIN1.sql = "Select " & SelectList & " from " & TABLE_NAME & SQLX
            dst.Tables.Add(ASCDATA1.GetDataTable("", TABLE_NAME, Number_of_Key_Fields))
        End If
    End Sub

#End Region

#Region "GL Routines"

    Sub GL_Rounding(ByVal DETL_CTL_DATE As Date)

        'Stop ' watch 1st on thru

        For Each row As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("GLTINTF1"), "JOURNAL_TYPE", "JOURNAL_NO", "OPS_YYYYPP").Rows 'ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1"), New String() {"JOURNAL_TYPE", "JOURNAL_NO", "OPS_YYYYPP"}).Rows
            Dim sqlwx As String = "JOURNAL_TYPE = '" & row.Item("JOURNAL_TYPE") & "'" _
                            & " and JOURNAL_NO = '" & row.Item("JOURNAL_NO") & "'" _
                            & " and OPS_YYYYPP = '" & row.Item("OPS_YYYYPP") & "'"
            'Stop ' ROUND
            Dim OOBAL_ROUND_AMT As Decimal = Val(ASCDATA1.GetDataValue("Select JOURNAL_OOBAL_ROUND_AMT from GLTTYPE1 where JOURNAL_TYPE = '" & row.Item("JOURNAL_TYPE") & "'") & "")
            If OOBAL_ROUND_AMT = 0 Then
                OOBAL_ROUND_AMT = Val(ROWs("GLTPARM1").Item("GL_PARM_OOBAL_ROUND_AMT") & "")
            End If
            Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("GLTINTF1").Compute("SUM (DETL_POSTING_AMT)", sqlwx) & "")
            Dim JOURNAL_LNO As Integer = Val(dst.Tables("GLTINTF1").Compute("MAX (JOURNAL_LNO)", sqlwx) & "")

            If DETL_POSTING_AMT <> 0 And Abs(DETL_POSTING_AMT) <= OOBAL_ROUND_AMT Then
                Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                rowGLTINTF1.Item("JOURNAL_TYPE") = row.Item("JOURNAL_TYPE")
                rowGLTINTF1.Item("JOURNAL_NO") = row.Item("JOURNAL_NO")
                rowGLTINTF1.Item("JOURNAL_LNO") = JOURNAL_LNO + 1
                rowGLTINTF1.Item("OPS_YYYYPP") = row.Item("OPS_YYYYPP")
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("GLTPARM1").Item("GL_PARM_ACCT_ROUNDING")
                rowGLTINTF1.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowGLTINTF1.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowGLTINTF1.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                rowGLTINTF1.Item("DETL_POSTING_AMT") = -1 * DETL_POSTING_AMT
                rowGLTINTF1.Item("DETL_CTL_DATE") = DETL_CTL_DATE
                rowGLTINTF1.Item("DETL_EXE_NO") = XNO
                dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
            End If
        Next
    End Sub

    Sub GL_Update()

        Update_Record_TDA("GLTJRNL1")

        Create_TDA(dst.Tables.Add, "GLTDETL1", "*")

        For Each row As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("GLTINTF1"), "OPS_YYYYPP", "JOURNAL_NO").Rows
            Dim JOURNAL_NO As String = row.Item("JOURNAL_NO")
            Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP")
            dst.Tables("GLTDETL1").Rows.Clear()
            Dim sqlx As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and JOURNAL_NO = '" & JOURNAL_NO & "'"
            For Each rowGLTINTF1 As DataRow In dst.Tables("GLTINTF1").Select(sqlx)
                Dim rowGLTDETL1 As DataRow = dst.Tables("GLTDETL1").NewRow
                For i As Integer = 0 To rowGLTINTF1.ItemArray.Length - 1
                    Dim COLUMN_NAME As String = dst.Tables("GLTINTF1").Columns(i).ColumnName
                    If COLUMN_NAME = "DIST_CODE" Or COLUMN_NAME = "JOURNAL_TYPE" Then
                    Else
                        rowGLTDETL1.Item(COLUMN_NAME) = rowGLTINTF1.Item(COLUMN_NAME)
                    End If
                Next
                dst.Tables("GLTDETL1").Rows.Add(rowGLTDETL1)
            Next
            InterCompany(JOURNAL_NO, OPS_YYYYPP)
            Update_Record_TDA("GLTDETL1")
            Update_GLTACCT3(JOURNAL_NO, OPS_YYYYPP)
        Next

    End Sub

    Sub Prepare_Journal()

        If Not dst.Tables.Contains("GLTTYPE1") Then
            dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTTYPE1"))
        End If

        If Not dst.Tables.Contains("GLTSEGM1") Then
            dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTSEGM1"))
        End If

        If ASCMAIN1.DBS_COMPANY = "TFP" Then
            For Each row As DataRow In dst.Tables("GLTINTF1").Rows
                If row.Item("SEG2_CODE") = "E" _
                Or row.Item("SEG2_CODE") = "W" _
                Or row.Item("SEG2_CODE") = "M" _
                Or row.Item("SEG2_CODE") = "F" Then
                    row.Item("SEG2_CODE") = "A"
                End If
            Next
        End If

        Prepare_GL_Account_Activity_Recaps("GLTINTF1")

        Create_TDA(dst.Tables.Add, "GLTJRNL1", "*")

        Dim JOURNAL_NO As String

        For Each row As DataRow In _
        ASCMAIN1.Distinct_Values("", dst.Tables("GLTINTF1"), _
        "JOURNAL_NO", "JOURNAL_TYPE").Rows
            JOURNAL_NO = row("JOURNAL_NO")
            Dim rowGLTJRNL1 As DataRow
            'rowGLTJRNL1 = dst.Tables("GLTJRNL1").Rows.Find(JOURNAL_NO)
            rowGLTJRNL1 = dst.Tables("GLTJRNL1").NewRow
            rowGLTJRNL1("JOURNAL_NO") = JOURNAL_NO
            rowGLTJRNL1("JOURNAL_DESC") = dst.Tables("GLTTYPE1").Rows.Find(row("JOURNAL_TYPE")).Item("JOURNAL_TYPE_DESC")
            rowGLTJRNL1("JOURNAL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
            rowGLTJRNL1("OPS_YYYYPP") = ASCMAIN1.CYP
            rowGLTJRNL1("JOURNAL_TYPE") = row("JOURNAL_TYPE")
            rowGLTJRNL1("INIT_OPER") = ASCMAIN1.USER_ID
            rowGLTJRNL1("INIT_DATE") = DATETIME_STAMP
            'rowGLTJRNL1("LAST_OPER") = ASCMAIN1.USER_ID
            'rowGLTJRNL1("LAST_DATE") = DATETIME_STAMP
            rowGLTJRNL1("SEG2_CODE") = ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
            rowGLTJRNL1("SEG3_CODE") = ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
            rowGLTJRNL1("SEG4_CODE") = ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
            rowGLTJRNL1("REGISTER_XNO") = XNO
            dst.Tables("GLTJRNL1").Rows.Add(rowGLTJRNL1)
        Next

        GL_Rounding(DATETIME_STAMP.Date)

        For i As Integer = 2 To 4
            Dim z As String = "SEG" & Format$(i, "0") & "_CODE"
            For Each row As DataRow In dst.Tables("GLTINTF1").Select(z & " is Null or " & z & " = ''")
                row.Item(z) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & Format$(i, "0"))
            Next
        Next

        TDAs("GLTINTF1").Update(dst.Tables("GLTINTF1"))
    End Sub

    Sub Print_GL(Optional ByVal BY_DR_CR_IND As Boolean = False)

        Prepare_Journal()

        Get_WKCodes("GLTINTF1", "OPS_YYYYPP", "GLTPARM2")
        Get_WKCodes("GLTINTF1", "ACCT_CODE", "GLTACCT1")

        CR_params.Add("GL_PARM_SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("GL_PARM_SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("GL_PARM_SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
        CR_params.Add("GL_PARM_ACCT_RECAPS", ROWs("GLTPARM1").Item("GL_PARM_ACCT_RECAPS") & "")
        If BY_DR_CR_IND Then
            CR_params.Add("BY_DR_CR_IND", "Y")
        End If
        ' CR_params.Add("BY_DR_CR_IND", IIf(BY_DR_CR_IND, "Y", "N"))

        F.Generate_Report("GLRINTF1")
        If Not Validate_JE() Then
            If RWU = "R" Then
                RWU = "N"
            End If
        End If
    End Sub

    Function Validate_JE() As Boolean

        Dim tblGLTERROR As New DataTable("GLTERROR")
        With tblGLTERROR
            .Columns.Add("OPS_YYYYPP", GetType(System.String))
            .Columns.Add("JOURNAL_NO", GetType(System.String))
            .Columns.Add("ACCT_CODE", GetType(System.String))
            .Columns.Add("ACCT_SEG_ID", GetType(System.String))
            .Columns.Add("ACCT_SEG_CODE", GetType(System.String))
            .Columns.Add("ERROR_TEXT", GetType(System.String))
        End With

        dst.Tables.Add(tblGLTERROR)

        Dim ERROR_TEXT As String = ""

        Dim ACCT_SEG_ID As String = ""
        Dim ACCT_SEG_CODE As String = ""
        Dim ACCT_CODE As String = ""
        Dim JOURNAL_TYPEs As New List(Of String)

        If RWU = "R" Then
            For Each row As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("GLTINTF1"), "JOURNAL_NO").Rows
                Dim JOURNAL_NO As String = row.Item("JOURNAL_NO")
                Dim OPS_YYYYPP As String = dst.Tables("GLTINTF1").Compute("MIN(OPS_YYYYPP)", "JOURNAL_NO = '" & JOURNAL_NO & "'")
                If OPS_YYYYPP < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                    ERROR_TEXT = "Some Entries are Attempting to Post to a Closed GL Period"
                    Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, "", "", "", ERROR_TEXT)
                Else
                    Dim GL_PARM_SOFT_CLOSE_YYYYPP As String = ROWs("GLTPARM1").Item("GL_PARM_SOFT_CLOSE_YYYYPP")
                    If OPS_YYYYPP <= ROWs("GLTPARM1").Item("GL_PARM_SOFT_CLOSE_YYYYPP") Then
                        ERROR_TEXT = $"Some Entries are Attempting to Post to a Soft-Closed GL Period ({GL_PARM_SOFT_CLOSE_YYYYPP})"
                        Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, "", "", "", ERROR_TEXT)
                    End If
                End If
            Next
        End If

        For Each row As DataRow In ASCMAIN1.Distinct_Values _
            ("", dst.Tables("GLTINTF1"), "OPS_YYYYPP", "JOURNAL_NO", "JOURNAL_TYPE").Rows
            ' ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1"), New String() {"OPS_YYYYPP", "JOURNAL_NO"}).Rows
            Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP")
            Dim JOURNAL_NO As String = row.Item("JOURNAL_NO")
            Dim JOURNAL_TYPE As String = row.Item("JOURNAL_TYPE")

            If Not JOURNAL_TYPEs.Contains(JOURNAL_TYPE) Then
                JOURNAL_TYPEs.Add(JOURNAL_TYPE)

                Dim rowGLTTYPE1 As DataRow = LookUp("GLTTYPE1", JOURNAL_TYPE)
                If rowGLTTYPE1 Is Nothing Then
                    ERROR_TEXT = "Journal Type " & JOURNAL_TYPE & " is not on File"
                    Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, "", "", "", ERROR_TEXT)
                Else
                    If rowGLTTYPE1.Item("JOURNAL_STATUS") & "" <> "A" Then
                        ERROR_TEXT = "Journal Type " & JOURNAL_TYPE & " is not Active"
                        Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, "", "", "", ERROR_TEXT)
                    End If
                End If
            End If

            ACCT_SEG_ID = ""
            ACCT_SEG_CODE = ""
            ACCT_CODE = ""
            Dim sqlx As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and JOURNAL_NO = '" & JOURNAL_NO & "'"
            Dim T As Decimal = dst.Tables("GLTINTF1").Compute("SUM (DETL_POSTING_AMT)", sqlx)
            If Abs(T) >= 0.01 Then
                ERROR_TEXT = "Journal is Out of Balance by " & Format(T, "##,##0.00")
                Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
            End If

            For Each rowACCT_CODE As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("GLTINTF1"), "ACCT_CODE").Rows ' ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1").Select(sqlx), New String() {"ACCT_CODE"}).Rows
                ACCT_CODE = rowACCT_CODE.Item("ACCT_CODE") & ""
                ACCT_SEG_ID = ""
                ACCT_SEG_CODE = ""
                Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                If rowGLTACCT1 Is Nothing Then
                    ERROR_TEXT = "Invalid Account Code"
                    Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                Else
                    If rowGLTACCT1.Item("ACCT_STATUS") & "" <> "A" Then
                        ERROR_TEXT = "Account Status is not Active"
                        Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                    Else
                        ' Dim sqly As String = sqlx & " and ACCT_CODE = '" & ACCT_CODE & "'"
                        For i As Integer = 2 To 4
                            ACCT_SEG_ID = CStr(i)
                            Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                            Dim ACCT_SEG_DEFAULT As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                            For Each rowACCT_SEG_CODE As DataRow In ASCMAIN1.Distinct_Values("", "ACCT_CODE = '" & ACCT_CODE & "'", dst.Tables("GLTINTF1"), COLUMN_NAME).Rows
                                ACCT_SEG_CODE = rowACCT_SEG_CODE.Item(COLUMN_NAME) & ""
                                Dim ACCT_SEG_TYPE As String = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                                If ACCT_SEG_TYPE = "" Then
                                    ACCT_SEG_TYPE = "Segment " & ACCT_SEG_ID
                                End If
                                cdr = LookUp("GLTSEGM1", New String() {ACCT_SEG_ID, ACCT_SEG_CODE})
                                If cdr Is Nothing Then
                                    ERROR_TEXT = "Invalid " & ACCT_SEG_TYPE & " Code"
                                    Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                                Else
                                    If cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                                        ERROR_TEXT = ACCT_SEG_TYPE & " Code not Permitted for J/E"
                                        Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                                    End If
                                    If cdr.Item("ACCT_SEG_STATUS") & "" <> "A" Then
                                        ERROR_TEXT = ACCT_SEG_TYPE & " Code not Active"
                                        Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                                    End If

                                    If ROWs("GLTPARM1").Item("GL_PARM_MAND_SEG_CTL") & "" = "1" Then
                                        ' 0 or D = Requires Default
                                        ' 1 or N = Requires Non-Default
                                        ' 2 OR A = Any Value
                                        ' ELIMINATING OLD 0/1/2 HAVING CHANGED GLTACCT1 AND ASTCODE1, AND TO MAKE WAY FOR NEW VALUES

                                        If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "D" Then
                                            If ACCT_SEG_CODE <> ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                                ERROR_TEXT = "Acct " & ACCT_CODE & " requires Default Value (" & ACCT_SEG_DEFAULT & ") for " & ACCT_SEG_TYPE
                                                Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                                            End If
                                        End If
                                        If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "N" Then
                                            If ACCT_SEG_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                                ERROR_TEXT = "Acct " & ACCT_CODE & " requires non-Default Value (" & ACCT_SEG_DEFAULT & ") for " & ACCT_SEG_TYPE
                                                Call Validate_JE_Error(OPS_YYYYPP, JOURNAL_NO, ACCT_CODE, ACCT_SEG_ID, ACCT_SEG_CODE, ERROR_TEXT)
                                            End If
                                        End If
                                    End If

                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Next

        If dst.Tables("GLTERROR").Rows.Count > 0 Then
            F.Generate_Report("GLRERROR", , "Issues with GL Distribution")
        End If

        xErrMsg = ERROR_TEXT

        Return (dst.Tables("GLTERROR").Rows.Count = 0)

    End Function

    Sub Validate_JE_Error( _
    ByVal OPS_YYYYPP As String, _
    ByVal JOURNAL_NO As String, _
    ByVal ACCT_CODE As String, _
    ByVal ACCT_SEG_ID As String, _
    ByVal ACCT_SEG_CODE As String, _
    ByVal ERROR_TEXT As String)

        Dim rowGLTERROR As DataRow = dst.Tables("GLTERROR").NewRow
        rowGLTERROR.Item("OPS_YYYYPP") = OPS_YYYYPP
        rowGLTERROR.Item("JOURNAL_NO") = JOURNAL_NO
        rowGLTERROR.Item("ACCT_CODE") = ACCT_CODE
        rowGLTERROR.Item("ACCT_SEG_ID") = ACCT_SEG_ID
        rowGLTERROR.Item("ACCT_SEG_CODE") = ACCT_SEG_CODE
        rowGLTERROR.Item("ERROR_TEXT") = ERROR_TEXT
        dst.Tables("GLTERROR").Rows.Add(rowGLTERROR)

    End Sub

    Sub Special_Routines_for_ACCT_TYPE()

        ' Special Sequence for ACCT_TYPE

        Dim j As Integer = Val(SQLA("ACCT_TYPE", "SEQUENCE"))
        If j <> 0 Then
            Dim zx As String = "Decode (G" & CStr(j) & ",'A','1','L','2','E','3','I','4','X','5','0')"
            sql = "Update " & ASTSRPT1 & " Set G" & CStr(j) & " = " & zx & " || G" & CStr(j)
            ASCDATA1.ExecuteSQL(sql)
        End If

        ' Write Group Records

        Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":1A", "A", "Asset")
        Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":2L", "L", "Liability")
        Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":3E", "E", "Equity")
        Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":4I", "I", "Income")
        Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":5X", "X", "Expense")
    End Sub

#End Region

    Sub Check_if_Empty(ByVal TABLE_NAME As String)

        dst.Tables(TABLE_NAME).AcceptChanges()
        If dst.Tables(TABLE_NAME).Rows.Count = 0 Then
            RWU &= "0"
            xErrMsg = "No Eligible Records"
        End If
    End Sub

#Region "Data Query"

    Sub Initialize_Data_Query()

        Format_grdASTDSQL1_pre()

        ' Setup DQ Column Sequencing Control

        Dim dt As New DataTable
        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            dt.Columns.Add(COLUMN_NAMEs(i))
            dt.Columns(i).Caption = COLUMN_CAPTIONs(i)
        Next
        Dim row As DataRow = dt.NewRow
        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            row.Item(i) = i + 1
        Next
        dt.Rows.Add(row)
        grdDQseq.DataSource = dt
        With grdDQseq.DisplayLayout.Bands(0)
            .CardView = True
            .CardSettings.LabelWidth = 100
            .CardSettings.ShowCaption = False
            .CardSettings.Width = 1
        End With


        ' Setup Treeview Column Structure

        tvwDQ.Nodes.Clear()
        With tvwDQ
            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
            rootColumnSet.Columns.Clear()
            For i As Integer = 1 To COLUMN_NAMEs.Count
                Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add("G" & CStr(i))
            Next
        End With


        ' Show CheckBoxes, OptionSet and Additional GridColumn if COLUMN_NAME_appended <> ""
        'TODO: Need to work on display member vs value member for these appended columns, and sort order
        For Each ctl As Control In grpRECAPS.Controls
            Try
                Dim CHK As ABSCS.ABSCheckBox = DirectCast(ctl, ABSCS.ABSCheckBox)
                CHK = Nothing
            Catch ex As Exception

            End Try
        Next

        Dim RR As Integer = tblASTRECAP.Rows.Count
        If RR = 0 Then
            optRECAPSORT.ValueList = Nothing
        Else
            ReDim DATA_TYPEs(RR)
            Dim i As Integer = 0
            Dim VL As New ValueList
            Dim t As Integer = 0
            For Each rowASTRECAP As DataRow In tblASTRECAP.Rows
                VL.ValueListItems.Add(rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO"), rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CAPTION") & "")
                Dim chk As New ABSCS.ABSCheckBox
                chk.Checked = True
                chk.Text = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CAPTION") & ""
                chk.Tag = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & ""
                chk.Parent = grpRECAPS
                chk.Top = t + 5
                chk.Left = optRECAPSORT.Left
                t += chk.Height
                chk.Visible = True
                i += 1
                DATA_TYPEs(i) = chk.Text
                AddHandler chk.CheckedChanged, AddressOf chkRECAP_CheckedChanged
            Next
            grpRECAPS.Tag = "".PadLeft(i, "1")
            optRECAPSORT.ValueList = VL
            optRECAPSORT.Tag = "X"
            optRECAPSORT.CheckedIndex = 0
            optRECAPSORT.Tag = ""
        End If
        grpRECAPS.Visible = (RR <> 0)
        grpRECAPSORT.Visible = (RR <> 0)
        tabRECAP.Tabs("Sort").Visible = (RR <> 0)
        tabRECAP.Tabs("Recaps").Visible = (RR <> 0)


        ' Set up Main DQ Grid

        tblASTDSQL1 = New DataTable("ASTDSQL1")
        If dst.Tables.Contains(tblASTDSQL1.TableName) Then
            dst.Tables.Remove(tblASTDSQL1)
        End If
        dst.Tables.Add(tblASTDSQL1)
        tblASTDSQL1.Columns.Add("CODE_VALUE")
        tblASTDSQL1.Columns.Add("DESC_VALUE")
        tblASTDSQL1.Columns.Add("SORT_VALUE")
        tblASTDSQL1.Columns.Add("RANK_VALUE")

        grdASTDSQL1.DataSource = Nothing

        For COL_sfx As Integer = Sign(RR) To RR
            For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
                Dim dtype As String = COLUMN_NAME_sum(COLUMN_NAME)
                Dim COLUMN_NAME_sfx As String = COLUMN_NAME
                If COL_sfx <> 0 Then
                    COLUMN_NAME_sfx &= "_" & CStr(COL_sfx)
                End If
                Select Case dtype
                    Case "DEC"
                        tblASTDSQL1.Columns.Add(COLUMN_NAME_sfx, GetType(System.Decimal))
                    Case "AMT"
                        tblASTDSQL1.Columns.Add(COLUMN_NAME_sfx, GetType(System.Double))
                    Case "QTY"
                        tblASTDSQL1.Columns.Add(COLUMN_NAME_sfx, GetType(System.Int64))
                End Select
            Next
        Next

        grdASTDSQL1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdASTDSQL1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdASTDSQL1.DataSource = tblASTDSQL1

        With grdASTDSQL1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"CODE_VALUE", "DESC_VALUE", "RANK_VALUE"}
                .Columns(COLUMN_NAME).Group = .Groups.Add(COLUMN_NAME)
                .Groups(COLUMN_NAME).Header.Appearance.BackColor = Color.Yellow
                .Groups(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.Yellow
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            Next

            If RR <> 0 Then
                .LevelCount = RR

                .Groups.Add("DATA_TYPE", "Data Type")
                .Groups("DATA_TYPE").Header.Appearance.BackColor = Color.Yellow
                .Groups("DATA_TYPE").Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                For i As Integer = 1 To RR
                    COLUMN_NAME = "DATA_TYPE_" & CStr(i)
                    .Columns.Add(COLUMN_NAME, DATA_TYPEs(i))
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.Yellow
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Group = .Groups("DATA_TYPE")
                    .Columns(COLUMN_NAME).Level = i - 1
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                Next
            End If



            For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
                .Groups.Add(COLUMN_NAME)
                For COL_sfx As Integer = Sign(RR) To RR
                    Dim COLUMN_NAME_sfx As String = COLUMN_NAME
                    If COL_sfx <> 0 Then
                        COLUMN_NAME_sfx &= "_" & CStr(COL_sfx)
                    End If
                    .Columns(COLUMN_NAME_sfx).Group = .Groups(COLUMN_NAME)
                    If COL_sfx <> 0 Then
                        .Columns(COLUMN_NAME_sfx).Level = COL_sfx - 1
                    End If
                Next
            Next
        End With



        tblASTDSQL1_copy = tblASTDSQL1.Clone

        With grdASTDSQL1.DisplayLayout.Bands(0)
            .Groups("CODE_VALUE").Header.Fixed = True
            .Groups("DESC_VALUE").Header.Fixed = True
            .Groups("RANK_VALUE").Header.Fixed = True
            If RR <> 0 Then
                .Groups("DATA_TYPE").Header.Fixed = True
            End If
        End With

        With grdASTDSQL1.DisplayLayout.Bands(0)
            .Summaries.Clear()
        End With

        'Create_Summary(grdASTDSQL1, "CODE_VALUE", "Count")
        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim COLUMN_NAME As String = rowASTDSQLS.Item("COLUMN_NAME")
            For COL_sfx As Integer = Sign(RR) To RR
                Dim COLUMN_NAME_sfx As String = COLUMN_NAME
                If COL_sfx <> 0 Then
                    COLUMN_NAME_sfx &= "_" & CStr(COL_sfx)
                End If

                With grdASTDSQL1.DisplayLayout.Bands(0).Columns(COLUMN_NAME_sfx)
                    .Width = rowASTDSQLS.Item("COLUMN_WIDTH")
                    If rowASTDSQLS.Item("COLUMN_FORMAT") & "" <> "" Then
                        .Format = rowASTDSQLS.Item("COLUMN_FORMAT")
                    End If

                    .Header.Caption = rowASTDSQLS.Item("COLUMN_CAPTION")
                    .Header.Appearance.BackColor = Color.LightGreen '  .LightSteelBlue
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .Header.Appearance.TextHAlign = HAlign.Right

                End With
                'Create_Summary(grdASTDSQL1, COLUMN_NAME_sfx)
            Next
            With grdASTDSQL1.DisplayLayout.Bands(0).Groups(COLUMN_NAME)
                .Header.Caption = rowASTDSQLS.Item("COLUMN_CAPTION")
                .Header.Appearance.BackColor = Color.LightGreen '  .LightSteelBlue
                .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
        Next
        grdASTDSQL1.DisplayLayout.Bands(0).ColHeadersVisible = False
        grdASTDSQL1.DisplayLayout.Bands(0).Groups("RANK_VALUE").Hidden = True


        grdASTDSQL1.DisplayLayout.Override.FixedRowStyle = UltraWinGrid.FixedRowStyle.Bottom
        grdASTDSQL1.DisplayLayout.Override.FixedRowIndicator = UltraWinGrid.FixedRowIndicator.Button


        ' Set up Top N by Column

        Dim dtTOP As New DataTable
        dtTOP.Columns.Add("COLUMN_NAME")
        dtTOP.Columns.Add("COLUMN_CAPTION")
        COLUMN_NAME_sum_first = ""
        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim COLUMN_NAME As String = rowASTDSQLS.Item("COLUMN_NAME")
            Dim COLUMN_CAPTION As String = rowASTDSQLS.Item("COLUMN_CAPTION")
            If COLUMN_NAME_sum_first = "" Then
                COLUMN_NAME_sum_first = COLUMN_NAME
            End If
            dtTOP.Rows.Add(New String() {COLUMN_NAME, COLUMN_CAPTION})
        Next

        cbeTopN.DataSource = dtTOP
        cbeTopN.ValueMember = "COLUMN_NAME"
        cbeTopN.DisplayMember = "COLUMN_CAPTION"
        cbeTopN.Value = COLUMN_NAME_sum_first

        Format_grdASTDSQL1_post(grdASTDSQL1)

        Generate_Inquiry()

    End Sub

    Overridable Sub Format_grdASTDSQL1_pre()

    End Sub

    Overridable Sub Format_grdASTDSQL1_post(ByVal grdASTDSQL1 As UltraWinGrid.UltraGrid)

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = (UltraTabControl1.SelectedTab.Key = "Data Query") Or (UltraTabControl1.SelectedTab.Key = "Data Exports")
        UltraExplorerBar1.Groups("Data Query").Visible = ASCMAIN1.USER_ID = "wjzz" And (UltraTabControl1.SelectedTab.Key = "Data Query")
        UltraExplorerBar1.Groups("Data Grid").Visible = ASCMAIN1.USER_ID = "wjzz" And (UltraTabControl1.SelectedTab.Key = "Data Grid")
    End Sub

    Private Sub tvw_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwDQ.Click
        Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
        Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
        Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

        Click_Node(tnode)
    End Sub

    Sub Click_Node(ByVal tnode As UltraWinTree.UltraTreeNode)
        If tnode IsNot Nothing Then
            Dim L As Integer = tnode.Level
            tblASTDSQL1.Rows.Clear()

            With grdASTDSQL1.DisplayLayout.Bands(0)
                .Groups("CODE_VALUE").Header.Caption = COLUMN_CAPTION_by_Lvl(L + 1)
                .Groups("DESC_VALUE").Header.Caption = "Description"
                .Columns("CODE_VALUE").Header.Caption = COLUMN_CAPTION_by_Lvl(L + 1)
                .Columns("DESC_VALUE").Header.Caption = "Description"
            End With

            Dim sql As String = ""
            If L < G_by_Lvl.Length - 1 Then
                For iLVL As Integer = L + 1 To G_by_Lvl.Length - 1
                    sql &= " and G" & G_by_Lvl(iLVL) & " <> '" & aRC & "'"
                Next
                sql = Mid(sql, 6)
            End If

            If L > 0 Then
                Dim caption As String = ""
                For Lvl As Integer = 1 To L
                    sql &= " and G" & CStr(G_by_Lvl(Lvl)) & " = '" & tnode.Cells(Lvl - 1).Text & "'"
                    caption &= ", " & COLUMN_CAPTION_by_Lvl(Lvl) & " " & Split(tnode.Cells(Lvl - 1).Text, ":")(1)
                Next
                'sql = Mid(sql, 5)
                grdASTDSQL1.Text = Mid(caption, 2)
            Else
                grdASTDSQL1.Text = tnode.Text
            End If

            If COLUMN_NAME_sum_first = "" Then
                Exit Sub
            Else
                'If ASCMAIN1.USER_ID = "wjz" Then Stop ' WHEN IS THIS FIELD <> ""? - WHEN THE FIELDS TO TOTAL ARE DEFINED IN REPORT MAINT - SEE ARRATBR1
            End If

            Dim c0 As Integer = dst.Tables("ASTSRPT1").Columns(COLUMN_NAME_sum_first).Ordinal
            Dim RR As Integer = tblASTRECAP.Rows.Count

            Dim orderby As String = "G" & G_by_Lvl(L + 1)
            Dim CODE_VALUE As String = ""
            ReDim Totals_All(COLUMN_NAME_sum.Count * IIf(RR = 0, 1, RR))

            Dim RECAP_ROW_INDEX As New Dictionary(Of Integer, Integer)
            Dim RI As Integer = 0
            If RR <> 0 Then
                For Each rowASTRECAP As DataRow In tblASTRECAP.Select("", "ASTSRPT1_RECAP_ROW_NO", DataViewRowState.CurrentRows)
                    RI += 1
                    RECAP_ROW_INDEX.Add(Val(rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & ""), RI)
                Next
            End If

            For Each row As DataRow In dst.Tables("ASTSRPT1").Select(sql, orderby)
                If row.Item(orderby) & "" <> CODE_VALUE Then
                    If CODE_VALUE <> "" Then Add_DQ_Row(CODE_VALUE)
                    CODE_VALUE = row.Item(orderby) & ""
                    ReDim Totals_Row(COLUMN_NAME_sum.Count * IIf(RR = 0, 1, RR))
                End If
                Dim ASTSRPT1_RECAP_ROW_NO As Integer = 1
                If RR <> 0 Then
                    ASTSRPT1_RECAP_ROW_NO = RECAP_ROW_INDEX(Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & ""))
                End If
                For j As Integer = 1 To COLUMN_NAME_sum.Count
                    Dim k As Integer = j + c0 - 1
                    Totals_Row(j + (ASTSRPT1_RECAP_ROW_NO - 1) * COLUMN_NAME_sum.Count) += Val(row.Item(k) & "")
                Next
            Next
            If CODE_VALUE <> "" Then Add_DQ_Row(CODE_VALUE)

            tblASTDSQL1_copy.Rows.Clear()
            tblASTDSQL1_copy.Merge(tblASTDSQL1)

            Add_Totals_Row()

            If chkTopN.Checked Then
                Show_TopN()
                'grdASTDSQL1.Text &= " - Top " & CStr(numTopN.Value) & " Based on " & cbeTopN.Text
            Else
                'Sort_grdColumns (grdASTDSQL1 ,
            End If
        End If
    End Sub

    Sub Add_DQ_Row(ByVal CODE_VALUE As String)

        Dim rowASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(CODE_VALUE)
        Dim rowDQ As DataRow = tblASTDSQL1.NewRow
        If rowASTGROUP IsNot Nothing Then
            rowDQ.Item("CODE_VALUE") = rowASTGROUP.Item("GROUP_CODE")
            rowDQ.Item("DESC_VALUE") = rowASTGROUP.Item("GROUP_DESC")
            rowDQ.Item("SORT_VALUE") = 1
            For j As Integer = 1 To UBound(Totals_Row)
                rowDQ.Item(3 + j) = Totals_Row(j)
                Totals_All(j) += Totals_Row(j)
            Next
            Totals_All(0) += 1
            tblASTDSQL1.Rows.Add(rowDQ)
        End If
    End Sub

    Sub Add_Totals_Row(Optional ByVal rowOthers As DataRow = Nothing)

        Dim rowDQ As DataRow = tblASTDSQL1.NewRow
        rowDQ.Item("CODE_VALUE") = "Totals"
        rowDQ.Item("DESC_VALUE") = CStr(Totals_All(0))
        rowDQ.Item("SORT_VALUE") = 3
        For j As Integer = 1 To UBound(Totals_All)
            rowDQ.Item(3 + j) = Totals_All(j)
        Next
        tblASTDSQL1.Rows.Add(rowDQ)

        grdASTDSQL1.Rows.FixedRows.Clear()
        Dim RI As Integer

        If rowOthers IsNot Nothing Then
            RI = tblASTDSQL1.Rows.IndexOf(rowOthers)
            grdASTDSQL1.Rows.FixedRows.Add(grdASTDSQL1.Rows.GetRowWithListIndex(RI))
        End If

        RI = tblASTDSQL1.Rows.IndexOf(rowDQ)
        'grdASTDSQL1.Rows(grdASTDSQL1.Rows.Count - 1).Fixed = True
        grdASTDSQL1.Rows.FixedRows.Add(grdASTDSQL1.Rows.GetRowWithListIndex(RI))
    End Sub

    Sub Generate_Inquiry()

        If COLUMN_NAMEs.Count = 0 Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-Configuring Data Query Tree")

        ReDim COLUMN_NAME_by_Lvl(COLUMN_NAMEs.Count)
        ReDim COLUMN_CAPTION_by_Lvl(COLUMN_NAMEs.Count)
        ReDim G_by_Lvl(COLUMN_NAMEs.Count)
        For G As Integer = 1 To COLUMN_NAMEs.Count
            Dim GC As UltraWinGrid.UltraGridColumn = grdDQseq.DisplayLayout.Bands(0).Columns(COLUMN_NAMEs(G - 1))
            Dim Lvl As Integer = GC.Header.VisiblePosition + 1
            COLUMN_NAME_by_Lvl(Lvl) = COLUMN_NAMEs(G - 1)
            COLUMN_CAPTION_by_Lvl(Lvl) = GC.Header.Caption
            G_by_Lvl(Lvl) = G
        Next

        Dim Gs() As String = Nothing
        ReDim Gs(COLUMN_NAMEs.Count - 1)
        Dim orderby As String = ""
        For Lvl As Integer = 1 To COLUMN_NAMEs.Count
            Gs(Lvl - 1) = "G" & CStr(G_by_Lvl(Lvl))
            orderby &= ",G" & CStr(G_by_Lvl(Lvl))
        Next

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
        Dim CODE_VALUE_at_Lvl() As String = Nothing
        ReDim CODE_VALUE_at_Lvl(COLUMN_NAMEs.Count)

        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        tvwDQ.Nodes.Clear()

        Dim cur_Node_at_Lvl() As Infragistics.Win.UltraWinTree.UltraTreeNode
        ReDim cur_Node_at_Lvl(COLUMN_NAMEs.Count)
        If COLUMN_CAPTION_by_Lvl.Length = 1 Then
            aNode = tvwDQ.Nodes.Add("*", "All")
        Else
            aNode = tvwDQ.Nodes.Add("*", "All (" & COLUMN_CAPTION_by_Lvl(1) & ")")
        End If
        cur_Node_at_Lvl(0) = aNode
        Dim TBL As DataTable = ASCDATA1.SelectDistinct("ASTSRPT1", Gs)
        Dim last_level_set As Integer = 0
        If COLUMN_NAMEs.Count > 1 Then ' no nodes (other than All) when there is only 1 level
            Dim sqlx As String = ""
            For Each Gx As String In Gs
                sqlx &= " and " & Gx & " <> '" & aRC & "'"
            Next
            sqlx = Mid(sqlx, 6)
            For Each row As DataRow In TBL.Select(sqlx, Mid(orderby, 2))
                If row.Item(0) & "" <> aRC Then
                    For Lvl As Integer = 1 To COLUMN_NAMEs.Count - 1
                        If CODE_VALUE_at_Lvl(Lvl) <> row.Item(Lvl - 1) & "" Or last_level_set < Lvl Then
                            If CODE_VALUE_at_Lvl(Lvl) <> aRC Then
                                last_level_set = Lvl
                                aNode = cur_Node_at_Lvl(Lvl - 1).Nodes.Add
                                cur_Node_at_Lvl(Lvl) = aNode
                                'aNode.Key = KEY_PREFIX & KEY
                                Dim CAPTION As String = Split(row.Item(Lvl - 1) & ":", ":")(1)
                                aNode.Text = CAPTION
                                'aNode.Tag = row.Item("MENU_ID") & Chr(1) & KEY
                                aNode.Expanded = False
                                CODE_VALUE_at_Lvl(Lvl) = row.Item(Lvl - 1) & ""
                                If last_level_set = COLUMN_NAMEs.Count - 1 Then
                                    aNode.LeftImages.Add(ASCMAIN1.Get_Image(IMAGE_FOLDER, "ITEM_green")) ' "graph_node"))
                                Else
                                    aNode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M")
                                    aNode.Override.ExpandedNodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN")
                                End If
                                For iLvl As Integer = 1 To Lvl
                                    aNode.Cells(iLvl - 1).Value = CODE_VALUE_at_Lvl(iLvl)
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        End If
        cur_Node_at_Lvl(0).Expanded = True
        Click_Node(cur_Node_at_Lvl(0))

        Sort_grdColumns(grdASTDSQL1, "SORT_VALUE,CODE_VALUE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdDQseq_AfterColPosChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdDQseq.AfterColPosChanged
        Generate_Inquiry()
    End Sub

    Private Sub chkTopN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTopN.CheckedChanged
        numTopN.Visible = chkTopN.Checked
        cbeTopN.Visible = chkTopN.Checked
        If chkTopN.Checked Then
            Show_TopN()
        Else
            tblASTDSQL1.Rows.Clear()
            tblASTDSQL1.Merge(tblASTDSQL1_copy)
            'If tvwDQ.ActiveNode Is Nothing Then
            '    Click_Node(tvwDQ.Nodes(0))
            'Else
            '    Click_Node(tvwDQ.ActiveNode)
            'End If
        End If
    End Sub
#End Region

    Private Sub cbeTopN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeTopN.ValueChanged
        If chkTopN.Checked Then
            Show_TopN()
        End If
    End Sub

    Sub Show_TopN()

        tblASTDSQL1.Rows.Clear()

        Dim N As Integer = numTopN.Value
        Dim I As Integer = 0
        Dim C As Integer = 0


        Dim RR As Integer = tblASTRECAP.Rows.Count

        Dim T() As Decimal
        ReDim T(COLUMN_NAME_sum.Count * IIf(RR = 0, 1, RR))
        Dim c0 As Integer = 4 ' tblASTDSQL1.Columns(COLUMN_NAME_sum_first).Ordinal

        Dim COLUMN_NAME As String = cbeTopN.Value
        If RR <> 0 Then
            Dim ci As Integer = optRECAPSORT.CheckedIndex
            COLUMN_NAME &= "_" & CStr(ci + 1)
        End If

        Dim rowASTDSQL1 As DataRow
        For Each rowASTDSQL1_copy As DataRow In tblASTDSQL1_copy.Select _
        ("SORT_VALUE = '1'", COLUMN_NAME & " DESC", DataViewRowState.CurrentRows)
            I += 1
            If I <= N Then
                rowASTDSQL1 = tblASTDSQL1.NewRow
                rowASTDSQL1.ItemArray = rowASTDSQL1_copy.ItemArray
                tblASTDSQL1.Rows.Add(rowASTDSQL1)
            Else
                C += 1
                For j As Integer = 1 To UBound(T)
                    Dim k As Integer = j + c0 - 1
                    T(j) += Val(rowASTDSQL1_copy.Item(k) & "")
                Next
            End If
        Next
        If C > 0 Then
            rowASTDSQL1 = tblASTDSQL1.NewRow
            rowASTDSQL1.Item("CODE_VALUE") = "All Others"
            rowASTDSQL1.Item("DESC_VALUE") = CStr(C)
            rowASTDSQL1.Item("SORT_VALUE") = 2
            For j As Integer = 1 To UBound(T)
                Dim k As Integer = j + c0 - 1
                rowASTDSQL1.Item(k) = T(j)
            Next
            tblASTDSQL1.Rows.Add(rowASTDSQL1)
            Add_Totals_Row(rowASTDSQL1)
        Else
            Add_Totals_Row()
        End If

        'Sort_grd(COLUMN_NAME, True)
        Sort_grd(cbeTopN.Value, True)
    End Sub

    Private Sub numTopN_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles numTopN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter And chkTopN.Checked Then
            Show_TopN()
        End If
    End Sub

    Private Sub numTopN_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles numTopN.Leave
        If chkTopN.Checked Then
            Show_TopN()
        End If
    End Sub

    Private Sub numTopN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numTopN.ValueChanged

    End Sub

    Private Sub grdASTDSQL1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTDSQL1.Click
        Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
        Dim tt As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        'Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

        Dim grid As Infragistics.Win.UltraWinGrid.UltraGrid = DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid)

        'Get the last element that the mouse entered
        Dim lastElementEntered As Infragistics.Win.UIElement = grid.DisplayLayout.UIElement.LastElementEntered

        'See if there's a RowUIElement in the chain.
        Dim hElement As Infragistics.Win.UltraWinGrid.HeaderUIElement

        If TypeOf lastElementEntered Is Infragistics.Win.UltraWinGrid.HeaderUIElement Then
            hElement = DirectCast(lastElementEntered, Infragistics.Win.UltraWinGrid.HeaderUIElement)
        Else
            hElement = DirectCast(lastElementEntered.GetAncestor(GetType(Infragistics.Win.UltraWinGrid.HeaderUIElement)), Infragistics.Win.UltraWinGrid.HeaderUIElement)
        End If

        If Not hElement Is Nothing Then
            Dim GROUP_NAME As String = hElement.Header.Group.Key
            If GROUP_NAME = "DATA_TYPE" Then Exit Sub

            If grid.DisplayLayout.Bands(0).ColHeadersVisible Then Exit Sub
            Sort_grd(GROUP_NAME)
        End If
    End Sub

    Sub Sort_grd( _
    ByVal GROUP_NAME As String, _
    Optional ByVal force_decending As Boolean = False)

        Dim COLUMN_NAME As String = GROUP_NAME
        If tblASTRECAP.Rows.Count <> 0 Then
            If GROUP_NAME <> "CODE_VALUE" And GROUP_NAME <> "DESC_VALUE" Then
                COLUMN_NAME &= "_" & CStr(optRECAPSORT.CheckedIndex + 1)
            End If
        End If

        With grdASTDSQL1.DisplayLayout.Bands(0)
            Dim DESCENDING As Boolean = False
            If force_decending _
            Or .Groups(GROUP_NAME).Header.Appearance.ForeColor = Color.Blue Then
                DESCENDING = True
            End If
            .SortedColumns.Clear()
            For I As Integer = 0 To grdASTDSQL1.DisplayLayout.Bands(0).Groups.Count - 1
                .Groups(I).Header.Appearance.ForeColor = Color.Black
            Next
            .SortedColumns.Add("SORT_VALUE", False)
            .SortedColumns.Add(COLUMN_NAME, DESCENDING)
            If DESCENDING Then
                .Groups(GROUP_NAME).Header.Appearance.ForeColor = Color.Red
            Else
                .Groups(GROUP_NAME).Header.Appearance.ForeColor = Color.Blue
            End If
        End With

        grdASTDSQL1.DisplayLayout.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

    End Sub

    Private Sub grdASTDSQL1_DoubleClickHeader(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs) Handles grdASTDSQL1.DoubleClickHeader
        '  Stop
    End Sub

    Private Sub grdASTDSQL1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTDSQL1.InitializeLayout

    End Sub

    Private Sub grdASTDSQL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTDSQL1.InitializeRow
        Dim RR As Integer = tblASTRECAP.Rows.Count
        If RR <> 0 Then
            For I As Integer = 1 To RR
                e.Row.Cells("DATA_TYPE_" & CStr(I)).Value = DATA_TYPEs(I)
            Next
            If e.Row.Cells("SORT_VALUE").Text = "3" Then
                e.Row.CellAppearance.BackColor = Color.Khaki          ' Color.Coral     '  Color.Lavender    ' Color.LightSteelBlue
                e.Row.Appearance.BackColor = Color.Khaki          ' Color.Salmon     '  Color.Lavender    ' Color.LightSteelBlue
            ElseIf e.Row.Cells("SORT_VALUE").Text = "2" Then
                e.Row.CellAppearance.BackColor = Color.LightGray           ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
                e.Row.Appearance.BackColor = Color.LightGray        ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
            Else
                If grdASTDSQL1.DisplayLayout.Bands(0).LevelCount > 1 Then
                    Dim R As Integer = e.Row.VisibleIndex
                    ' e.Row.Cells("DATA_TYPE_1").Value = R
                    If R Mod 2 = 1 Then
                        e.Row.CellAppearance.BackColor = Color.PowderBlue      ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
                        e.Row.Appearance.BackColor = Color.PowderBlue      ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
                    Else
                        e.Row.CellAppearance.BackColor = Color.White       ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
                        e.Row.Appearance.BackColor = Color.White       ' Color.CornflowerBlue     '  Color.Lavender    ' Color.LightSteelBlue
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub optRECAPSORT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRECAPSORT.ValueChanged

    End Sub

    Private Sub chkRECAP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim chk As ABSCS.ABSCheckBox = DirectCast(sender, ABSCS.ABSCheckBox)
        Dim ROW_NO As Integer = Val(chk.Tag)
        Dim CHKS As String = grpRECAPS.Tag
        Mid(CHKS, ROW_NO, 1) = IIf(chk.Checked, "1", "0")
        grpRECAPS.Tag = CHKS

        Dim RR As Integer = tblASTRECAP.Rows.Count

        With grdASTDSQL1.DisplayLayout.Bands(0)
            Dim L As Integer = 0
            Dim CHKS_COUNT = Replace(CHKS, "0", "").Length
            If .LevelCount < CHKS_COUNT Then .LevelCount = CHKS_COUNT
            For i As Integer = 1 To RR
                Dim CHKD As String = Mid(CHKS, i, 1)
                For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
                    Dim COLUMN_NAME_sfx As String = COLUMN_NAME & "_" & CStr(i)
                    .Columns(COLUMN_NAME_sfx).Hidden = (CHKD = "0")
                    If CHKD = "1" Then .Columns(COLUMN_NAME_sfx).Level = L

                Next
                If CHKD = "1" Then L += 1
            Next
            If CHKS_COUNT > 0 Then
                .LevelCount = CHKS_COUNT
            End If
        End With


        With grdASTDSQL1.DisplayLayout.Bands(0)
            .Summaries.Clear()
        End With

        'grdASTDSQL1.DisplayLayout .Bands(0).Override .SummaryDisplayArea.BottomFixed
        Create_Summary(grdASTDSQL1, "CODE_VALUE", "Count")
        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim COLUMN_NAME As String = rowASTDSQLS.Item("COLUMN_NAME")
            For COL_sfx As Integer = Sign(RR) To RR
                Dim COLUMN_NAME_sfx As String = COLUMN_NAME
                If COL_sfx <> 0 Then
                    COLUMN_NAME_sfx &= "_" & CStr(COL_sfx)
                End If
                Create_Summary(grdASTDSQL1, COLUMN_NAME_sfx)
            Next
        Next

    End Sub

    Sub Recursive_Export_to_Excel(ByVal tnode As UltraWinTree.UltraTreeNode, _
                                  ByVal wb As Infragistics.Documents.Excel.Workbook)
        Click_Node(tnode)
        Export_to_Excel_Add_grd(wb, grdASTDSQL1, False, tnode.Text)
        If tnode.HasNodes Then
            For Each cnode As UltraWinTree.UltraTreeNode In tnode.Nodes
                Recursive_Export_to_Excel(cnode, wb)
            Next
        End If
    End Sub

    Public Overridable Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1
        clsASCBASE1 = New ASCBASE1

        Dim sqlw As String = CStr(parms(0))

        ' REPORT SPECIFIC CODE GOES HERE

        Return clsASCBASE1
    End Function

    Public Overridable Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

    End Sub

    Private Sub cmdAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAll.Click
        grdSetup.ActiveRow.Cells("CODE_VALUES").Value = ""
        DirectCast(grd.DataSource, DataTable).Rows.Clear()
        cmdAll.Visible = False
    End Sub

    Function Get_PreKey_filter(Optional ByVal prepend_leading_AND As Boolean = False) As String

        Dim FILTER As String = ""

        For Each COLUMN_NAME As String In ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Keys
            Dim CODE_VALUES As String = SQLA(COLUMN_NAME, , True)
            If InStr(CODE_VALUES, ",") <> 0 Or CODE_VALUES = "" Then
                FILTER = " and 1<>1"
                Exit For
            End If
            FILTER &= " and " & COLUMN_NAME & " in (" & CODE_VALUES & ")"
        Next

        If Not prepend_leading_AND Then
            FILTER = Mid(FILTER, 6)
        End If

        Return FILTER
    End Function

#Region "grdASTROPT4"

    Private Sub grdASTROPT4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "USER_ID"
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_NAME")
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_EMAIL")
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_STATUS")
        End Select

        If grdASTROPT4.ActiveCell.Column.Key = "USER_ID" And grdASTROPT4.ActiveCell.Value & "" <> "" AndAlso grdASTROPT4.ActiveCell.Value <> grdASTROPT4.ActiveCell.Value.ToString.ToLower Then
            grdASTROPT4.ActiveCell.Value = grdASTROPT4.ActiveCell.Value.ToString.ToLower
        End If

    End Sub

    Private Sub grdASTROPT4_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT4.AfterRowActivate
        With grdASTROPT4.DisplayLayout.Bands(0)
            If grdASTROPT4.ActiveRow.IsAddRow Then
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdASTROPT4.ActiveCell = grdASTROPT4.ActiveRow.Cells("USER_ID")
                grdASTROPT4.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdASTROPT4_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT4.AfterRowsDeleted
        For Each USER_GROUP_ID As String In USER_GROUP_IDs
            Delete_Rows("ASTROPTA", "USER_GROUP_ID = '" & USER_GROUP_ID & "'")
        Next
        USER_GROUP_IDs.Clear()
    End Sub

    Private Sub grdASTROPT4_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTROPT4.AfterRowUpdate
        If e.Row.Band.Key = "ASTROPT4" Then

            If 1 = 1 Then Exit Sub ' until we support ASTROPTA in ASFSRPTM

            Delete_Rows("ASTROPTA", "USER_GROUP_ID = '" & e.Row.Cells("USER_ID").Text & "'")

            If e.Row.Cells("USER_STATUS").Text = "G" Then
                ASCMAIN1.sql = "Select ASTUSER3.*,ASTUSER1.USER_NAME,ASTUSER1.USER_EMAIL from ASTUSER3,ASTUSER1" _
                & " where ASTUSER1.USER_ID = ASTUSER3.USER_ID and ASTUSER3.USER_GROUP_ID = :PARM1"

                For Each rowASTUSER3 As DataRow In ASCDATA1.GetDataTable _
                    (ASCMAIN1.sql, , , , , "V", New Object() {e.Row.Cells("USER_ID").Text}).Rows
                    Dim rowASTROPTA As DataRow = dst.Tables("ASTROPTA").NewRow
                    rowASTROPTA.Item("USER_GROUP_ID") = rowASTUSER3.Item("USER_GROUP_ID")
                    rowASTROPTA.Item("USER_ID") = rowASTUSER3.Item("USER_ID")
                    rowASTROPTA.Item("USER_NAME") = rowASTUSER3.Item("USER_NAME")
                    rowASTROPTA.Item("USER_EMAIL") = rowASTUSER3.Item("USER_EMAIL")
                    dst.Tables("ASTROPTA").Rows.Add(rowASTROPTA)
                Next
            End If
        End If
    End Sub

    Private Sub grdASTROPT4_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdASTROPT4.BeforeCellUpdate
        'grdFieldFormat(grdASTROPT4)
        'If grdASTROPT4.ActiveCell.Column.Key = "USER_ID" And grdASTROPT4.ActiveCell.Value & "" <> "" AndAlso grdASTROPT4.ActiveCell.Value <> grdASTROPT4.ActiveCell.Value.ToString.ToLower Then
        '    grdASTROPT4.ActiveCell.Value = grdASTROPT4.ActiveCell.Value.ToString.ToLower
        'End If
    End Sub


    Private Sub grdASTROPT4_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdASTROPT4.BeforeExitEditMode
    End Sub

    Private Sub grdASTROPT4_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTROPT4.BeforeRowsDeleted

        'If grdASTROPT4.Selected.Rows.Count <> 1 Then
        '    e.Cancel = True
        'Else
        'End If
        USER_GROUP_IDs.Clear()
        For Each gr As UltraWinGrid.UltraGridRow In grdASTROPT4.Selected.Rows
            USER_GROUP_IDs.Add(gr.Cells("USER_ID").Text)
        Next
    End Sub

    Private Sub grdASTROPT4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTROPT4.BeforeRowUpdate
        With grdASTROPT4

            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", e.Row.Cells("USER_ID").Text)
            If rowASTUSER1 Is Nothing Then
                e.Cancel = True
            End If
            If e.Row.IsAddRow Then
                .ActiveRow.Cells("FORM_NAME").Value = FORM_NAME
                .ActiveRow.Cells("SET_ID").Value = SET_ID
            End If
        End With
    End Sub

    Private Sub grdASTROPT4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT4.ClickCellButton
        Dim sql_where As String = ""
        If grdASTROPT4.ActiveCell.Column.Key = "USER_ID" Then
            'sql_where = " and VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        End If
        Call grdClickCellButton(grdASTROPT4, sql_where, False)
    End Sub

    Private Sub grdASTROPT4_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdASTROPT4.Error
        grdASTROPT4.ActiveRow.CancelUpdate()
    End Sub
#End Region

    Sub Prepare_XLS(Optional ByVal xls_where As String = "", Optional ByVal ASTSRPT1 As String = "ASTSRPT1")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Exporting Report File to Excel")

        'Dim wkb As New GemBox.Spreadsheet.ExcelFile
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

        Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile
        Dim ws As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets.Add(MENU_ITEM_OBJECT)

        Dim COLs() As String = Split(Prepare_XLS_Summary_Columns(COLUMN_NAME_sum), ",")

        Dim colors() As System.Drawing.Color = _
        {Color.Beige, Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige, _
         Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige, _
         Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige}

        'ws.Cells(R, i).Style.FillPattern.PatternForegroundColor = colors(i)
        'ws.Cells(R, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid

        Dim G_Colors(9) As System.Drawing.Color
        G_Colors(1) = Color.Purple
        G_Colors(2) = Color.Green
        G_Colors(3) = Color.DarkOrange
        G_Colors(4) = Color.Blue
        G_Colors(5) = Color.Olive
        G_Colors(6) = Color.Brown
        G_Colors(7) = Color.Gold
        G_Colors(8) = Color.DarkMagenta
        G_Colors(9) = Color.Red
        Dim ROTATE As Integer = 0
        Do While ROTATE <> 0
            ROTATE -= 1
            For CLR As Integer = 1 To 9
                G_Colors(CLR - 1) = G_Colors(CLR)
            Next
            G_Colors(9) = G_Colors(0)
            For CLR As Integer = 1 To 3
                colors(CLR - 1) = colors(CLR)
            Next
            colors(3) = colors(0)
        Loop

        Dim FS As New Dictionary(Of String, String)
        Dim XLC As New Dictionary(Of String, String)

        Dim C As Integer = 0
        Dim R As Integer = 0
        Dim GMAX As Integer = COLUMN_NAMEs.Count

        C = GMAX + 1 + 1
        Dim FORMATS As New Dictionary(Of String, String)
        For Each SCN As String In COLs
            C += 1
            Dim FORMAT As String = ""
            With grdASTSRPT1.DisplayLayout.Bands(0).Columns(SCN)
                If COLUMN_NAME_sum.ContainsKey(SCN) Then
                    Select Case COLUMN_NAME_sum(SCN)
                        Case "QTY"
                            FORMAT = "#,##0"
                        Case "AMT"
                            FORMAT = "#,##0.00"
                        Case "DEC"
                            FORMAT = "#,##0.00"
                    End Select
                Else
                    Select Case dst.Tables("ASTSRPT1").Columns(SCN).DataType.ToString
                        Case "System.Int64", "System.Int32", "System.Integer"
                            FORMAT = "#,##0"
                        Case "System.Decimal"
                            FORMAT = "#,##0.00"
                        Case Else
                            FORMAT = ""
                    End Select
                End If
            End With
            FORMATS.Add(SCN, FORMAT)
            If dst.Tables("ASTSRPT1").Columns(SCN).Expression <> "" Then
                Dim FORMULA As String = "=" & Replace(dst.Tables("ASTSRPT1").Columns(SCN).Expression, "IIF", "IF")
                FS.Add(SCN, FORMULA)
            End If

            Dim CP As Integer = (C - 1) \ 26
            Dim XL As String = Chr(64 + C - CP * 26)
            If CP > 0 Then
                XL = Chr(64 + CP) & XL
            End If
            XLC.Add(SCN, XL & "#")
        Next

        For C = 1 To GMAX + 1 + COLs.Length
            ws.Columns(C - 1).Style.Font.Name = "Verdana"
        Next

        ws.Cells(1, 0).Style.Font.Color = Color.Blue
        ws.Cells(1, 0).Style.Font.Size = 300
        ws.Cells(1, 0).Style.Font.Name = "Times New Roman"
        ws.Cells(1, 0).Value = MENU_ITEM_DESC
        ws.Cells(0, 1).Value = TABLE_NAME
        ws.Cells(2, 0).Value = SUBT
        R = 3

        With ws.Cells(0, 0)
            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Left
            .Style.NumberFormat = "mm/dd/yy;@"
            .Value = Now
        End With

        Dim XL1 As Integer = 0
        Dim XL2 As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""


        Dim G() As String = Nothing
        Dim GK() As String = Nothing
        Dim B As Integer = 0
        Dim ST() As String = Nothing

        R += 1
        For C = 1 To GMAX
            ws.Cells(R - 1, C - 1).Value = COLUMN_CAPTION_by_Lvl(C)
        Next
        ws.Cells(R - 1, C - 1).Value = "Description"

        C += 1
        For Each SCN As String In COLs
            Dim row() As DataRow = tblASTDSQLS.Select("COLUMN_NAME = '" & SCN & "'")
            Dim COLUMN_CAPTION As String = SCN
            If row.Length = 1 Then COLUMN_CAPTION = row(0).Item("COLUMN_CAPTION") & ""
            C += 1
            ws.Cells(R - 1, C - 1).Value = COLUMN_CAPTION
        Next

        For C = 1 To GMAX + 1 + 1 + COLs.Length
            ws.Cells(R - 1, C - 1).Style.FillPattern.PatternForegroundColor = Color.LightGray
            ws.Cells(R - 1, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
        Next


        Dim GS As String = ""
        For I As Integer = 1 To GMAX
            GS &= "," & "G" & CStr(I)
        Next

        Dim IMGFILENAME2 As String = ""

        Dim sqlw As String = ""
        For I As Integer = 1 To GMAX
            sqlw &= " and G" & CStr(I) & " <> '" & aRC & "'"
        Next
        sqlw = Mid(sqlw, 5)

        If xls_where <> "" Then
            sqlw &= " and " & xls_where
        End If

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select(sqlw, Mid(GS, 2))
            For I As Integer = 1 To GMAX
                If G Is Nothing OrElse GK(I) <> row.Item("G" & CStr(I)) & "" Then
                    B = I

                    If G Is Nothing Then
                        ' REPORT HEADING
                        ReDim G(GMAX)
                        ReDim GK(GMAX)
                        ReDim ST(GMAX)
                    Else
                        If B < GMAX Then
                            Prepare_XLS_SubTotals(B, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, ws)
                            XL1 = 0
                            XL2 = 0
                        End If
                    End If

                    For J As Integer = B To GMAX
                        GROUP_KEY = row.Item("G" & CStr(J)) & ""
                        rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)

                        GK(J) = GROUP_KEY
                        G(J) = rowASTGROUP.Item("GROUP_CODE")
                        GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                        R += 1 ' HEADING

                        'ws.Rows(R - 1).Style.Font.Color = G_Colors(J)

                        ws.Rows(R - 1).OutlineLevel = GMAX
                        ws.Cells(R - 1, GMAX).Value = GROUP_DESC
                        ws.Cells(R - 1, GMAX).Style.Indent = J - 1
                        If J <> GMAX Then
                            ws.Cells(R - 1, GMAX).Style.Font.Color = G_Colors(J)
                        End If

                        For C = 1 To J
                            ws.Cells(R - 1, C - 1).Value = G(C)
                            If C <> GMAX Then
                                ws.Cells(R - 1, C - 1).Style.Font.Color = G_Colors(C)
                            End If
                        Next

                        If J <> GMAX Then
                            ws.Rows(R - 1).Style.Font.Color = G_Colors(J)

                            'For C = 1 To GMAX + 1
                            '    ws.Cells(R - 1, C - 1).Style.FillPattern.PatternForegroundColor = colors(J) ' Color.Beige
                            '    ws.Cells(R - 1, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            'Next
                        End If
                    Next
                End If
            Next

            Prepare_XLS_Prepare_row(row)

            C = GMAX + 1 + 1
            For Each SCN As String In COLs
                C += 1
                If FS.ContainsKey(SCN) Then
                    Dim FORMULA As String = FS(SCN)
                    For Each SCN2 As String In COLs
                        If InStr(FORMULA, SCN2) <> 0 Then
                            FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                        End If
                    Next
                    FORMULA = Replace(FORMULA, "#", CStr(R))
                    ws.Cells(R - 1, C - 1).Formula = FORMULA

                Else
                    ws.Cells(R - 1, C - 1).Value = row.Item(SCN)
                End If
            Next

            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("4D")

            Dim IMGFILENAME As String = ""
            Dim col_for_image As Integer = 0

            Try
                ws.Rows(R - 1).OutlineLevel = GMAX
                If XL1 = 0 Then XL1 = R
                XL2 = R


                IMGFILENAME = Prepare_XLS_GetImage(row, GMAX, col_for_image)

                If IMGFILENAME2 <> "" Then
                    IMGFILENAME = IMGFILENAME2
                End If
                'IMGFILENAME = ""
            Catch ex As Exception
                If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
            End Try


            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("4E")

            If IMGFILENAME <> "" Then
                IMGFILENAME2 = IMGFILENAME
                ws.Rows(R - 1).Height = ws.Rows(R - 1).Height * 4
                If My.Computer.FileSystem.FileExists(IMGFILENAME) Then
                    Try
                        ws.Pictures.Add(IMGFILENAME, _
                        GemBox.Spreadsheet.PositioningMode.MoveAndSize, _
                        New GemBox.Spreadsheet.AnchorCell(ws.Columns(col_for_image), ws.Rows(R - 1), True), _
                        New GemBox.Spreadsheet.AnchorCell(ws.Columns(col_for_image), ws.Rows(R - 1), False))

                    Catch ex As Exception
                        Stop

                    End Try
                End If
            End If
        Next

        If ST Is Nothing Then Exit Sub

        Prepare_XLS_SubTotals(0, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, ws)

        Dim xlsFileName As String = MENU_ITEM_OBJECT

        Dim tryagain As Integer = 0
        Dim FILENAME As String = ""

        Do
            xlsFileName = MENU_ITEM_OBJECT
            If tryagain > 0 Then
                xlsFileName &= "_" & CStr(tryagain)
            End If
            'FILENAME = ASCMAIN1.Folders("Temp") & xlsFileName & ".xlsx"
            FILENAME = ASCMAIN1.Folders("Temp") & xlsFileName & ".xls"

            If My.Computer.FileSystem.FileExists(FILENAME) Then
                Try
                    My.Computer.FileSystem.DeleteFile(FILENAME)
                Catch ex As Exception

                End Try
            End If

            Try
                'myWorkbook.SaveXlsx(FILENAME)
                myWorkbook.SaveXls(FILENAME)
                tryagain = -1
            Catch ex As Exception
                tryagain += 1
            End Try

        Loop While tryagain >= 0 And tryagain < 10

        myWorkbook.ClosePreservedXlsx()
        myWorkbook = Nothing

        Dim excel As New Process
        excel.StartInfo.Arguments = """" + xlsFileName + """ /e"
        excel.StartInfo.FileName = FILENAME
        excel.Start()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Overridable Function Prepare_XLS_Summary_Columns( _
    ByVal COLUMN_NAME_sum As Dictionary(Of String, String)) _
    As String
        Dim COLUMN_NAMEs As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            COLUMN_NAME &= "," & COLUMN_NAME
        Next
        Return Mid(COLUMN_NAMEs, 2)
    End Function

    Overridable Function Prepare_XLS_GetImage( _
    ByVal row As DataRow, _
    ByVal GMAX As Integer, _
    ByRef col As Integer) As String
        Return ""
    End Function

    Overridable Sub Prepare_XLS_Prepare_row(ByVal row As DataRow)

    End Sub

    Sub Prepare_XLS_SubTotals( _
    ByVal B As Integer, _
    ByRef R As Integer, _
    ByVal GMAX As Integer, _
    ByVal XL1 As Integer, _
    ByVal XL2 As Integer, _
    ByVal ST() As String, _
    ByVal G() As String, _
    ByVal GK() As String, _
    ByVal COLs() As String, _
    ByVal FS As Dictionary(Of String, String), _
    ByVal XLC As Dictionary(Of String, String), _
    ByVal G_Colors() As System.Drawing.Color, _
    ByVal ws As GemBox.Spreadsheet.ExcelWorksheet)

        Dim C As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""

        For Slvl As Integer = GMAX - 1 To B Step -1

            R += 1 ' SUB-TOTAL
            ws.Rows(R - 1).Style.Font.Color = G_Colors(Slvl)
            For J As Integer = Slvl To 1 Step -1
                ws.Cells(R - 1, J - 1).Value = G(J)
                ws.Cells(R - 1, J - 1).Style.Font.Color = G_Colors(J)
            Next

            ST(Slvl) &= ",X" & CStr(R)

            If Slvl = 0 Then
                GROUP_DESC = "Totals"
            Else
                GROUP_KEY = GK(Slvl)
                rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)
                GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                ws.Cells(R - 1, GMAX).Style.Indent = Slvl - 1
            End If
            ws.Cells(R - 1, GMAX).Value = GROUP_DESC
            ws.Cells(R - 1, GMAX).Style.Font.Color = G_Colors(Slvl)

            C = GMAX + 1 + 1
            For Each SCN As String In COLs
                C += 1
                Dim CP As Integer = (C - 1) \ 26
                Dim XL As String = Chr(64 + C - CP * 26)
                If CP > 0 Then
                    XL = Chr(64 + CP) & XL
                End If

                If FS.ContainsKey(SCN) Then
                    Dim FORMULA As String = FS(SCN)
                    For Each SCN2 As String In COLs
                        If InStr(FORMULA, SCN2) <> 0 Then
                            FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                        End If
                    Next
                    FORMULA = Replace(FORMULA, "#", CStr(R))
                    ws.Cells(R - 1, C - 1).Formula = FORMULA

                Else
                    If Slvl = GMAX - 1 Then
                        ws.Cells(R - 1, C - 1).Formula = "=SUM(" & XL & XL1 & ":" & XL & XL2 & ")"
                    Else
                        ws.Cells(R - 1, C - 1).Formula = "=SUM(" & Replace(Mid(ST(Slvl + 1), 2), "X", XL) & ")"
                    End If
                End If

                ws.Rows(R - 1).OutlineLevel = Slvl
                ws.Cells(R - 1, C - 1).Style.Font.Color = G_Colors(Slvl)
            Next
            ST(Slvl + 1) = ""

            Dim CC As System.Drawing.Color = Color.PaleGoldenrod
            If Slvl = 0 Then CC = Color.PaleGreen

            For C = 1 To GMAX + 1 + 1 + COLs.Length
                ws.Cells(R - 1, C - 1).Style.FillPattern.PatternForegroundColor = CC
                ws.Cells(R - 1, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
            Next
            R += 1
            ws.Rows(R - 1).Height = ws.Rows(R - 1).Height * 0.25
        Next

    End Sub

    Private Sub grdSetup_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSetup.BeforeCellUpdate
        ' e.NewValue = "x"
        'If e.Cell.Column.Key = "CODE_VALUES" Then
        '    Dim CODE_VALUES As String = e.Cell.Value & ""
        '    If InStr(CODE_VALUES, vbCrLf) <> 0 Then
        '        CODE_VALUES = Replace(CODE_VALUES, vbCrLf, ",")
        '        If CODE_VALUES.EndsWith(",") Then
        '            CODE_VALUES = Mid(CODE_VALUES, 1, Len(CODE_VALUES) - 1)
        '        End If
        '        e.Cell.Value = CODE_VALUES
        '    End If
        'End If

    End Sub

    Private Sub grdSetup_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdSetup.KeyUp
        'If e.KeyCode = Keys.V And e.Control Then

        '    If grdSetup.ActiveCell IsNot Nothing AndAlso grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
        '        Dim CODE_VALUES As String = grdSetup.ActiveCell.Value & ""
        '        If InStr(CODE_VALUES, vbCrLf) <> 0 Then
        '            CODE_VALUES = Replace(CODE_VALUES, vbCrLf, ",")
        '            If CODE_VALUES.EndsWith(",") Then
        '                CODE_VALUES = Mid(CODE_VALUES, 1, Len(CODE_VALUES) - 1)
        '            End If
        '            grdSetup.ActiveCell.Value = CODE_VALUES
        '            grdSetup.ActiveRow.Update()
        '        End If
        '    End If

        'End If
    End Sub

    Public Overrides Function Data_Export_Context() As ABSolution.ASFBASE0.Data_Export_Entity

        Dim E As New Data_Export_Entity
        E.enabled = True
        ASTDATA1s.Clear()
        For Each T As DataTable In dst.Tables
            If Not T.TableName Like "AST*" Then
                ASTDATA1s.Add(T.TableName, T.TableName)
            Else
                If T.TableName = "ASTSRPT1" Or T.TableName = "ASTGROUP" Then ASTDATA1s.Add(T.TableName, T.TableName)
            End If
        Next
        Return E
    End Function

    Private Sub cmdPivot_Click(sender As System.Object, e As System.EventArgs) Handles cmdPivot.Click
        Dim dt As DataTable = dst.Tables("ASTSRPT1").Copy
        Pivot_Prepare(dt)
        Pivot_Show(dt)
    End Sub

    Public Overridable Sub Pivot_Prepare(dt As DataTable)
        Pivot_Prepare_PreProcess(dt)

        For I As Integer = 1 To 9
            If I > COLUMN_NAMEs.Count Then
                dt.Columns.Remove("G" & CStr(I))
            Else
                dt.Columns(I - 1).ColumnName = COLUMN_CAPTIONs(I - 1)
            End If
        Next

        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim COLUMN_NAME As String = rowASTDSQLS.Item("COLUMN_NAME")
            Dim COLUMN_CAPTION As String = rowASTDSQLS.Item("COLUMN_CAPTION")
            dt.Columns(COLUMN_NAME).ColumnName = COLUMN_CAPTION
        Next

        For Each row As DataRow In dt.Rows
            For i As Integer = 1 To COLUMN_NAMEs.Count
                If row.Item(i - 1).ToString.StartsWith(aRC) Then
                Else
                row.Item(i - 1) = Split(row.Item(i - 1), ":", 2)(1)
                End If
            Next
        Next

        Pivot_Prepare_PostProcess(dt)
    End Sub

    Public Overridable Sub Pivot_Prepare_PreProcess(dt As DataTable)

    End Sub

    Public Overridable Sub Pivot_Prepare_PostProcess(dt As DataTable)

    End Sub

    Public Overridable Sub Pivot_Show(dt As DataTable)
        'Dim F As New ABS_OLAP.ASFPIVOT(dt, Me.Text)
        'F.Show()
    End Sub

    Public Overridable Sub OLAP_Show(dt As DataTable)
        Dim F As New ABS_OLAP.ASFOLAP1(dt, Me.Text)
        F.Show()
    End Sub

    Private Sub cmdOLAP_Click(sender As System.Object, e As System.EventArgs) Handles cmdOLAP.Click
        Dim dt As DataTable = dst.Tables("ASTSRPT1").Copy
        Pivot_Prepare(dt)
        OLAP_Show(dt)
    End Sub

    Public Overridable Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Return ""
    End Function

    Function Get_Level_for(COLUMN_NAME As String) As Integer
        Dim Level As Integer = 0
        'If COLUMN_NAME_by_Lvl IsNot Nothing AndAlso COLUMN_NAME_by_Lvl.Count > 0 Then
        '    For I As Integer = 1 To COLUMN_NAME_by_Lvl.Count - 1
        '        If COLUMN_NAME_by_Lvl(I) = COLUMN_NAME Then
        '            Level = I
        '            Exit For
        '        End If
        '    Next
        'End If
        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
        If rowASTDSQLA IsNot Nothing Then
            Level = Val(rowASTDSQLA.Item("SEQUENCE") & "")
        End If
        Return Level
    End Function

    Function Get_Data(TABLE_NAME As String, _
        ByVal sql_Sum As String, _
        ByVal sql_Sum_Cols As String, _
        ByVal sql_filter As String, _
        ByVal sql_filter2 As String, _
        ByVal sql_Having As String, _
        ByVal sql_Appended_Cols As String) As Int64

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & sql_Appended_Cols & vbCrLf & sql_Sum _
            & " from " & TABLE_NAME & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
            & " group by " & IIf(sql_GROUP_BY_cols = "", "'x'", sql_GROUP_BY_cols) & vbCrLf _
            & sql_Appended_Cols & vbCrLf _
            & IIf(sql_Having = "", "", " having " & sql_Having)

        If sql_Sum_Cols.StartsWith(",") Then
            sql_Sum_Cols = Mid(sql_Sum_Cols, 2)
        End If

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & "," & sql_Sum_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        Return ASCDATA1.ExecuteSQL()
    End Function

    Public Overridable Sub Set_Parameters(Optional JOB_PARMS As Dictionary(Of String, String) = Nothing)

    End Sub

End Class