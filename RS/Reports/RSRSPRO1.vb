Imports System.Threading
Imports System.Reflection

Public Class RSRSPRO1
    Dim tblARTCUST1 As DataTable
    Dim tblARTCUST2 As DataTable
    Dim sqlARTCUST2 As String = ""

    Dim ARTCUST2 As String

    Dim rowGLTPARM3 As DataRow
    Dim RYM As String
    Dim REL_WEEK As Integer
    Dim WEEK_LEGEND As String

    Dim RYW_2LY As String
    Dim RYM01 As String

    Dim CUST_CODE As String
    Dim CUST_STORE_NOs As New List(Of String)

    Dim WKS(10, 1) As String

#Region "Thread"

    Dim RSTSPRO1 As String
    Dim RSTSPRO2 As String
    Dim RSTSPRO3 As String
    Dim RSTSPRO4 As String
    Dim RSTSPRO5 As String
    Dim RSTSPRO6 As String
    Dim RSTSPRO7 As String
    Dim RSTSPRO8 As String

    Dim RSTSPROB As String
    Dim RSTSPROC As String
    Dim RSTSPROE As String

#End Region
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
        tblARTCUST1 = ASCDATA1.GetDataTable

        sqlARTCUST2 = "Select CUST_CODE, CUST_STORE_NO, CUST_STORE_NAME, CUST_STORE_ADDR1, CUST_STORE_CITY, CUST_STORE_STATE, MALL_CODE, SELL_CODE from ARTCUST2"
        ASCMAIN1.sql = sqlARTCUST2 & " where ROWNUM < 1"
        tblARTCUST2 = ASCDATA1.GetDataTable
        tblARTCUST2.Columns.Add("SEL")
        tblARTCUST2.Columns("SEL").DefaultValue = "0"

        grdARTCUST2.DataSource = tblARTCUST2
        With grdARTCUST2.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("CUST_CODE").Hidden = True
            .Columns("CUST_STORE_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.Gold
                End With

                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = System.Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        Create_Summary(grdARTCUST2, "SEL")
        Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")
        grdARTCUST2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Show_Filter(grdARTCUST2, True)
        grdARTCUST2.DisplayLayout.GroupByBox.Hidden = False

    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        CUST_CODE = txtCUST_CODE.Text

        CUST_STORE_NOs.Clear()

        For Each row As DataRow In tblARTCUST2.Select("SEL='1'")
            CUST_STORE_NOs.Add(row.Item("CUST_STORE_NO"))
        Next


        If 1 = 1 Then
            Main_Process_1()
        Else

            'create a thread to handle communication with connected client
            Dim clientThread As New Thread(New ParameterizedThreadStart(AddressOf HandleClientComm))

            Dim dic As New Dictionary(Of String, Object)
            dic.Add("CUST_CODE", CUST_CODE)
            dic.Add("RYW", RYW)
            dic.Add("tblARTCUST2", tblARTCUST2)

            clientThread.Start(dic) ' dic

            ' PACK ALL UI SETTINGS INTO A CLASS OR KEY/VALUE PAIR

            ' INSTANTIATE RSCSPRO1 ON A NEW THREAD

            ' MAKE SURE FORM IS DISABLED
            ' ADD KILL SWITCH

        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub HandleClientComm(client As Object)
        Dim ads As New AppDomainSetup()
        ads.ShadowCopyFiles = True

        Try
            ' EventLog.WriteEntry(eventLogSource, "Creating new app domain...", EventLogEntryType.Information, 4)
            Dim ad = AppDomain.CreateDomain("", Nothing, ads)

            ' EventLog.WriteEntry(eventLogSource, "Creating ABS control class...", EventLogEntryType.Information, 4)
            Dim directory = AppDomain.CurrentDomain.BaseDirectory
            Dim ABSDomain = CType(ad.CreateInstanceFromAndUnwrap(directory & "ABSDomain.dll", "ABSDomain.ABSDomain", True, BindingFlags.CreateInstance, Nothing, Nothing, Nothing, Nothing), ABSDomain)
            ' EventLog.WriteEntry(eventLogSource, "Starting ABS control class...", EventLogEntryType.Information, 4)
            ABSDomain.Start(Nothing)
            '  EventLog.WriteEntry(eventLogSource, "Unloading app domain...", EventLogEntryType.Information, 100)
            AppDomain.Unload(ad)
        Catch ex As Exception
            '  EventLog.WriteEntry(eventLogSource, String.Format("App domain error: {0}", ex.Message), EventLogEntryType.Error, 50)
            MsgBox(ex.Message)
        End Try
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

        Dim MO As String = Format(Val(Mid(RYM, 5, 2)) - 1, "00")
        If MO = "00" Then MO = "12"

        RPT = "RSRSPRO0"
        CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
        CR_params.Add("TY", Mid(RYW, 1, 4))
        CR_params.Add("LY", Format(Val(Mid(RYW, 1, 4)) - 1, "0000"))
        CR_params.Add("MO", MO)
        Generate_Report(RPT, "Store Profile", SUBT)


        Dim TBL As DataTable = dst.Tables("ARTCUST2").Copy

        For Each CUST_STORE_NO As String In CUST_STORE_NOs

            dst.Tables("ARTCUST2").Rows.Clear()
            Dim row As DataRow = TBL.Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            dst.Tables("ARTCUST2").Rows.Add(row.ItemArray)

            Dim CUST_STORE_NAME As String = row.Item("CUST_STORE_NAME") & ""
            If CUST_STORE_NAME = "" Then CUST_STORE_NAME = CUST_CODE & "_" & CUST_STORE_NO
            CUST_STORE_NAME = Replace(CUST_STORE_NAME, "#", "")
            Dim FILENAME As String = CUST_STORE_NAME & "_" & XNO

            CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
            CR_params.Add("TY", Mid(RYW, 1, 4))
            CR_params.Add("LY", Format(Val(Mid(RYW, 1, 4)) - 1, "0000"))
            CR_params.Add("MO", MO)

            Generate_Report(RPT, CUST_STORE_NAME, "Store Profile")

            'Dim REPORT_NO As String = Generate_Report(RPT, "Store Profile", CUST_STORE_NAME, "", "PDF", FILENAME, False)
            'Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & "." & "PDF")
        Next
        '    Print_Report_End(, True)

        dst.Tables("ARTCUST2").Rows.Clear()
        dst.Tables("ARTCUST2").Merge(TBL)


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If txtCUST_CODE.Text = "" Then
                EMsg &= vbCr & "You Must Specify a Customer"
            Else
                If LookUp("ARTCUST1", txtCUST_CODE.Text) Is Nothing Then
                    EMsg &= vbCr & "Customer Specified (" & txtCUST_CODE.Text & ") is Invalid"
                Else
                    If tblARTCUST2.Select("SEL='1'").Length = 0 Then
                        EMsg &= vbCr & "You Must Select 1 Customer Store"
                    End If
                    'If tblARTCUST2.Select("SEL='1'").Length > 1 Then
                    '    EMsg &= vbCr & "You Must Select Only 1 Customer Store"
                    'End If
                End If
            End If
        End If
    End Sub

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        Dim CUST_CODE = txtCUST_CODE.Text
        If CUST_CODE = "" OrElse tblARTCUST1.Rows.Find(CUST_CODE) Is Nothing Then
            grdARTCUST2.Visible = False
        Else
            grdARTCUST2.Visible = True

            tblARTCUST2.Rows.Clear()
            ASCMAIN1.sql = sqlARTCUST2 & " where CUST_CODE = '" & CUST_CODE & "'"
            tblARTCUST2.Merge(ASCDATA1.GetDataTable)
            Sort_grdColumns(grdARTCUST2, "CUST_STORE_NO")
        End If
    End Sub

    Sub Main_Process_1()

        Create_Worktable()

        ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC" & vbCrLf _
            & ", COLLECTION_CODE, PROD_CODE, ITEM_RETAIL_PRICE, ITEM_EAN_CODE, ITEM_SNU_CODE, ITEM_CLASS_CODE" & vbCrLf _
            & " from ICTITEM1 where ITEM_CODE in " & vbCrLf _
            & " (Select Distinct ITEM_CODE from " & RSTSPRO1 & " union Select Distinct ITEM_CODE from " & RSTSPRO4 & " union Select Distinct ITEM_CODE from " & RSTSPROE & ")"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False)
        Fill_Records("ICTITEM1")

        ASCMAIN1.sql = "Select * from ICTCOLL1"
        Create_TDA(dst.Tables.Add, "ICTCOLL1", "**", 0, False)
        Fill_Records("ICTCOLL1")

        ASCMAIN1.sql = "Select * from ICTCOLL0"
        Create_TDA(dst.Tables.Add, "ICTCOLL0", "**", 0, False)
        Fill_Records("ICTCOLL0")

        ASCMAIN1.sql = "Select * from ICTBRAN1"
        Create_TDA(dst.Tables.Add, "ICTBRAN1", "**", 0, False)
        Fill_Records("ICTBRAN1")

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME" & vbCrLf _
            & " from ARTCUST1 where CUST_CODE in " & vbCrLf _
            & " (Select Distinct CUST_CODE from " & ARTCUST2 & ")"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False)
        Fill_Records("ARTCUST1")

        ASCMAIN1.sql = "Select *" & vbCrLf _
            & " from ARTCUST2 where (CUST_CODE, CUST_STORE_NO) in " & vbCrLf _
            & " (Select Distinct CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")"
        Create_TDA(dst.Tables.Add, "ARTCUST2", "**", 0, False)
        Fill_Records("ARTCUST2")

        ASCMAIN1.sql = "Select * from " & RSTSPRO2
        Create_TDA(dst.Tables.Add, "RSTSPRO2", "**", 0, False)
        Fill_Records("RSTSPRO2")
        With dst.Tables("RSTSPRO2")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPRO3
        Create_TDA(dst.Tables.Add, "RSTSPRO3", "**", 0, False)
        Fill_Records("RSTSPRO3")
        With dst.Tables("RSTSPRO3")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPRO8
        Create_TDA(dst.Tables.Add, "RSTSPRO8", "**", 0, False, "", 3)
        Fill_Records("RSTSPRO8")

        'With dst.Tables("RSTSPRO8")
        '    .Columns.Add("TY_WTD", GetType(System.Decimal), "1000 * TY_WK" & CStr(REL_WEEK))
        '    .Columns.Add("LY_WTD", GetType(System.Decimal), "1000 * LY_WK" & CStr(REL_WEEK))
        'End With

        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & RYM & "'"
        Create_TDA(dst.Tables.Add, "GLTPARM3", "**", 0, False, "", 1)

        ASCMAIN1.sql = "Select * from GLTPARM3"
        Create_TDA(dst.Tables.Add, "GLTPARM3_ALL", "**", 0, False, "", 1)

        Fill_Records("GLTPARM3")

        dst.Tables("GLTPARM3").Columns.Add("LEGEND2")
        For Each row As DataRow In dst.Tables("GLTPARM3").Select("")
            Dim LEGEND As String = row.Item("LEGEND")
            LEGEND = Mid(LEGEND, 10, 7)
            row.Item("LEGEND2") = Mid(LEGEND, 1, 3) & " Wk " & Mid(LEGEND, 5, 1)
        Next

        Fill_Records("GLTPARM3_ALL")

        dst.Tables("GLTPARM3_ALL").Columns.Add("LEGEND2")
        For Each row As DataRow In dst.Tables("GLTPARM3_ALL").Select("")
            Dim LEGEND As String = row.Item("LEGEND")
            LEGEND = Mid(LEGEND, 10, 7)
            row.Item("LEGEND2") = Mid(LEGEND, 1, 3) & " Wk " & Mid(LEGEND, 5, 1)
        Next



        With dst.Tables.Add("RSTSPRO9")
            .Columns.Add("CUST_CODE")
            .Columns.Add("CUST_STORE_NO")
            .Columns.Add("CHECKBOOK")
            .Columns.Add("OPS_YYYYWW")
            .Columns.Add("TY_WTD", GetType(System.Decimal))
            .Columns.Add("LY_WTD", GetType(System.Decimal))
            .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("CUST_STORE_NO"), .Columns("CHECKBOOK"), .Columns("OPS_YYYYWW")}
        End With

        With dst.Tables.Add("RSTSPROA")
            .Columns.Add("CUST_CODE")
            .Columns.Add("CUST_STORE_NO")
            .Columns.Add("OPS_YYYYWW")
            .Columns.Add("TY_WTD", GetType(System.Decimal))
            .Columns.Add("LY_WTD", GetType(System.Decimal))
            .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("CUST_STORE_NO"), .Columns("OPS_YYYYWW")}
        End With


        ASCMAIN1.sql = "Select * from " & RSTSPROC
        Create_TDA(dst.Tables.Add, "RSTSPROC", "**", 0, False, "", 0)
        Fill_Records("RSTSPROC")
        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, OPS_YYYYWW, SUM (TY_SLS) TY_SLS, SUM (LY_SLS) LY_SLS from " & RSTSPROC & " group by CUST_CODE, CUST_STORE_NO, OPS_YYYYWW"
        Create_TDA(dst.Tables.Add, "RSTSPROD", "**", 0, False, "", 0)
        Fill_Records("RSTSPROD")



        For Each row As DataRow In dst.Tables("RSTSPRO8").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim CHECKBOOK As String = row.Item("CHECKBOOK")
            'dst.Tables("RSTSPRO9").Rows.Clear()
            For I As Integer = 1 To 6
                If WKS(I, 0) <> "" Then
                    Dim rowRSTSPRO9 As DataRow = dst.Tables("RSTSPRO9").NewRow
                    With rowRSTSPRO9
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                        .Item("CHECKBOOK") = CHECKBOOK
                        .Item("OPS_YYYYWW") = WKS(I, 0)
                        .Item("TY_WTD") = 1000 * Val(row.Item("TY_WK" & Format(I, "0")) & "")
                        .Item("LY_WTD") = 1000 * Val(row.Item("LY_WK" & Format(I, "0")) & "")
                    End With
                    dst.Tables("RSTSPRO9").Rows.Add(rowRSTSPRO9)

                    Dim rowRSTSPROA As DataRow = dst.Tables("RSTSPROA").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, WKS(I, 0)})
                    If rowRSTSPROA Is Nothing Then
                        rowRSTSPROA = dst.Tables("RSTSPROA").NewRow
                        With rowRSTSPROA
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("CUST_STORE_NO") = CUST_STORE_NO
                            .Item("OPS_YYYYWW") = WKS(I, 0)
                            .Item("TY_WTD") = 0
                            .Item("LY_WTD") = 0
                        End With
                        dst.Tables("RSTSPROA").Rows.Add(rowRSTSPROA)
                    End If
                    With rowRSTSPROA
                        .Item("TY_WTD") += 1000 * Val(row.Item("TY_WK" & Format(I, "0")) & "")
                        .Item("LY_WTD") += 1000 * Val(row.Item("LY_WK" & Format(I, "0")) & "")
                    End With
                End If
            Next
        Next

        Dim sqlA As String = "" _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.OPS_YYYYWW, SUM (RSTRETL1.AMT_SOLD) SLS" & vbCrLf _
            & " from RSTRETL1,ICTITEM1" & vbCrLf _
            & " where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and RSTRETL1.CUST_STORE_NO in (Select CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW between X000000X and X000001X" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.OPS_YYYYWW"

        ASCMAIN1.sql = "" _
            & Replace(Replace(sqlA, "X000000X", "'" & WKS(1, 0) & "'"), "X000001X", "'" & WKS(REL_WEEK, 0) & "'") _
            & vbCrLf & " union " _
            & Replace(Replace(sqlA, "X000000X", "'" & WKS(1, 1) & "'"), "X000001X", "'" & WKS(REL_WEEK, 1) & "'")

        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        tbl.PrimaryKey = New DataColumn() {tbl.Columns("CUST_CODE"), tbl.Columns("CUST_STORE_NO"), tbl.Columns("OPS_YYYYWW")}

        For Each rowRSTSPROA As DataRow In dst.Tables("RSTSPROA").Select("")
            Dim CUST_CODE As String = rowRSTSPROA.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowRSTSPROA.Item("CUST_STORE_NO")
            Dim OPS_YYYYWW As String = rowRSTSPROA.Item("OPS_YYYYWW")

            Dim TY_WTD As Decimal = 0
            Dim LY_WTD As Decimal = 0

            Dim row As DataRow
            row = tbl.Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, OPS_YYYYWW})
            If row IsNot Nothing Then
                TY_WTD = Val(row.Item("SLS") & "")
            End If
            Dim OPS_YYYYWW_ly As String = ASCMAIN1.Week_Calc(OPS_YYYYWW, -52)
            If ASCMAIN1.CLIENT = "INT" Then
                OPS_YYYYWW_ly = Format(Val(Mid(OPS_YYYYWW, 1, 4)) - 1, "0000") & Mid(OPS_YYYYWW, 5, 2)
            End If
            row = tbl.Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, OPS_YYYYWW_ly})
            If row IsNot Nothing Then
                LY_WTD = Val(row.Item("SLS") & "")
            End If

            rowRSTSPROA.Item("TY_WTD") = TY_WTD
            rowRSTSPROA.Item("LY_WTD") = LY_WTD
        Next


        ASCMAIN1.sql = "Select * from " & RSTSPRO4
        Create_TDA(dst.Tables.Add, "RSTSPRO4", "**", 0, False, "", 3)
        Fill_Records("RSTSPRO4")

        ASCMAIN1.sql = "Select * from " & RSTSPRO7 & " where TY_STD454 <> 0 or LY_STD454 <> 0"
        Create_TDA(dst.Tables.Add, "RSTSPRO7", "**", 0, False, "", 2)
        Fill_Records("RSTSPRO7")
        With dst.Tables("RSTSPRO7")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPROE
        Create_TDA(dst.Tables.Add, "RSTSPROE", "**", 0, False, "", 5)
        Fill_Records("RSTSPROE")

        ASCMAIN1.sql = "Select * from " & RSTSPRO6 & " where TY_STD454 <> 0 or LY_STD454 <> 0 or LY_NXX <> 0"
        If Mid(RYW, 5, 2) <= "26" Then
            ASCMAIN1.sql &= " OR (LY_M01 <> 0 OR LY_M02 <> 0 OR LY_M03 <> 0 OR LY_M04 <> 0 OR LY_M05 <> 0 OR LY_M06 <> 0)"
        Else
            ASCMAIN1.sql &= " OR (LY_M07 <> 0 OR LY_M08 <> 0 OR LY_M09 <> 0 OR LY_M10 <> 0 OR LY_M11 <> 0 OR LY_M12 <> 0)"
        End If
        Create_TDA(dst.Tables.Add, "RSTSPRO6", "**", 0, False, "", 3)
        Fill_Records("RSTSPRO6")
        With dst.Tables("RSTSPRO6")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        For Each CUST_STORE_NO As String In CUST_STORE_NOs
            For Each T As String In New String() {"RSTSPRO2", "RSTSPRO6", "RSTSPRO7"}
                Dim RANK As Integer = 0
                Dim sqlCS As String = "CUST_CODE = '" & CUST_CODE & "' and CUST_STORE_NO = '" & CUST_STORE_NO & "'"
                If T = "RSTSPRO7" Then sqlCS = "CUST_CODE = '" & CUST_CODE & "'"
                For Each row As DataRow In dst.Tables(T).Select(sqlCS, "TY_STD454 DESC, LY_STD454 DESC")
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                    If rowICTITEM1 IsNot Nothing AndAlso rowICTITEM1.Item("ITEM_SNU_CODE") & "" = "S" AndAlso rowICTITEM1.Item("ITEM_CLASS_CODE") & "" <> "511" Then
                        ' Exclude class 511 (Bath & Body) as per DM/SP 05/11/22
                        'If rowICTITEM1 IsNot Nothing AndAlso rowICTITEM1.Item("ITEM_SNU_CODE") & "" = "S" Then
                        RANK += 1
                        row.Item("RANK") = RANK
                    End If
                Next
            Next
        Next


        ASCMAIN1.sql = "Select SPTCWRX2.*" & vbCrLf _
            & " from SPTCWRX2,SPTCWRXC,ICTCOLL1" & vbCrLf _
            & " where (SPTCWRX2.CUST_CODE,SPTCWRX2.CUST_STORE_NO) in " & vbCrLf _
            & " (Select Distinct CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & "   and SPTCWRX2.OPS_YYYYWW >= :PARM1 and SPTCWRX2.OPS_YYYYWW <= :PARM2" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SPTCWRX2.COLLECTION_CODE" & vbCrLf _
            & "   and SPTCWRXC.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and (SPTCWRXC.COLLECTION_GENDER = ICTCOLL1.COLLECTION_GENDER or SPTCWRXC.COLLECTION_GENDER = 'U')"
        Create_TDA(dst.Tables.Add, "SPTCWRX2", "**", 0, False, "VV", 0)

        Fill_Records("SPTCWRX2", New String() {ASCMAIN1.Week_Calc(RYW, -25), RYW})

    End Sub

    Sub Create_Worktable()
        ASCMAIN1.sql = "Select * from ARTCUST2" & vbCrLf _
            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and CUST_STORE_NO in ('" & Join(CUST_STORE_NOs.ToArray, "','") & "')"
        ARTCUST2 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "Alter Table " & ARTCUST2 & " Add Primary Key (CUST_CODE, CUST_STORE_NO)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim sqlwCUST_CODE As String = " and CUST_CODE = '" & CUST_CODE & "'"
        Dim sqlw = sqlwCUST_CODE
        If CUST_STORE_NOs.Count = 1 Then
            sqlw &= " and CUST_STORE_NO = '" & CUST_STORE_NOs(0) & "'"
        Else
            sqlw &= " and CUST_STORE_NO in (Select CUST_STORE_NO from " & ARTCUST2 & ")"
        End If

        'ASCMAIN1.Progress("Now Loading Data")

        rowGLTPARM3 = LookUp("GLTPARM3", RYW)
        RYM = rowGLTPARM3.Item("YYYYMM")

        Dim YTD_M As Integer = Val(Mid(RYM, 5, 2)) - 1
        If YTD_M = 0 Then YTD_M = 12

        REL_WEEK = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        Dim LEGEND_WK As String = rowGLTPARM3.Item("LEGEND")
        WEEK_LEGEND = "Week Ending " & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yy") & ",  " & LEGEND_WK
        WEEK_LEGEND = "Week Ending " & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yy") & ",  " & Mid(LEGEND_WK, 10, 3) & " week " & Mid(LEGEND_WK, 14, 1)

        RYW_2LY = ASCMAIN1.Week_Calc(RYW, -52 * 2)
        If ASCMAIN1.CLIENT = "INT" Then
            RYW_2LY = Format(Val(Mid(RYW, 1, 4)) - 2, "0000") & Mid(RYW, 5, 2)
        End If

        RYM01 = IIf(Mid(RYM, 5, 2) = "01",
                                  Format(Val(Mid(RYM, 1, 4)) - 1, "0000") & "02",
                                  Mid(RYM, 1, 4) & "02")



        Dim WKS110(110) As String
        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYWW between '" & RYW_2LY & "' and '" & RYW & "'"
        Dim W110 As Integer = -1
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "YYYYWW DESC")
            W110 += 1
            WKS110(W110) = row.Item("YYYYWW")
        Next

        Dim YPs(12, 1) As String
        For iYP As Integer = 1 To 12
            YPs(iYP, 0) = ASCMAIN1.Period_Calc(RYM01, iYP - 1)
            YPs(iYP, 1) = ASCMAIN1.Period_Calc(YPs(iYP, 0), -12)
        Next

        'Dim WKS(10, 1) As String
        ' 1-6 are for weeks
        ' 7 = HTD445
        ' 8 = STD454
        ' 9 = YTD445
        ' 10 = YTD454

        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & RYM & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "YYYYWW")
            Dim RW As Integer = Val(row.Item("REL_WEEK") & "")
            Dim YW As String = row.Item("YYYYWW")
            WKS(RW, 0) = YW
            WKS(RW, 1) = ASCMAIN1.Week_Calc(YW, -52)
            If ASCMAIN1.CLIENT = "INT" Then
                WKS(RW, 1) = Format(Val(Mid(YW, 1, 4)) - 1, "0000") & Mid(YW, 5, 2)
            End If
        Next

        If Mid(RYM, 5, 2) >= "07" Then
            WKS(7, 0) = Mid(RYW, 1, 4) & "23"  'HTD445
        Else
            WKS(7, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -4) 'HTD445
        End If
        If Mid(RYW, 5, 2) >= "27" Then
            WKS(8, 0) = Mid(RYW, 1, 4) & "27" 'STD454
        Else
            WKS(8, 0) = Mid(RYW, 1, 4) & "01" 'STD454
        End If

        If Mid(RYM, 5, 2) = "01" Then
            WKS(9, 0) = ASCMAIN1.Week_Calc(RYW, 1 - REL_WEEK) 'YTD445
        Else
            WKS(9, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -4) 'YTD445
        End If

        WKS(10, 0) = Mid(RYW, 1, 4) & "01" 'YTD454

        For W As Integer = 7 To 10
            WKS(W, 1) = ASCMAIN1.Week_Calc(WKS(W, 0), -52)
            If ASCMAIN1.CLIENT = "INT" Then
                WKS(W, 1) = Format(Val(Mid(WKS(W, 0), 1, 4)) - 1, "0000") & Mid(WKS(W, 0), 5, 2)
            End If
        Next


        Dim YWMIN As String = ""
        Dim YWMAX As String = ""
        For W As Integer = 1 To 10
            For Y As Integer = 0 To 1
                If WKS(W, Y) <> "" Then
                    If YWMIN = "" Or WKS(W, Y) < YWMIN Then YWMIN = WKS(W, Y)
                    If YWMAX = "" Or WKS(W, Y) > YWMAX Then YWMAX = WKS(W, Y)
                End If
            Next
        Next
        YWMIN = Format(Val(Mid(YWMIN, 1, 4)) - 1, "0000") & "01" ' some columns might be prior to min week - like LY_M01


        Dim XX As Integer = 10
        Dim D As String = "AMT_SOLD / 1000"

        Dim M12TYLY As String = ""
        For Each XY As String In New String() {"TY", "LY"}
            Dim Y As Integer = IIf(XY = "TY", 0, 1)
            For M As Integer = 1 To 12
                M12TYLY &= ", Sum (Decode (OPS_YYYYPP,'" & YPs(M, Y) & "'," & D & ",0)) " & XY & "_M" & Format(M, "00") & vbCrLf
            Next
            M12TYLY &= ", Sum (Case when OPS_YYYYPP between '" & YPs(1, Y) & "' and '" & YPs(YTD_M, Y) & "' then " & D & " else 0 End) " & XY & "_YTD" & vbCrLf
        Next

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(1, 0) & "'," & D & ",0)) TY_WK1" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(1, 1) & "'," & D & ",0)) LY_WK1" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(2, 0) & "'," & D & ",0)) TY_WK2" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(2, 1) & "'," & D & ",0)) LY_WK2" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(3, 0) & "'," & D & ",0)) TY_WK3" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(3, 1) & "'," & D & ",0)) LY_WK3" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(4, 0) & "'," & D & ",0)) TY_WK4" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(4, 1) & "'," & D & ",0)) LY_WK4" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(5, 0) & "'," & D & ",0)) TY_WK5" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(5, 1) & "'," & D & ",0)) LY_WK5" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(6, 0) & "'," & D & ",0)) TY_WK6" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(6, 1) & "'," & D & ",0)) LY_WK6" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(1, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_MTD" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(1, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_MTD" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(7, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_HTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(7, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_HTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(8, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_STD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(8, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_STD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(9, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_YTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(9, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_YTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(10, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_YTD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(10, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_YTD454" & vbCrLf _
            & M12TYLY & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX - 1) & "' and '" & WKS110(0) & "' THEN " & D & " ELSE 0 END) TY_LXX" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX + 52 - 1) & "' and '" & WKS110(1 + 52 - 1) & "' THEN " & D & " ELSE 0 END) LY_LXX" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX + 52 - 10 - 1) & "' and '" & WKS110(1 + 52 - 10 - 1) & "' THEN " & D & " ELSE 0 END) LY_NXX" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(REL_WEEK, 0) & "',QTY_EOW,0)) TYTW_ONH" & vbCrLf _
            & " from RSTRETL1" & vbCrLf _
            & " where OPS_YYYYWW between '" & YWMIN & "' and '" & YWMAX & "'" & vbCrLf _
            & sqlw _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE"

        '& " where OPS_YYYYWW between '" & WKS(9, 1) & "' and '" & WKS(REL_WEEK, 0) & "'" & vbCrLf _
        Dim sqlRSTSPRO1 As String = ASCMAIN1.sql
        RSTSPRO1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "Alter Table " & RSTSPRO1 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = Replace(sqlRSTSPRO1, D, "QTY_SOLD")
        RSTSPRO5 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "Alter Table " & RSTSPRO5 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = Replace(Replace(Replace(sqlRSTSPRO1, D, "QTY_SOLD"), sqlw, sqlwCUST_CODE), "CUST_CODE, CUST_STORE_NO", "CUST_CODE")
        RSTSPRO7 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "Alter Table " & RSTSPRO7 & " Add Primary Key (CUST_CODE, ITEM_CODE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim sqlCs As String = "Select HC_CODE from SATAUTH1 " & vbCrLf _
            & " where OPS_YYYYPP_OPENED IS NOT NULL" & vbCrLf _
            & "   and (SATAUTH1.OPS_YYYYPP_CLOSED IS NULL OR SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _
            & sqlw

        Dim sqlD As String = "Delete from " & RSTSPRO1 & " where ITEM_CODE in (" & vbCrLf _
            & " Select X.ITEM_CODE from " & RSTSPRO1 & " X, ICTITEM1, ICTCOLL1, ICTBRAN1" & vbCrLf _
            & "  where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "    and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "    and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "    and (NVL(ICTBRAN1.BRAND_STATUS,'?') <> 'A' or ICTCOLL1.HC_CODE Not in (" & sqlCs & ")))"
        ASCDATA1.ExecuteSQL(sqlD)
        ASCDATA1.ExecuteSQL(Replace(sqlD, RSTSPRO1, RSTSPRO5))
        ASCDATA1.ExecuteSQL(Replace(sqlD, RSTSPRO1, RSTSPRO7))

        Dim sqlSum As String = "" _
            & ", Sum(X.TY_WK1) TY_WK1, Sum(X.TY_WK2) TY_WK2, Sum(X.TY_WK3) TY_WK3, Sum(X.TY_WK4) TY_WK4, Sum(X.TY_WK5) TY_WK5, Sum(X.TY_WK6) TY_WK6" & vbCrLf _
            & ", Sum(X.LY_WK1) LY_WK1, Sum(X.LY_WK2) LY_WK2, Sum(X.LY_WK3) LY_WK3, Sum(X.LY_WK4) LY_WK4, Sum(X.LY_WK5) LY_WK5, Sum(X.LY_WK6) LY_WK6" & vbCrLf _
            & ", Sum(X.TY_MTD) TY_MTD, Sum(X.TY_HTD445) TY_HTD445, Sum(X.TY_STD454) TY_STD454, Sum(X.TY_YTD445) TY_YTD445, Sum(X.TY_YTD454) TY_YTD454" & vbCrLf _
            & ", Sum(X.LY_MTD) LY_MTD, Sum(X.LY_HTD445) LY_HTD445, Sum(X.LY_STD454) LY_STD454, Sum(X.LY_YTD445) LY_YTD445, Sum(X.LY_YTD454) LY_YTD454" & vbCrLf _
            & ", Sum(X.TY_M01) TY_M01, Sum(X.TY_M02) TY_M02, Sum(X.TY_M03) TY_M03, Sum(X.TY_M04) TY_M04, Sum(X.TY_M05) TY_M05, Sum(X.TY_M06) TY_M06" & vbCrLf _
            & ", Sum(X.TY_M07) TY_M07, Sum(X.TY_M08) TY_M08, Sum(X.TY_M09) TY_M09, Sum(X.TY_M10) TY_M10, Sum(X.TY_M11) TY_M11, Sum(X.TY_M12) TY_M12" & vbCrLf _
            & ", Sum(X.TY_YTD) TY_YTD" & vbCrLf _
            & ", Sum(X.LY_M01) LY_M01, Sum(X.LY_M02) LY_M02, Sum(X.LY_M03) LY_M03, Sum(X.LY_M04) LY_M04, Sum(X.LY_M05) LY_M05, Sum(X.LY_M06) LY_M06" & vbCrLf _
            & ", Sum(X.LY_M07) LY_M07, Sum(X.LY_M08) LY_M08, Sum(X.LY_M09) LY_M09, Sum(X.LY_M10) LY_M10, Sum(X.LY_M11) LY_M11, Sum(X.LY_M12) LY_M12" & vbCrLf _
            & ", Sum(X.LY_YTD) LY_YTD" & vbCrLf _
            & ", Sum(X.TY_LXX) TY_LXX, Sum(X.LY_LXX) LY_LXX, Sum(X.LY_NXX) LY_NXX" & vbCrLf _
            & ", Sum(X.TYTW_ONH * ICTITEM1.ITEM_RETAIL_PRICE / 1000) TYTW_ONH" & vbCrLf


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & sqlSum _
            & " from " & RSTSPRO1 & " X" & vbCrLf _
            & ",ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE"
        RSTSPRO2 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO2 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & Replace(sqlSum, " * ICTITEM1.ITEM_RETAIL_PRICE / 1000", "") _
            & " from " & RSTSPRO5 & " X" & vbCrLf _
            & ",ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE"
        RSTSPRO6 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  CUST_CODE" & vbCrLf _
            & ", CUST_STORE_NO" & vbCrLf _
            & ", COLLECTION_GENDER" & vbCrLf _
            & ", BRAND_CODE" & vbCrLf _
            & ", HC_CODE" & vbCrLf _
            & Replace(sqlSum, " * ICTITEM1.ITEM_RETAIL_PRICE / 1000", "") _
            & " from " & RSTSPRO2 & " X" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, COLLECTION_GENDER, BRAND_CODE, HC_CODE"
        RSTSPRO3 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", SPTCWRXC.CHECKBOOK" & vbCrLf _
            & ", X.BRAND_CODE" & vbCrLf _
            & Replace(sqlSum, " * ICTITEM1.ITEM_RETAIL_PRICE / 1000", "") _
            & " from " & RSTSPRO2 & " X, SPTCWRXC" & vbCrLf _
            & " where SPTCWRXC.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and (SPTCWRXC.COLLECTION_GENDER = X.COLLECTION_GENDER or SPTCWRXC.COLLECTION_GENDER = 'U')" & vbCrLf _
            & " group by X.CUST_CODE, X.CUST_STORE_NO, SPTCWRXC.CHECKBOOK, X.BRAND_CODE"
        RSTSPRO8 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (INV_DATE) INV_DATE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", 0 ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", NULL INV_DATE" & vbCrLf _
            & " from SOTORDR2" & vbCrLf _
            & " where ORDR_STATUS between 'O' and 'P'" & vbCrLf _
            & "   and (CUST_CODE, CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", 0 ORDR_QTY_PICK" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (SOTINVH1.INV_DATE) INV_DATE" & vbCrLf _
            & " from SOTINVH2,SOTINVH1,ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and ((ICTITEM1.ITEM_SNU_CODE = 'S' and SOTINVH2.OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -3) & "' and SOTINVH2.OPS_YYYYWW <= '" & RYW & "')" & vbCrLf _
            & "    or  (ICTITEM1.ITEM_SNU_CODE <> 'S' and SOTINVH2.OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -8) & "' and SOTINVH2.OPS_YYYYWW <= '" & RYW & "'))" & vbCrLf _
            & "   and (SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ")  group by CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        RSTSPRO4 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO4 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO4 & " Add ORDR_QTY_SHIP_INV_DATE NUMBER (8,0)")
        ASCMAIN1.sql = "" _
            & " Begin " & vbCrLf _
            & "  Declare Cursor C1 is " & vbCrLf _
            & "   Select * from " & RSTSPRO4 & " where INV_DATE is NOT Null for Update;" & vbCrLf _
            & "  Begin For R1 in C1 Loop" & vbCrLf _
            & "    Update " & RSTSPRO4 & " Set ORDR_QTY_SHIP_INV_DATE =" & vbCrLf _
            & "     (Select Sum (ORDR_QTY_SHIP) from SOTINVH2,SOTINVH1" & vbCrLf _
            & "       where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "         and SOTINVH1.CUST_CODE = R1.CUST_CODE and SOTINVH1.CUST_STORE_NO = R1.CUST_STORE_NO" & vbCrLf _
            & "         and SOTINVH1.INV_DATE = R1.INV_DATE" & vbCrLf _
            & "         and SOTINVH2.ITEM_CODE = R1.ITEM_CODE)" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE, ORDR_CUST_PO, ORDR_SHIP_DATE" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (INV_DATE) INV_DATE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, NVL(SOTORDR1.ORDR_ARRIVAL_DATE,NVL(ORDR_ORIG_SHIP_DATE,ORDR_SHIP_DATE)) ORDR_SHIP_DATE" & vbCrLf _
            & ", 0 ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", NULL INV_DATE" & vbCrLf _
            & " from SOTORDR2,SOTORDR1" & vbCrLf _
            & " where SOTORDR2.ORDR_STATUS between 'O' and 'P'" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and (SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & " group by SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, NVL(SOTORDR1.ORDR_ARRIVAL_DATE,NVL(ORDR_ORIG_SHIP_DATE,ORDR_SHIP_DATE))" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, NVL(SOTORDR1.ORDR_ARRIVAL_DATE,NVL(ORDR_ORIG_SHIP_DATE,ORDR_SHIP_DATE)) ORDR_SHIP_DATE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", 0 ORDR_QTY_PICK" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (SOTINVH1.INV_DATE) INV_DATE" & vbCrLf _
            & " from SOTINVH2,SOTINVH1,ICTITEM1,SOTORDR1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and ((ICTITEM1.ITEM_SNU_CODE = 'S' and SOTINVH2.OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -8) & "' and SOTINVH2.OPS_YYYYWW <= '" & RYW & "')" & vbCrLf _
            & "    or  (ICTITEM1.ITEM_SNU_CODE <> 'S' and SOTINVH2.OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -8) & "' and SOTINVH2.OPS_YYYYWW <= '" & RYW & "'))" & vbCrLf _
            & "   and (SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, NVL(SOTORDR1.ORDR_ARRIVAL_DATE,NVL(ORDR_ORIG_SHIP_DATE,ORDR_SHIP_DATE))" & vbCrLf _
            & ")  group by CUST_CODE, CUST_STORE_NO, ITEM_CODE, ORDR_CUST_PO, ORDR_SHIP_DATE"
        RSTSPROE = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPROE & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE, ORDR_CUST_PO, ORDR_SHIP_DATE)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPROE & " Add ORDR_QTY_SHIP_INV_DATE NUMBER (8,0)")
        ASCMAIN1.sql = "" _
            & " Begin " & vbCrLf _
            & "  Declare Cursor C1 is " & vbCrLf _
            & "   Select * from " & RSTSPROE & " where INV_DATE is NOT Null for Update;" & vbCrLf _
            & "  Begin For R1 in C1 Loop" & vbCrLf _
            & "    Update " & RSTSPROE & " Set ORDR_QTY_SHIP_INV_DATE =" & vbCrLf _
            & "     (Select Sum (ORDR_QTY_SHIP) from SOTINVH2,SOTINVH1,SOTORDR1" & vbCrLf _
            & "       where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "         and SOTINVH1.CUST_CODE = R1.CUST_CODE and SOTINVH1.CUST_STORE_NO = R1.CUST_STORE_NO" & vbCrLf _
            & "         and SOTINVH1.INV_DATE = R1.INV_DATE" & vbCrLf _
            & "         and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "         and SOTORDR1.ORDR_CUST_PO = R1.ORDR_CUST_PO" & vbCrLf _
            & "         and NVL(SOTORDR1.ORDR_ARRIVAL_DATE,NVL(ORDR_ORIG_SHIP_DATE,ORDR_SHIP_DATE)) = R1.ORDR_SHIP_DATE" & vbCrLf _
            & "         and SOTINVH2.ITEM_CODE = R1.ITEM_CODE)" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_SHIP NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_PICK NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_OPEN NUMBER (8,0)")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & RSTSPRO4 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & RSTSPRO6 & vbCrLf _
            & "    Set ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
            & "       ,ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
            & "       ,ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE" & vbCrLf _
            & "      and CUST_STORE_NO = R1.CUST_STORE_NO" & vbCrLf _
            & "      and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
 
        ASCMAIN1.sql = "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "RSTRETL1.OPS_YYYYWW, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_GENDER, SUM (RSTRETL1.AMT_SOLD) AMT_SOLD" & vbCrLf _
            & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and RSTRETL1.CUST_STORE_NO in ('" & Join(CUST_STORE_NOs.ToArray, "','") & "')" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and NVL(RSTRETL1.AMT_SOLD,0) <> 0" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and OPS_YYYYWW between '" & RYW_2ly & "' and '" & RYW & "'" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "RSTRETL1.OPS_YYYYWW, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_GENDER"
        RSTSPROB = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "Alter Table " & RSTSPROB & " Add CHECKBOOK VARCHAR2(6)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from SPTCWRXC;" & vbCrLf _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update " & RSTSPROB & " Set CHECKBOOK = R1.CHECKBOOK" & vbCrLf _
            & "  where BRAND_CODE = R1.BRAND_CODE" & vbCrLf _
            & "    and (COLLECTION_GENDER = R1.COLLECTION_GENDER or R1.COLLECTION_GENDER = 'U');" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, OPS_YYYYWW, CHECKBOOK" & vbCrLf _
            & ", SUM (AMT_SOLD) TY_SLS from " & RSTSPROB & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, OPS_YYYYWW, CHECKBOOK"
        RSTSPROC = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCMAIN1.sql = "Alter Table " & RSTSPROC & " Add LY_SLS NUMBER (13,2)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        For W As Integer = 0 To 25

            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -1 * W)
            Dim YW_L As String = ASCMAIN1.Week_Calc(YW, -52)
            If ASCMAIN1.CLIENT = "INT" Then
                YW_L = Format(Val(Mid(YW, 1, 4)) - 1, "0000") & Mid(YW, 5, 2)
            End If
            ASCMAIN1.sql = "Update " & RSTSPROC & " X" & vbCrLf _
                & " Set LY_SLS = (Select TY_SLS from " & RSTSPROC & vbCrLf _
                & " where CUST_CODE = X.CUST_CODE" & vbCrLf _
                & "   and CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
                & "   and NVL(CHECKBOOK,'?') = NVL(X.CHECKBOOK,'?')" & vbCrLf _
                & "   and OPS_YYYYWW = '" & YW_L & "')" & vbCrLf _
                & " where OPS_YYYYWW = '" & YW & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Next
    End Sub
End Class