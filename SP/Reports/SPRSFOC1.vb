Public Class SPRSFOC1

#Region "General Declarations"
    Dim SEASON_CODE As String

    Dim SPTSFOC9 As String
    Dim SPTCOOPA As String
    Dim SPTSFOC1_NO_CS As String

    Dim rowICTSEAS1 As DataRow
    Dim SEASON_YEAR As String
    Dim SEASON_YEAR_LY As String
    Dim SEASON_TYPE As String
    Dim SEASON_DESC As String
    Dim SEASON_DESC_LY As String

    Dim SELL_CODE As String
    Dim SELL_TYPE As String
    Dim sqlSELL_CODE As String

    Dim REGION_CODE As String

    Dim RUNAS As String = ""

    Dim YPs(,) As String
    Dim Weeks_TY As New Dictionary(Of String, Integer)
    Dim RAs As New Dictionary(Of String, Integer)
    Dim RLs As New Dictionary(Of String, Integer)
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim XLS_FILENAME_base As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Dim YM As String = ASCMAIN1.CYM
        YM = ASCMAIN1.Period_Calc(YM, 1)
        Absx1.txtFor("SEASON_CODE").Text = Mid(YM, 1, 4) & IIf(Val(Mid(YM, 5, 2)) < 8, "S", "F")

        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP0", YM, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, ByRef Optional sql_where As String = "", ByRef Optional Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)
        If COLUMN_NAME = "SEASON_CODE" Then
            Dim LY As String = Format(Now.Date.Year - 1, "0000")
            Dim NY As String = Format(Now.Date.Year + 1, "0000")
            sql_where = "SEASON_YEAR >= '" & LY & "' and SEASON_YEAR <= '" & NY & "'"
        End If

    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Work File")

        Dim sqlw As String = ""
        'sqlw &= "   and SPTCOOP1.OPS_YYYYPP <= '" & RYP0 & "'"

        SEASON_CODE = Absx1.txtFor("SEASON_CODE").Text
        SELL_CODE = cmbSELL_CODE.Value
        REGION_CODE = cmbREGION_CODE.Value

        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE)
        If rowSOTSELL1 IsNot Nothing Then
            SELL_TYPE = rowSOTSELL1.Item("SELL_TYPE") & ""
        End If

        If SELL_TYPE = "AC" Then
            sqlSELL_CODE = "ARTCUST2.SELL_CODE_AC"
        Else
            sqlSELL_CODE = "ARTCUST2.SELL_CODE"
        End If

        Prepare_dst(True, New Object() {SEASON_CODE, RYP0, RYP1})
        '   Check_if_Empty("SPTSFOC1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Sub Print_Report()

        Dim sqlARTCUST2 As String = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
            & ", " & sqlSELL_CODE & " SELL_CODE, ARTCUST2.SDS_CODE" & vbCrLf _
            & " from ARTCUST2,ARTCUST1,SOTTCLS1,SOTSELL1" & vbCrLf _
            & " where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   And ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
            & "   And SOTSELL1.SELL_CODE = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   And SOTTCLS1.CHANNEL_CODE = '1'"

        XLS_FILENAME_base = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No(Me.Name & "_SFC") & "SFC"

        If RUNAS = "T" Then

            ASCMAIN1.sql = sqlARTCUST2 & " and " & sqlSELL_CODE & " = '" & SELL_CODE & "'"

            Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
            Create_XLS("T", SELL_CODE)

            ASCMAIN1.sql = "Select Distinct SDS_CODE from ARTCUST2 where " & sqlSELL_CODE & " = '" & SELL_CODE & "' and SDS_CODE is Not Null"
            For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("", "SDS_CODE")
                Dim SDS_CODE As String = ROW.Item("SDS_CODE")

                ASCMAIN1.sql = sqlARTCUST2 & " and " & sqlSELL_CODE & " = '" & SELL_CODE & "' and SDS_CODE = '" & SDS_CODE & "'"
                Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
                Create_XLS("S", SDS_CODE)
            Next

        ElseIf RUNAS = "R" Then

            ASCMAIN1.sql = sqlARTCUST2 & " and SOTSELL1.REGION_CODE = '" & REGION_CODE & "'"
            Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
            Create_XLS("R", REGION_CODE)

            ASCMAIN1.sql = "Select Distinct SELL_CODE from SOTSELL1 where SOTSELL1.REGION_CODE = '" & REGION_CODE & "'"
            For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("", "SELL_CODE")
                Dim SELL_CODE As String = ROW.Item("SELL_CODE")

                ASCMAIN1.sql = sqlARTCUST2 & " and ARTCUST2.SELL_CODE = '" & SELL_CODE & "'"
                Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
                'Dim rows() As DataRow = dst.Tables("ARTCUST2").Select("CUST_CODE <> 'IPLBAE'")
                If dst.Tables("ARTCUST2").Rows.Count > 0 Then
                    Create_XLS("T", SELL_CODE)
                End If
            Next

        ElseIf RUNAS = "C" Then
            dst.Tables("ARTCUST2").Rows.Clear()
            Create_XLS()
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            Set_YPs()
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
            Dim rowICTSEAS1 As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
            If rowICTSEAS1 Is Nothing Then
                EMsg &= vbCrLf & "Invalid Season"
            Else
                Dim SEASON_YEAR As String = rowICTSEAS1.Item("SEASON_YEAR") & ""
                If SEASON_YEAR <> Mid(ASCMAIN1.CYM, 1, 4) And
                   SEASON_YEAR <> Format(Val(Mid(ASCMAIN1.CYM, 1, 4)) - 1, "0000") And
                   SEASON_YEAR <> Format(Val(Mid(ASCMAIN1.CYM, 1, 4)) + 1, "0000") Then
                    EMsg &= vbCrLf & "Invalid Season Year"
                End If
                Dim SEASON_TYPE As String = rowICTSEAS1.Item("SEASON_TYPE") & ""
                If SEASON_TYPE <> "S" And SEASON_TYPE <> "F" Then
                    EMsg &= vbCrLf & "Invalid Season Type"
                End If

                Dim YPX As String = ""
                YPX = Absx1.cmbFor("RYP0").Value & ""
                YPX = Mid(YPX, 1, 4) & Mid(YPX, 6, 2)
                If YPX < SEASON_YEAR & IIf(SEASON_TYPE = "S", "01", "07") Or YPX > SEASON_YEAR & IIf(SEASON_TYPE = "S", "06", "12") Then
                    EMsg &= vbCrLf & "Start Month outside Season Range"
                End If
                YPX = Absx1.cmbFor("RYP1").Value & ""
                YPX = Mid(YPX, 1, 4) & Mid(YPX, 6, 2)
                If YPX < SEASON_YEAR & IIf(SEASON_TYPE = "S", "01", "07") Or YPX > SEASON_YEAR & IIf(SEASON_TYPE = "S", "06", "12") Then
                    EMsg &= vbCrLf & "End Month outside Season Range"
                End If

                If Absx1.cmbFor("RYP1").Value & "" < Absx1.cmbFor("RYP0").Value & "" Then
                    EMsg &= vbCrLf & "Period Range out of sequence"
                End If
            End If


            If optRUNAS.Value = "T" Then
                If cmbSELL_CODE.Value & "" = "" Then
                    EMsg &= vbCrLf & "You Must Select an AE"
                End If
            ElseIf optRUNAS.Value = "R" Then
                If cmbREGION_CODE.Value & "" = "" Then
                    EMsg &= vbCrLf & "You Must Select an ASD"
                End If
            End If

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        RUNAS = optRUNAS.Value
        Create_Work_Tables(True)

        With dst

            ASCMAIN1.sql = "Select * from " & SPTSFOC1_NO_CS & " SPTSFOC1_NO_CS"
            Create_TDA(.Tables.Add, "SPTSFOC1_NO_CS", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & SPTSFOC9 & " SPTSFOC9"
            Create_TDA(.Tables.Add, "SPTSFOC9", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select * from SPTSFOC1 where EVENT_GROUP_NO in (Select Distinct EVENT_GROUP_NO from " & SPTSFOC9 & " union Select Distinct EVENT_GROUP_NO from " & SPTSFOC1_NO_CS & ")"
            Create_TDA(.Tables.Add, "SPTSFOC1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from SPTSFOC3 where EVENT_GROUP_NO in (Select Distinct EVENT_GROUP_NO from " & SPTSFOC9 & " union Select Distinct EVENT_GROUP_NO from " & SPTSFOC1_NO_CS & ")"
            Create_TDA(.Tables.Add, "SPTSFOC3", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SPTSFOC1.EVENT_GROUP_NO" & vbCrLf _
                & ", DECODE(SPTSFOC1.VEHICLE_CODE,'MA',SPTSFOC3.ITEM_CODE,NULL) ITEM_CODE" & vbCrLf _
                & ", MAX(ICTCOLL1.BRAND_CODE) BRAND_CODE" & vbCrLf _
                & ", MAX(ICTCOLL1.COLLECTION_NAME) COLLECTION_NAME" & vbCrLf _
                & ", DECODE(SPTSFOC1.VEHICLE_CODE,'MA',SPTSFOC3.FEATURE_DESC,NULL) FEATURE_DESC" & vbCrLf _
                & ", MAX(SPTSFOC1.EVENT_DATE_CHANGED) EVENT_DATE_CHANGED" & vbCrLf _
                & " from ICTCOLL1,SPTSFOC3,SPTSFOC1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = SPTSFOC3.COLLECTION_CODE" & vbCrLf _
                & "   And SPTSFOC1.EVENT_GROUP_NO = SPTSFOC3.EVENT_GROUP_NO" & vbCrLf _
                & "   And SPTSFOC1.STATUS_CODE = 'O'" & vbCrLf _
                & "   And SPTSFOC3.EVENT_GROUP_NO in (Select Distinct EVENT_GROUP_NO from " & SPTSFOC9 & " union Select Distinct EVENT_GROUP_NO from " & SPTSFOC1_NO_CS & ")" & vbCrLf _
                & " group by SPTSFOC1.EVENT_GROUP_NO" & vbCrLf _
                & ", DECODE(SPTSFOC1.VEHICLE_CODE,'MA',SPTSFOC3.ITEM_CODE,NULL)" & vbCrLf _
                & ", DECODE(SPTSFOC1.VEHICLE_CODE,'MA',SPTSFOC3.FEATURE_DESC,NULL)"
            Create_TDA(.Tables.Add, "SPTSFOCE", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select * from SPTENOT2 where SEASON_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SPTENOT2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM between :PARM1 and :PARM2"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "VV", 1)

            ASCMAIN1.sql = "Select SPTCOOP1.*" & vbCrLf _
                & " from SPTCOOP1," & SPTCOOPA & " SPTCOOPA" & vbCrLf _
                & " where SPTCOOP1.AUTH_NO = SPTCOOPA.AUTH_NO"
            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, False, "", 1)
            With .Tables("SPTCOOP1").Columns
                .Add("YYYYWW_END") ' YW that this event ends
            End With

            ASCMAIN1.sql = "Select SPTCOOPB.AUTH_NO, SPTCOOP1.CUST_CODE, SPTCOOPB.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, " & sqlSELL_CODE & " SELL_CODE, ARTCUST2.SDS_CODE, SOTSELL1.REGION_CODE" & vbCrLf _
                & " from ARTCUST2,SPTCOOPB,SPTCOOP1," & SPTCOOPA & " SPTCOOPA,SOTSELL1" & vbCrLf _
                & " where SPTCOOPB.AUTH_NO = SPTCOOPA.AUTH_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
                & "   And ARTCUST2.CUST_STORE_NO = SPTCOOPB.CUST_STORE_NO" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = " & sqlSELL_CODE & vbCrLf _
                & "   And SPTCOOP1.AUTH_NO = SPTCOOPB.AUTH_NO"
            Create_TDA(.Tables.Add, "SPTCOOPA", "**", 0, False, "", 0)


            ASCMAIN1.sql = "Select Distinct SPTCOOP3.AUTH_NO" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_NAME, SPTCOOP3.FEATURE_DESC" & vbCrLf _
                & " from SPTCOOP1,ICTCOLL1,SPTCOOP3," & SPTCOOPA & " SPTCOOPA" & vbCrLf _
                & " where SPTCOOP1.AUTH_NO = SPTCOOPA.AUTH_NO" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
                & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO"
            Create_TDA(.Tables.Add, "SPTCOOPE", "**", 0, False, "", 0)

            Create_TDA(.Tables.Add, "ICTBRAN1", "*", 0, False)
            Fill_Records("ICTBRAN1")

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 0, False)

            Create_TDA(.Tables.Add, "GLTPARM2", "*", 0, False)
            Fill_Records("GLTPARM2")
        End With

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        'If parms.Length > 0 Then
        '    sqlw = parms(0)
        'End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        SEASON_CODE = parms(0)
        RYP0 = parms(1)
        RYP1 = parms(2)

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        SEASON_YEAR = rowICTSEAS1.Item("SEASON_YEAR")
        SEASON_YEAR_LY = Format(Val(SEASON_YEAR) - 1, "0000")
        SEASON_TYPE = rowICTSEAS1.Item("SEASON_TYPE")
        SEASON_DESC = rowICTSEAS1.Item("SEASON_DESC")
        SEASON_DESC_LY = Replace(SEASON_DESC, SEASON_YEAR, SEASON_YEAR_LY)

        Dim YP As String = ""
        If SEASON_TYPE = "F" Then
            YP = SEASON_YEAR & "08"
        ElseIf SEASON_TYPE = "S" Then
            YP = SEASON_YEAR & "02"
        End If

        ' Get YP information into an array for the 6 months of the chosen season, TY & LY 

        ReDim YPs(6, 1) ' 0 = TY, 1 = LY
        For I As Integer = 1 To 6
            YPs(I, 0) = ASCMAIN1.Period_Calc(YP, I - 1)
            YPs(I, 1) = ASCMAIN1.Period_Calc(YP, I - 1 - 12)
        Next

        ' Pull all of the weeks for the current season into a datatable, and then determine the week which is 52 weeks prior

        'Fill_Records("GLTPARM3", New String() {YPs(1, 0), YPs(6, 0)})
        Fill_Records("GLTPARM3", New String() {RYP0, RYP1})

        Dim W As Integer = 0
        Weeks_TY.Clear()
        For Each rowGLTPARM3 As DataRow In dst.Tables("GLTPARM3").Select("", "YYYYWW")
            Dim YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
            W += 1
            Weeks_TY.Add(YYYYWW, W)
        Next

        ' Fill Work Table of Customer Stores

        Create_Work_Tables(False)

        EnforceConstraints(False)

        Fill_Records("SPTSFOC1_NO_CS")
        Fill_Records("SPTSFOC9")
        Fill_Records("SPTSFOC1")
        Fill_Records("SPTSFOC3")
        Fill_Records("SPTSFOCE")

        Fill_Records("SPTENOT2", SEASON_CODE)

        ' Pull all of the Promo Information into the workbook for events which start in any of the weeks in scope for the selected season, TY and LY
        ' Problem - what about an even which starts prior to the 1st week of TY, but ends sometime in the time period reflected by the SxS report

        'Fill_Records("SPTCOOP1", New String() {RYP0, RYP1}, True)
        'Fill_Records("SPTCOOPA", New String() {RYP0, RYP1}, True)
        'Fill_Records("SPTCOOPE", New String() {RYP0, RYP1}, True)
        Fill_Records("SPTCOOP1")
        Fill_Records("SPTCOOPA")
        Fill_Records("SPTCOOPE")

        EnforceConstraints(True)

    End Sub

    Sub Create_Work_Tables(initialize As Boolean)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        Dim sqlSPTSFOC9 As String = "Select SPTSFOC9.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", " & sqlSELL_CODE & " SELL_CODE, SOTSELL1.REGION_CODE, ARTCUST2.SDS_CODE" & vbCrLf _
                & " from SPTSFOC9,ARTCUST2,SPTSFOC1,SOTSELL1" & vbCrLf _
                & " where SPTSFOC1.EVENT_GROUP_NO = SPTSFOC9.EVENT_GROUP_NO" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = " & sqlSELL_CODE & vbCrLf _
                & "   and SPTSFOC1.STATUS_CODE = 'O'" & vbCrLf _
                & "   and SPTSFOC1.APPR_STATUS_CODE In ('A')" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTSFOC9.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTSFOC9.CUST_STORE_NO"

        Dim sqlSPTSFOC1_NO_CS As String = "Select SPTSFOC1.EVENT_GROUP_NO" & vbCrLf _
                & " from SPTSFOC1" & vbCrLf _
                & " where VEHICLE_CODE = 'BF'" & vbCrLf _
                & "   and OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
                & "   and STATUS_CODE = 'O'" & vbCrLf _
                & "   and SPTSFOC1.APPR_STATUS_CODE In ('A')" & vbCrLf _
                & " Minus Select Distinct EVENT_GROUP_NO from SPTSFOC9"

        '  & "   And SPTCOOP1.CUST_CODE In (Select Distinct CUST_CODE from " & SPTSFOC9 & ")" & vbCrLf _
        Dim sqlSPTCOOPA As String = "Select SPTCOOP1.AUTH_NO" & vbCrLf _
                & " from SPTCOOP1, GLTPARM3" & vbCrLf _
                & " where SPTCOOP1.APPR_STATUS_CODE In ('A')" & vbCrLf _
                & "   and SPTCOOP1.EXPENSE_TYPE_CODE in ('COOP','SCENT','NATMEDIA','RTLEVENTS','VISUAL')" & vbCrLf _
                & "   and SPTCOOP1.EVENT_GROUP_NO IS NULL" & vbCrLf _
                & "   and GLTPARM3.YYYYWW = SPTCOOP1.OPS_YYYYWW"
        '                & " where SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _

        If initialize Then
            ASCMAIN1.sql = sqlSPTSFOC9 & " and ROWNUM < 1"
            SPTSFOC9 = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTSFOC9 & " Add Primary Key (EVENT_GROUP_NO,CUST_CODE,CUST_STORE_NO)")

            ASCMAIN1.sql = sqlSPTSFOC1_NO_CS
            SPTSFOC1_NO_CS = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTSFOC1_NO_CS & " Add Primary Key (EVENT_GROUP_NO)")

            ASCMAIN1.sql = sqlSPTCOOPA & " and ROWNUM < 1"
            SPTCOOPA = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCOOPA & " Add Primary Key (AUTH_NO)")
        Else
            ASCMAIN1.sql = sqlSPTSFOC9
            Select Case RUNAS
                Case "T"
                    ASCMAIN1.sql &= " and " & sqlSELL_CODE & " = '" & SELL_CODE & "'"
                Case "R"
                    ASCMAIN1.sql &= " and SOTSELL1.REGION_CODE = '" & REGION_CODE & "'"
                Case "C"
                    ' ASCMAIN1.sql &= " and ROWNUM < 1"
                Case Else
                    ASCMAIN1.sql &= " and ROWNUM < 1"
            End Select
            'ASCMAIN1.sql &= " and OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"
            ASCMAIN1.sql &= " and SPTSFOC1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"
            ASCDATA1.ExecuteSQL("Truncate Table " & SPTSFOC9)
            ASCDATA1.ExecuteSQL("Insert into " & SPTSFOC9 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = sqlSPTSFOC1_NO_CS
            ASCDATA1.ExecuteSQL("Truncate Table " & SPTSFOC1_NO_CS)
            ASCDATA1.ExecuteSQL("Insert into " & SPTSFOC1_NO_CS & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = sqlSPTCOOPA
            'ASCMAIN1.sql &= " and OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"
            ASCMAIN1.sql &= " and GLTPARM3.YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"
            ASCDATA1.ExecuteSQL("Truncate Table " & SPTCOOPA)
            ASCDATA1.ExecuteSQL("Insert into " & SPTCOOPA & " " & ASCMAIN1.sql)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Create_XLS(Optional summary_type As String = "",
                   Optional summary_code As String = "")

        worksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim XLS_FILENAME As String = XLS_FILENAME_base
        XLS_FILENAME &= IIf(summary_type = "", "_Corporate",
                            IIf(summary_type = "R", "_ASD_" & summary_code,
                            IIf(summary_type = "T", "_" & SELL_TYPE & "_", "_SDS_") & summary_code))
        XLS_FILENAME &= ".xlsX"

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        ' Get List of All Brands and Associate a background color with each

        Dim ET_COLORs As New List(Of SpreadsheetGear.Color)
        ET_COLORs.Add(SpreadsheetGear.Colors.DarkOrange)
        ET_COLORs.Add(SpreadsheetGear.Colors.Red)
        ET_COLORs.Add(SpreadsheetGear.Colors.DeepPink)
        ET_COLORs.Add(SpreadsheetGear.Colors.Green)
        ET_COLORs.Add(SpreadsheetGear.Colors.Purple)

        Dim iBrand As Integer = 0
        Dim BRAND_CODEs As New Dictionary(Of String, SpreadsheetGear.Color)
        For Each row As DataRow In dst.Tables("ICTBRAN1").Select("", "BRAND_STATUS,BRAND_CODE")
            Dim BRAND_CODE As String = row.Item(0)
            iBrand += 1
            Dim BRAND_COLOR As Int64 = Val(row.Item("BRAND_COLOR") & "")
            If BRAND_COLOR = 0 Then
                BRAND_CODEs.Add(BRAND_CODE, ET_COLORs((iBrand - 1) Mod ET_COLORs.Count))
            Else
                BRAND_CODEs.Add(BRAND_CODE, SpreadsheetGear.Color.FromArgb(BRAND_COLOR))
            End If
        Next
        BRAND_CODEs.Add("", SpreadsheetGear.Colors.Black)

        Dim CUST_CODEs As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTCUST2"), New String() {"CUST_CODE"}).Select()
            Dim CUST_CODE As String = row.Item(0)
            CUST_CODEs.Add(CUST_CODE)
        Next


        ' Create a worksheet for each Customer/Store

        Dim Sheets As New List(Of String)
        Sheets.Add("*")

        If summary_type = "T" Or summary_type = "S" Then
            For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("", "CUST_CODE,CUST_STORE_NO")
                Dim CUST_CODE As String = rowARTCUST2.Item("CUST_CODE")
                Dim CUST_STORE_NO As String = rowARTCUST2.Item("CUST_STORE_NO")
                Sheets.Add(CUST_CODE & "-" & CUST_STORE_NO)
            Next
        End If

        For Each CS As String In Sheets
            Dim CUST_CODE As String = ""
            Dim CUST_STORE_NO As String = ""
            Dim CUST_STORE_NAME As String = ""

            RAs.Clear()
            RLs.Clear()

            Dim sqlCS As String = ""

            Dim rowARTCUST2 As DataRow = Nothing
            Dim SHEET_DESC As String = ""

            If CS = "*" Then
                worksheet = workbook.Worksheets(0)

                Select Case summary_type
                    Case "R"
                        worksheet.Name = "ASD " & summary_code
                        Dim rowSOTSREG1 As DataRow = LookUp("SOTSREG1", summary_code)
                        If rowSOTSREG1 Is Nothing Then
                            SHEET_DESC = worksheet.Name
                        Else
                            SHEET_DESC = rowSOTSREG1.Item("REGION_DESC") & ""
                        End If
                    Case "T"
                        worksheet.Name = "AE " & summary_code
                        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", summary_code)
                        If rowSOTSELL1 Is Nothing Then
                            SHEET_DESC = worksheet.Name
                        Else
                            worksheet.Name = rowSOTSELL1.Item("SELL_TYPE") & " " & summary_code
                            SHEET_DESC = rowSOTSELL1.Item("SELL_NAME")
                        End If
                    Case "S"
                        worksheet.Name = "SDS " & summary_code
                        Dim rowSOTSDSC1 As DataRow = LookUp("SOTSDSC1", summary_code)
                        If rowSOTSDSC1 Is Nothing Then
                            SHEET_DESC = worksheet.Name
                        Else
                            SHEET_DESC = rowSOTSDSC1.Item("SDS_NAME")
                        End If
                    Case ""
                        worksheet.Name = "Corporate"
                        SHEET_DESC = worksheet.Name
                    Case Else
                        worksheet.Name = "Summary"
                        SHEET_DESC = worksheet.Name
                End Select

            Else
                CUST_CODE = Split(CS, "-")(0)
                CUST_STORE_NO = Split(CS, "-")(1)

                sqlCS = "CUST_CODE = '" & CUST_CODE & "' and CUST_STORE_NO = '" & CUST_STORE_NO & "'"

                rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                CUST_STORE_NAME = rowARTCUST2.Item("CUST_STORE_NAME") & ""

                worksheet = workbook.Worksheets.Add
                worksheet.Name = CS
                SHEET_DESC = CUST_STORE_NAME
            End If

            ' Set up Column Counts for Zones

            Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
            Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

            Dim C_Offset As Integer = 2
            Dim R_Offset As Integer = 1
            Dim WRows As Integer = 12

            Dim XR As Integer = 0
            Dim XC As Integer = 0

            Dim RM As New Dictionary(Of String, Integer)

            Dim W As Integer = 0
            For Each rowGLTPARM3 As DataRow In dst.Tables("GLTPARM3").Select("", "YYYYWW")
                Dim YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
                Dim YYYYMM As String = rowGLTPARM3.Item("YYYYMM")
                Dim REL_WEEK As String = Val(rowGLTPARM3.Item("REL_WEEK") & "")
                Dim WEEK_END_DATE As Date = CDate(rowGLTPARM3.Item("WEEK_END_DATE"))

                Dim MM As Integer = Val(Mid(YYYYMM, 5, 2))

                Dim rowSPTENOT2 As DataRow = dst.Tables("SPTENOT2").Rows.Find(New String() {SEASON_CODE, YYYYWW})
                W += 1
                Dim H As String = rowGLTPARM3.Item("LEGEND")

                XC = C_Offset + REL_WEEK

                If Not RM.ContainsKey(YYYYMM) Then
                    Dim M As Integer = RM.Keys.Count
                    RM.Add(YYYYMM, M * WRows + R_Offset)
                    If M > 0 Then worksheet.Cells(RM(YYYYMM) - 1, 0).PageBreak = SpreadsheetGear.PageBreak.Manual
                    worksheet.Cells(RM(YYYYMM) - 1, 2).Value = SHEET_DESC
                End If
                XR = RM(YYYYMM)

                worksheet.Cells(XR - 1, XC).Value = Mid(H, 10, 7)
                worksheet.Cells(XR - 0, XC).Value = Format(WEEK_END_DATE.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE, "MM/dd")
                worksheet.Cells(XR, XC).EntireColumn.ColumnWidth = 40
                worksheet.Cells(XR, XC).WrapText = True

                range = worksheet.Cells(XR - 1, XC, XR - 0, XC)
                With range
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    If MM Mod 2 = 0 Then
                        .Interior.Color = SpreadsheetGear.Colors.LightGray ' .Beige '.Orange
                    Else
                        .Interior.Color = SpreadsheetGear.Colors.LightGray ' Lavender '.Yellow
                    End If
                End With

                If rowSPTENOT2 IsNot Nothing Then
                    worksheet.Cells(XR + 2, XC).Value = rowSPTENOT2.Item("WEEKLY_NOTE_ADDL")
                    worksheet.Cells(XR + 5, XC).Value = rowSPTENOT2.Item("WEEKLY_NOTE_SAMP")
                End If
            Next

            worksheet.Cells(R_Offset + 0, C_Offset + 0).EntireColumn.ColumnWidth = 25
            XC = C_Offset + 0
            For Each YYYYMM As String In RM.Keys
                XR = RM(YYYYMM)
                For iRow As Integer = 1 To 7
                    worksheet.Cells(XR + iRow, XC).Value = New String() {
                        "VISUAL WEEKS (V)",
                        "ADDITIONAL COMMENTS",
                        "MASTER EVENT (Tracked)",
                        "BRAND FOCUS EVENT",
                        "NATIONAL SAMPLING",
                        "COOP/SCENT/NATMEDIA",
                        "RETAILER PROGRAMS"}(iRow - 1)
                    worksheet.Cells(XR + iRow, XC).EntireRow.VerticalAlignment = SpreadsheetGear.VAlign.Top
                    worksheet.Cells(XR + iRow, XC).EntireRow.WrapText = True
                Next

                Dim Weeks_in_Month As Integer = dst.Tables("GLTPARM3").Select("YYYYMM='" & YYYYMM & "'").Length
                With worksheet.Cells(XR, C_Offset, XR + 7, C_Offset + Weeks_in_Month)
                    .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With
            Next


            ' Event Groups (Store Focus)

            Dim rowEGs() As DataRow = Nothing
            If CS = "*" Then
                If summary_type = "S" Then
                    Dim sqlSDS As String = "SDS_CODE = '" & summary_code & "'"
                    rowEGs = dst.Tables("SPTSFOC9").Select(sqlSDS, "")
                Else
                    rowEGs = dst.Tables("SPTSFOC1").Select("", "OPS_YYYYWW")
                End If
            Else
                rowEGs = dst.Tables("SPTSFOC9").Select(sqlCS, "")
            End If

            Dim EVENT_GROUP_NOs As New List(Of String)
            For Each rowEG As DataRow In rowEGs
                Dim EVENT_GROUP_NO As String = rowEG.Item("EVENT_GROUP_NO")
                '                EVENT_GROUP_NOs.Add(EVENT_GROUP_NO)
                If Not EVENT_GROUP_NOs.Contains(EVENT_GROUP_NO) Then
                    EVENT_GROUP_NOs.Add(EVENT_GROUP_NO)
                End If
            Next

            ' ADD BF EVENTS WITH NO STORES
            For Each row As DataRow In dst.Tables("SPTSFOC1_NO_CS").Select
                Dim EVENT_GROUP_NO As String = row.Item("EVENT_GROUP_NO")
                If Not EVENT_GROUP_NOs.Contains(EVENT_GROUP_NO) Then
                    EVENT_GROUP_NOs.Add(EVENT_GROUP_NO)
                End If
            Next

            For Each EVENT_GROUP_NO As String In EVENT_GROUP_NOs
                Dim rowSPTSFOC1 As DataRow = dst.Tables("SPTSFOC1").Rows.Find(EVENT_GROUP_NO)
                Dim YYYYWW As String = rowSPTSFOC1.Item("OPS_YYYYWW")
                Dim VEHICLE_CODE As String = rowSPTSFOC1.Item("VEHICLE_CODE")
                Dim BOOKING_NAME As String = rowSPTSFOC1.Item("BOOKING_NAME")
                Dim EXPENSE_TYPE_CODE As String = rowSPTSFOC1.Item("EXPENSE_TYPE_CODE")
                Dim NOTES As String = rowSPTSFOC1.Item("NOTES") & ""
                Dim EVENT_FILE_LINK As String = rowSPTSFOC1.Item("EVENT_FILE_LINK") & ""
                Dim EVENT_DATE_CHANGED As Date = rowSPTSFOC1.Item("EVENT_DATE_CHANGED")

                If Weeks_TY.ContainsKey(YYYYWW) Then
                    W = Weeks_TY(YYYYWW)

                    Dim RX As Integer = 0
                    If VEHICLE_CODE = "MA" Then RX = 3
                    If VEHICLE_CODE = "BF" Then RX = 4

                    If RX <> 0 Then
                        Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
                        Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")
                        Dim YYYYMM As String = rowGLTPARM3.Item("YYYYMM") & ""

                        Dim CSs As String = ""

                        If summary_type = "T" Or summary_type = "S" Or summary_type = "R" Then
                            Dim sqlw As String = "EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                            If CS <> "*" Then sqlw &= " and " & sqlCS
                            If summary_type = "T" Then sqlw &= " and SELL_CODE = '" & summary_code & "'"
                            If summary_type = "S" Then sqlw &= " and SELL_CODE = '" & SELL_CODE & "' and SDS_CODE = '" & summary_code & "'"

                            For Each row As DataRow In dst.Tables("SPTSFOC9").Select(sqlw, "CUST_CODE, CUST_STORE_NO")
                                CSs &= vbCrLf & row.Item("CUST_STORE_NAME")
                            Next
                        End If

                        Dim BRAND_CODE As String = ""
                        Dim D As String = BOOKING_NAME
                        If NOTES <> "" Then D &= "-" & NOTES

                        For Each row As DataRow In dst.Tables("SPTSFOCE").Select _
                            ("EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'", "BRAND_CODE,COLLECTION_NAME")

                            Dim COLLECTION_NAME As String = row.Item("COLLECTION_NAME") & ""
                            Dim FEATURE_DESC As String = row.Item("FEATURE_DESC") & ""
                            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""

                            If BRAND_CODE = "" Then
                                BRAND_CODE = row.Item("BRAND_CODE") & ""
                            End If

                            If VEHICLE_CODE = "MA" And ITEM_CODE <> "" Then
                                D &= vbCrLf & ITEM_CODE
                                If FEATURE_DESC <> "" Then
                                    D &= " - " & FEATURE_DESC
                                End If
                            End If
                        Next

                        XR = Get_Next_Open(RM, YYYYMM, REL_WEEK, RX)

                        With worksheet.Cells(XR, C_Offset + REL_WEEK)
                            .WrapText = True
                            .Value = D & CSs
                            .Font.Color = BRAND_CODEs(BRAND_CODE)

                            If Format(EVENT_DATE_CHANGED, "yyyyMMdd") >= Format(Now.Date.AddDays(-28), "yyyyMMdd") Then
                                .Interior.Color = SpreadsheetGear.Colors.LightYellow
                            End If

                            If EVENT_FILE_LINK <> "" Then
                                XR = Get_Next_Open(RM, YYYYMM, REL_WEEK, RX)
                                worksheet.Hyperlinks.Add(worksheet.Cells(XR, C_Offset + REL_WEEK),
                                EVENT_FILE_LINK, "", "Click to Open File", EVENT_FILE_LINK)

                                'worksheet.Hyperlinks.Add(worksheet.Cells(XR, C_Offset + REL_WEEK),
                                'EVENT_FILE_LINK, "", "Click to Open File", D & CSs)

                            End If
                        End With
                    End If
                End If
            Next


            ' Events (COOP/SCENT/NATMEDIA) & VISUAL & RTLEVENTS

            Dim rowAUTHs() As DataRow = Nothing
            If CS = "*" Then
                rowAUTHs = dst.Tables("SPTCOOP1").Select("", "OPS_YYYYWW")
            Else
                rowAUTHs = dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "'", "OPS_YYYYWW")
            End If

            Dim AUTH_NOs_like As New List(Of String)

            For Each rowAUTH As DataRow In rowAUTHs
                Dim AUTH_NO As String = rowAUTH.Item("AUTH_NO")
                Dim rowSPTCOOP1 As DataRow = dst.Tables("SPTCOOP1").Rows.Find(AUTH_NO)
                Dim YYYYWW As String = rowSPTCOOP1.Item("OPS_YYYYWW")
                Dim VEHICLE_CODE As String = rowSPTCOOP1.Item("VEHICLE_CODE")
                Dim EXPENSE_TYPE_CODE As String = rowSPTCOOP1.Item("EXPENSE_TYPE_CODE")
                Dim CUST_CODE_X As String = rowSPTCOOP1.Item("CUST_CODE")

                Dim DATE_START As Date = CDate(rowSPTCOOP1.Item("DATE_START"))
                Dim DATE_END As Date = CDate(rowSPTCOOP1.Item("DATE_END"))
                Dim BOOKING_NAME As String = rowSPTCOOP1.Item("BOOKING_NAME") & ""
                Dim EVENT_FILE_LINK As String = rowSPTCOOP1.Item("EVENT_FILE_LINK") & ""
                Dim EVENT_DATE_CHANGED As Date = rowSPTCOOP1.Item("EVENT_DATE_CHANGED")
                Dim skip As Boolean = False
                If RUNAS = "C" Then
                    ' don't skip - print all for Corporate
                ElseIf CUST_CODEs.Contains(CUST_CODE_X) Then
                    ' don't skip if Customer is in Scope
                    ' however, if RTLEVENT and the store is not explicitly listed, then skip
                    ' WE ARE NOT LOOKING AT STORES INCLUDED UNTIL WE GET INTO THE NEXT BLOCK
                    'If EXPENSE_TYPE_CODE = "RTLEVENTS" And "" = "not one of my stores" Then
                    '    skip = True
                    'End If
                Else
                    skip = True
                End If

                If AUTH_NOs_like.Contains(AUTH_NO) Then
                    skip = True
                End If

                If Weeks_TY.ContainsKey(YYYYWW) And Not skip Then
                    W = Weeks_TY(YYYYWW)

                    Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
                    Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")
                    Dim WEEK_END_DATE As Date = CDate(rowGLTPARM3.Item("WEEK_END_DATE"))
                    Dim YYYYMM As String = rowGLTPARM3.Item("YYYYMM") & ""

                    Dim RX As Integer = 0
                    If EXPENSE_TYPE_CODE = "COOP" Or EXPENSE_TYPE_CODE = "SCENT" Or EXPENSE_TYPE_CODE = "NATMEDIA" Then
                        RX = 6
                    ElseIf EXPENSE_TYPE_CODE = "RTLEVENTS" Then
                        RX = 7
                    ElseIf EXPENSE_TYPE_CODE = "VISUAL" Then
                        RX = 1
                    End If

                    Dim CSs As String = ""

                    If summary_type = "T" Or summary_type = "S" Or summary_type = "R" Then
                        Dim sqlw As String = "AUTH_NO = '" & AUTH_NO & "'"

                        Dim door_specific As Boolean = False
                        If dst.Tables("SPTCOOPA").Select(sqlw).Length > 0 Then
                            door_specific = True
                        End If

                        If CS <> "*" Then sqlw &= " and " & sqlCS
                        If summary_type = "T" Then sqlw &= " and SELL_CODE = '" & summary_code & "'"
                        If summary_type = "S" Then sqlw &= " and SELL_CODE = '" & SELL_CODE & "' and SDS_CODE = '" & summary_code & "'"
                        If summary_type = "R" Then sqlw &= " and REGION_CODE = '" & summary_code & "'"

                        For Each row As DataRow In dst.Tables("SPTCOOPA").Select(sqlw, "CUST_CODE, CUST_STORE_NO")
                            CSs &= vbCrLf & row.Item("CUST_STORE_NAME")
                        Next

                        If door_specific And EXPENSE_TYPE_CODE = "RTLEVENTS" And CSs = "" Then
                            skip = True
                        End If
                    End If

                    If Not skip Then

                        Dim D As String = ""
                        If EXPENSE_TYPE_CODE = "COOP" Or EXPENSE_TYPE_CODE = "SCENT" Or EXPENSE_TYPE_CODE = "NATMEDIA" Then
                            D = CUST_CODE_X & "-" & VEHICLE_CODE & "-" & BOOKING_NAME
                        ElseIf EXPENSE_TYPE_CODE = "RTLEVENTS" Then
                            D = CUST_CODE_X & " (" & Format(DATE_START, "MM/dd") & "-" & Format(DATE_END, "MM/dd") & ") " & BOOKING_NAME
                        ElseIf EXPENSE_TYPE_CODE = "VISUAL" Then
                            D = CUST_CODE_X & " (" & Format(DATE_START, "MM/dd") & "-" & Format(DATE_END, "MM/dd") & ") " & BOOKING_NAME
                        End If

                        Dim AUTH_NOs_to_list As New List(Of String)
                        AUTH_NOs_to_list.Add(AUTH_NO)

                        If EXPENSE_TYPE_CODE = "COOP" Or EXPENSE_TYPE_CODE = "SCENT" Or EXPENSE_TYPE_CODE = "NATMEDIA" Then
                            ' find other AUTH_NOs like
                            For Each row2 As DataRow In rowAUTHs
                                Dim AUTH_NO2 As String = row2.Item("AUTH_NO") & ""
                                Dim VEHICLE_CODE2 As String = row2.Item("VEHICLE_CODE") & ""
                                Dim BOOKING_NAME2 As String = row2.Item("BOOKING_NAME") & ""
                                Dim CUST_CODE_X2 As String = row2.Item("CUST_CODE") & ""

                                If AUTH_NO2 <> AUTH_NO And CUST_CODE_X2 = CUST_CODE_X And VEHICLE_CODE2 = VEHICLE_CODE And BOOKING_NAME2 = BOOKING_NAME Then
                                    AUTH_NOs_like.Add(AUTH_NO2)
                                    AUTH_NOs_to_list.Add(AUTH_NO2)
                                End If
                            Next
                        End If

                        Dim BRAND_CODE As String = ""
                        If EXPENSE_TYPE_CODE = "RTLEVENTS" Then
                        Else
                            Dim FEATURE_DESC_last As String = ""
                            For Each AUTH_NO_to_list As String In AUTH_NOs_to_list
                                For Each row As DataRow In dst.Tables("SPTCOOPE").Select _
                                    ("AUTH_NO = '" & AUTH_NO_to_list & "'", "BRAND_CODE,COLLECTION_NAME")
                                    '             ("AUTH_NO = '" & AUTH_NO & "'", "BRAND_CODE,COLLECTION_NAME")
                                    If BRAND_CODE = "" Then BRAND_CODE = row.Item("BRAND_CODE") & ""
                                    ' Dim COLLECTION_NAME As String = row.Item("COLLECTION_NAME") & ""
                                    Dim FEATURE_DESC As String = row.Item("FEATURE_DESC") & ""
                                    If FEATURE_DESC <> "" And FEATURE_DESC_last <> FEATURE_DESC Then
                                        'D &= vbCrLf & FEATURE_DESC
                                        If EXPENSE_TYPE_CODE = "VISUAL" Then
                                        Else
                                            D &= "-" & FEATURE_DESC
                                        End If
                                    End If
                                    FEATURE_DESC_last = FEATURE_DESC
                                Next
                            Next

                        End If
                        '   If D = "LORD-08/20-08/20-August Trend Show" Then Stop

                        Dim more_to_come As Boolean = False
                        Do
                            XR = Get_Next_Open(RM, YYYYMM, REL_WEEK, RX)

                            With worksheet.Cells(XR, C_Offset + REL_WEEK)
                                .WrapText = True
                                .Value = D & CSs
                                .Font.Color = BRAND_CODEs(BRAND_CODE)
                                .Font.Underline = False

                                If Format(EVENT_DATE_CHANGED, "yyyyMMdd") >= Format(Now.Date.AddDays(-28), "yyyyMMdd") Then
                                    .Interior.Color = SpreadsheetGear.Colors.LightYellow
                                End If

                            End With

                            If EVENT_FILE_LINK <> "" Then
                                XR = Get_Next_Open(RM, YYYYMM, REL_WEEK, RX)
                                worksheet.Hyperlinks.Add(worksheet.Cells(XR, C_Offset + REL_WEEK),
                                    EVENT_FILE_LINK, "", "Click to Open File", EVENT_FILE_LINK)
                                worksheet.Cells(XR, C_Offset + REL_WEEK).WrapText = True
                                'worksheet.Hyperlinks.Add(worksheet.Cells(XR, C_Offset + REL_WEEK),
                                'EVENT_FILE_LINK, "", "Click to Open File", D & CSs)

                            End If

                            more_to_come = False
                            If EXPENSE_TYPE_CODE = "VISUAL" And Format(DATE_END, "yyyyMMdd") > Format(WEEK_END_DATE, "yyyyMMdd") Then

                                YYYYWW = ASCMAIN1.Week_Calc(YYYYWW, 1)
                                If Weeks_TY.ContainsKey(YYYYWW) Then

                                    rowGLTPARM3 = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
                                    REL_WEEK = Val(rowGLTPARM3.Item("REL_WEEK") & "")
                                    WEEK_END_DATE = CDate(rowGLTPARM3.Item("WEEK_END_DATE"))
                                    YYYYMM = rowGLTPARM3.Item("YYYYMM") & ""

                                    more_to_come = True

                                End If
                            End If
                        Loop While more_to_come
                    End If
                End If
            Next

            ' Top of Page

            worksheet.Cells(0, 0).Value = "'" & Format(Now, "MM/dd/yy")
            worksheet.Cells(1, 0).Value = ASCMAIN1.USER_ID
            worksheet.Cells(0, 1).EntireColumn.ColumnWidth = 1

            ' Headings

            With worksheet.PageSetup
                .Orientation = SpreadsheetGear.PageOrientation.Landscape
                '.PrintArea = Excel_Cell0(0, 0) & ":" + Excel_Cell0(0, 11)
                ''.PrintHeadings = True - this prints A B C across the top, and line numbers down the page
                '.PrintTitleRows = Excel_Cell0(0, 0) & ":" & Excel_Cell0(0 + 2, 1)
                .FitToPagesWide = 1
                .FitToPagesTall = 0
            End With

            range = worksheet.Cells("A:Z").EntireColumn
            '  range.AutoFit()

            'worksheet.Cells("D3").Activate()
            'worksheet.WindowInfo.FreezePanes = True
        Next

        workbook.Worksheets(0).Select()

        workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(XLS_FILENAME)
    End Sub

    Function Get_Next_Open(RM As Dictionary(Of String, Integer), YYYYMM As String, REL_WEEK As Integer, RX As Integer) As Integer

        Dim RA As Integer = 0
        Dim RA_line As String = YYYYMM & ":" & CStr(RX)
        Dim RA_key As String = YYYYMM & ":" & CStr(REL_WEEK) & ":" & CStr(RX)
        If Not RAs.ContainsKey(RA_key) Then

        End If
        If Not RAs.ContainsKey(RA_key) Then
            RAs.Add(RA_key, 1)
            If Not RAs.ContainsKey(RA_line) Then
                RAs.Add(RA_line, 1)

                If Not RLs.ContainsKey(RA_line) Then
                    For iRX As Integer = 1 To 7
                        RLs.Add(YYYYMM & ":" & CStr(iRX), 0)
                    Next
                End If
            End If
        Else
            RA = RAs(RA_key)
            RAs(RA_key) += 1

            If RAs(RA_key) > RAs(RA_line) Then
                RAs(RA_line) += 1

                Dim YMs As New List(Of String)
                For Each YM As String In RM.Keys
                    If YM > YYYYMM Then
                        YMs.Add(YM)
                    End If
                Next

                For Each YM As String In YMs
                    RM(YM) += 1
                Next
                If RX < 7 Then
                    For iRX As Integer = RX + 1 To 7
                        RLs(YYYYMM & ":" & CStr(iRX)) += 1
                    Next
                End If

                Dim pRA As Integer = RM(YYYYMM) + RX + RLs(YYYYMM & ":" & CStr(RX))
                worksheet.Cells(RA + pRA, 0).EntireRow.Insert()
                With worksheet.Cells(0 + pRA, 2, RA + pRA, 2)
                    .Merge()
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous

                End With

                Dim Weeks_in_Month As Integer = dst.Tables("GLTPARM3").Select("YYYYMM='" & YYYYMM & "'").Length
                With worksheet.Cells(0 + pRA, 2 + 1, RA + pRA, 2 + Weeks_in_Month)
                    .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.None
                    .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

            End If
        End If

        Return RA + RM(YYYYMM) + RX + RLs(YYYYMM & ":" & CStr(RX))
    End Function

    Private Sub optRUNAS_ValueChanged(sender As Object, e As EventArgs) Handles optRUNAS.ValueChanged
        cmbREGION_CODE.Visible = (optRUNAS.Value = "R")
        cmbSELL_CODE.Visible = (optRUNAS.Value = "T")
    End Sub

    Private Sub txtSEASON_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtSEASON_CODE.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Set_YPs()
    End Sub

    Sub Set_YPs()
        Dim SEASON_CODE As String = txtSEASON_CODE.Text
        If SEASON_CODE <> "" Then
            Dim row As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
            If row IsNot Nothing Then
                Dim SEASON_YEAR As String = row.Item("SEASON_YEAR")
                Dim SEASON_TYPE As String = row.Item("SEASON_TYPE")

                Dim YP1 As String = SEASON_YEAR & IIf(SEASON_TYPE = "S", "01", "07")
                Dim YP2 As String = SEASON_YEAR & IIf(SEASON_TYPE = "S", "06", "12")

                Absx1.cmbFor("RYP0").Value = LookUp("GLTPARM2", YP1).Item("LEGEND")
                Absx1.cmbFor("RYP1").Value = LookUp("GLTPARM2", YP2).Item("LEGEND")
            End If
        End If

    End Sub
End Class