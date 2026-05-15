Imports Microsoft.Office.Interop
Public Class GLRPROF1

#Region "General Declarations"

    Dim NYP As String

    Dim RCX As String
    Dim COL_LAYOUT_CODE As String
    Dim COL_LAYOUT_DESC As String
    Dim COL_CODE_RECAP As Integer
    Dim COL_CODE As String
    Dim RCOLS(,) As String
    Dim RCOLS_count As Integer
    Dim COL_ORA As Dictionary(Of String, String)

    Dim CHKTHOUSANDS As String
    Dim CAPTION_included As Boolean = False
    Dim CAPTIONSUB_included As Boolean = False

    Dim oWB As SpreadsheetGear.IWorkbook
    Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim XL_RH As Integer
    Dim XL_CH As Integer
    Dim XL_DESC As Integer

    Dim XL_ROWS As Integer
    Dim XL_COLS As Integer
    Dim XL_COLS_D As Integer
    Dim XL_ROWS_TOTAL As Integer
    Dim XL_COL_CAPTION As Integer

    Dim recap As Boolean
    Dim LINE_CAPTION As String
    Dim LINE_CAPTION_SUB As String

    Dim SHEET_ENTITY_DESC As String
    Dim SHEET_NAME As String
    Dim SI As Integer   ' Sheet Index

    Dim cfi As Integer ' Highest PB level where the code value is .Recap
    Dim cfmax As Integer

    Dim GLTPROF1 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        RWU = "N"
        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        ' Get Run-Time options

        CHKTHOUSANDS = IIf(Absx1.chkFor("THOUSANDS").Checked, "Y", "N")

        NYP = ASCMAIN1.Period_Calc(RYP, 12)                 ' Period Selected + 12; Next Year, Same Period
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)  ' Period Selected - 12; Last Year, Same Period
        Dim RY As String = Mid$(RYP, 1, 4)                  ' Year YYYY of Period Selected
        Dim RP As String = Mid$(RYP, 5, 2)                  ' Period PP of Period Selected
        Dim LY As String = Mid$(LYP, 1, 4)                  ' Last Year (Year of Period Selected -1)
        Dim NY As String = Mid$(NYP, 1, 4)                  ' Next Year (Year of Period Selected +1)

        RCX = " .Recap "

        ' Setup Column Expressions and Load into Report Layout Selected

        COL_ORA = New Dictionary(Of String, String)

        Dim COL_EXP As New Collection
        COL_EXP.Add("ASTSRPT1.ACT_LY_YTL", "ACT_LY_YTL")
        COL_EXP.Add("ASTSRPT1.ACT_LY_YTD", "ACT_LY_YTD")
        COL_EXP.Add("ASTSRPT1.ACT_LY_MTL", "ACT_LY_MTL")
        COL_EXP.Add("ASTSRPT1.BUD_TY_YTL", "BUD_TY_YTL")
        COL_EXP.Add("ASTSRPT1.BUD_TY_YTD", "BUD_TY_YTD")
        COL_EXP.Add("ASTSRPT1.BUD_TY_MTL", "BUD_TY_MTL")
        COL_EXP.Add("ASTSRPT1.ACT_TY_YTD", "ACT_TY_YTD")
        COL_EXP.Add("ASTSRPT1.ACT_TY_MTD", "ACT_TY_MTD")
        COL_EXP.Add("ASTSRPT1.ACT_TY_YTD - ASTSRPT1.ACT_LY_YTD", "VTY_LY_YTD") ' was ._ACT_LY_YTD
        COL_EXP.Add("ASTSRPT1.ACT_TY_YTD - ASTSRPT1.BUD_TY_YTD", "VTY_BD_YTD") ' was ._BUD_TY_YTD
        COL_EXP.Add("GLTPROF1.LINE_CAPTION", "CAPTION")
        COL_EXP.Add("ASTSRPT1.ACT_LY_YTL - ASTSRPT1.ACT_LY_YTD", "ACT_LY_YTG")
        COL_EXP.Add("ASTSRPT1.RBD_TY_YTL - ASTSRPT1.ACT_TY_YTD", "ACT_TY_YTG")
        COL_EXP.Add("GLTPROF1.LINE_CAPTION_SUB", "CAPTIONSUB")
        COL_EXP.Add("ASTSRPT1.ACT_TY_YTD + ASTSRPT1.BUD_TY_YTG", "ACT_TY_PRJ")
        COL_EXP.Add("ASTSRPT1.BUD_TY_YTG", "BUD_TY_YTG")
         
        COL_EXP.Add("IIF(ASTSRPT1.ACT_LY_YTD = 0, 0, 100 * (ASTSRPT1.ACT_TY_YTD - ASTSRPT1.ACT_LY_YTD) / ASTSRPT1.ACT_LY_YTD)", "PTY_LY_YTD")
        COL_ORA.Add("PTY_LY_YTD", "CASE WHEN ASTSRPT1.ACT_LY_YTD = 0 THEN 0 ELSE ROUND(100 * (ASTSRPT1.ACT_TY_YTD - ASTSRPT1.ACT_LY_YTD) / ASTSRPT1.ACT_LY_YTD,2) END")

        COL_EXP.Add("IIF(ASTSRPT1.BUD_TY_YTD = 0, 0, 100 * (ASTSRPT1.ACT_TY_YTD - ASTSRPT1.BUD_TY_YTD) / ASTSRPT1.BUD_TY_YTD)", "PTY_BD_YTD")
        COL_ORA.Add("PTY_BD_YTD", "CASE WHEN ASTSRPT1.BUD_TY_YTD = 0 THEN 0 ELSE ROUND(100 * (ASTSRPT1.ACT_TY_YTD - ASTSRPT1.BUD_TY_YTD) / ASTSRPT1.BUD_TY_YTD,2) END")

        COL_EXP.Add("IIF(ASTSRPT1.BUD_TY_YTL = 0, 0, 100 * (ASTSRPT1.ACT_LY_YTL - ASTSRPT1.BUD_TY_YTL) / ASTSRPT1.BUD_TY_YTL)", "PLY_BD_YTL")
        COL_ORA.Add("PLY_BD_YTL", "CASE WHEN ASTSRPT1.BUD_TY_YTL = 0 THEN 0 ELSE ROUND(100 * (ASTSRPT1.ACT_LY_YTL - ASTSRPT1.BUD_TY_YTL) / ASTSRPT1.BUD_TY_YTL,2) END")

        COL_EXP.Add("ASTSRPT1.ACT_LY_YTL - ASTSRPT1.BUD_TY_YTL", "VLY_BD_YTL")


        ASCMAIN1.sql = "Select T_CODE, T_DESC from ASTCODE1 " _
            & " where TABLE_NAME = 'GLTPROF2' " _
            & "   and COLUMN_NAME = 'COL_CODE'"
        Dim tblCOL_CODEs As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTCODE1", 1)

        COL_LAYOUT_CODE = Absx1.cmbFor("COL_LAYOUT_CODE").Text
        COL_CODE = Absx1.cmbFor("COL_CODE_RECAP").Text


        Dim rowGLTPROF2 As DataRow = LookUp("GLTPROF2", COL_LAYOUT_CODE)
        COL_LAYOUT_DESC = rowGLTPROF2.Item("COL_LAYOUT_DESC") & ""

        ReDim RCOLS(9, 3)

        CAPTION_included = False
        CAPTIONSUB_included = False

        For i As Integer = 1 To 9
            Dim COL_CODE_XX As String = rowGLTPROF2.Item("COL_CODE_" & Format$(i, "00")) & ""
            If COL_CODE_XX = "" Then
                Exit For
            End If

            If COL_CODE_XX = "CAPTION" Then
                CAPTION_included = True
            ElseIf COL_CODE_XX = "CAPTIONSUB" Then
                CAPTIONSUB_included = True
            End If

            RCOLS_count = i
            RCOLS(i, 0) = COL_CODE_XX

            If RCOLS(i, 0) = COL_CODE Then
                COL_CODE_RECAP = i
                RCOLS(0, 0) = COL_CODE
            End If
        Next i

        For i As Integer = 0 To RCOLS_count
            If RCOLS(i, 0) <> "" Then
                Dim rowCOL_CODE As DataRow = tblCOL_CODEs.Rows.Find(RCOLS(i, 0))
                Dim T_DESC As String = rowCOL_CODE.Item("T_DESC") & ""
                RCOLS(i, 1) = Split(T_DESC & ", ", ", ")(0)
                RCOLS(i, 2) = Split(T_DESC & ", ", ", ")(1)
                RCOLS(i, 3) = COL_EXP(RCOLS(i, 0))
            End If
        Next i

        Create_TDA(dst.Tables.Add, "GLTPROFZ", "*", 0, False, "", 0)

        ' Prepare Work File with Data from Server

        ASCMAIN1.Progress("Gathering Report Data", "")
        MyBase.Get_SQL("*")

        Dim sql_Cols As String = ""
        Dim sql_data As String = ""
        For Each C As String In COLUMN_NAME_sum.Keys
            sql_data &= ", SUM (GLTPROFZ." & C & IIf(CHKTHOUSANDS = "Y", "/1000", "") & ") AS " & C
            sql_Cols &= "," & C
        Next

        Dim sql_filter As String = " and GLTPROFZ.OPS_YYYYPP = '" & RYP & "'"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & COLUMN_NAMEs_appended & vbCrLf & sql_data _
            & " from GLTPROFZ" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & COLUMN_NAMEs_appended

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        ' Complete P&L for Each Page

        ASCMAIN1.Progress("Ensure Complete P/L Format", "")

        Dim SQLG As String = ""
        For i As Integer = 1 To 9
            SQLG &= Replace(", GX = NVL(GX,'x')", "GX", "G" & CStr(i))
        Next i
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set " & Mid(SQLG, 2)

        ASCMAIN1.sql = "Select Distinct G1, G2, G3, G4, G5, G6, G7, G8, G9, GLTPROF1.LINE_TAG, '0' AS MATCH from " & ASTSRPT1 & ", GLTPROF1"
        Dim GLTPROFQ As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFQ & " Add Primary Key (G1,G2,G3,G4,G5,G6,G7,G8,G9,LINE_TAG)")

        ASCMAIN1.sql = "Update " & GLTPROFQ & vbCrLf _
            & " Set MATCH = '1' where (G1, G2, G3, G4, G5, G6, G7, G8, G9, LINE_TAG)" & vbCrLf _
            & " in (Select G1, G2, G3, G4, G5, G6, G7, G8, G9, LINE_TAG from " & ASTSRPT1 & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & " (G1,G2,G3,G4,G5,G6,G7,G8,G9,LINE_TAG)" & vbCrLf _
            & " Select G1,G2,G3,G4,G5,G6,G7,G8,G9,LINE_TAG from " & GLTPROFQ & " where MATCH = '0'"
        ASCDATA1.ExecuteSQL()


        'where NVL(LINE_SUPPRESS,'0') <> '1'
        ASCMAIN1.sql = "Select X.*, ROWNUM RECORD_INDEX from (Select * from GLTPROF1 where NVL(LINE_SUPPRESS,'0') <> '1' order by LINE_NO) X"
        GLTPROF1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROF1 & " Add Primary Key (LINE_NO)")
        ASCDATA1.ExecuteSQL("Update " & GLTPROF1 & " Set LINE_NO = RECORD_INDEX")

        ASCMAIN1.sql = "Select * from " & GLTPROF1
        Create_TDA(dst.Tables.Add, "GLTPROF1", "**", 0, False, "", 1)
        Fill_Records("GLTPROF1")

    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        Dim LINE_NO_max As Integer = dst.Tables("GLTPROF1").Rows.Count
        Dim PROFBn As Integer = 9

        Dim TL As New Dictionary(Of String, Integer)


        Dim tbl As DataTable = dst.Tables("ASTSRPT1").Clone
        tbl.TableName = "ASTSRPT1_TOTAL"
        dst.Tables.Add(tbl)

        Dim tbl2 As DataTable = dst.Tables("ASTSRPT1").Clone

        Dim G1s As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), "G1").Select("")
            G1s.Add(row.Item(0))
        Next

        Dim LINE_TAGs As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), "LINE_TAG").Select("")
            LINE_TAGs.Add(row.Item(0))
        Next

        Dim RI As New Collection
        For Each rowGLTPROF1 As DataRow In dst.Tables("GLTPROF1").Select("")
            RI.Add(Val(rowGLTPROF1.Item("LINE_NO") & ""), rowGLTPROF1.Item("LINE_TAG") & "")
            TL.Add(rowGLTPROF1.Item("LINE_TAG"), rowGLTPROF1.Item("LINE_NO"))

            For Each G1 As String In G1s
                Dim rowTotal As DataRow = dst.Tables("ASTSRPT1_TOTAL").NewRow
                For i As Integer = 1 To 9
                    rowTotal.Item("G" & Format(i, "0")) = aRC
                Next i
                rowTotal.Item("G1") = G1
                rowTotal.Item("LINE_TAG") = rowGLTPROF1.Item("LINE_TAG") & ""
                dst.Tables("ASTSRPT1_TOTAL").Rows.Add(rowTotal)
            Next
        Next

        For Each LINE_TAG As String In LINE_TAGs
            If TL.ContainsKey(LINE_TAG) Then
                'OK
            Else
                ASCDATA1.DeleteRows(dst.Tables("ASTSRPT1"), "LINE_TAG = '" & LINE_TAG & "'")
            End If
        Next

        Create_Relation("ASTSRPT1_TOTAL", "ASTSRPT1", "G1,LINE_TAG")
        For Each COL As String In COLUMN_NAME_sum.Keys ' New String() {"ACT_TY_YTD", "ACT_TY_MTD"}
            dst.Tables("ASTSRPT1_TOTAL").Columns(COL).Expression = "SUM(CHILD." & COL & ")"
        Next


        For Each rowTotal As DataRow In dst.Tables("ASTSRPT1_TOTAL").Select("G1 = '" & aRC & "'")
            tbl2.Rows.Add(rowTotal.ItemArray)
        Next
         
        'For Each COL As String In COLUMN_NAME_sum.Keys ' New String() {"ACT_TY_YTD", "ACT_TY_MTD"}
        '    dst.Tables("ASTSRPT1_TOTAL").Columns(COL).Expression = ""
        'Next
        dst.Relations.Remove("ASTSRPT1_TOTAL_ASTSRPT1")
        dst.Tables("ASTSRPT1").Constraints.Remove("ASTSRPT1_TOTAL_ASTSRPT1")
        dst.Tables("ASTSRPT1_TOTAL").Rows.Clear()

        For Each rowTotal As DataRow In tbl2.Select("")
            dst.Tables("ASTSRPT1").Rows.Add(rowTotal.ItemArray)
        Next

        ' Do Calculations

        ASCMAIN1.Progress("Calculations", "")

        With dst.Tables("ASTSRPT1")
            .PrimaryKey = New DataColumn() {.Columns("G1"), .Columns("G2"), .Columns("G3"), .Columns("G4"), .Columns("G5"), .Columns("G6"), .Columns("G7"), .Columns("G8"), .Columns("G9"), .Columns("LINE_TAG")}
        End With

        Dim t(,) As Decimal
        ReDim t(LINE_NO_max, PROFBn)
        Dim KY(9) As String

        For Each rowGx As DataRow In ASCDATA1.SelectDistinct("ASTSRPT1", New String() {"G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9"}).Select("")

            Dim sqlX As String = ""
            For i As Integer = 1 To 9
                KY(i) = rowGx.Item(i - 1) & ""
                Dim z As String = rowGx.Item(i - 1) & ""
                sqlX &= " and G" & CStr(i) & " = '" & KY(i) & "'"
            Next i

            ReDim t(LINE_NO_max, PROFBn)

            For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select(Mid(sqlX, 6), "LINE_TAG")
                Dim LINE_TAG As String = rowASTSRPT1.Item("LINE_TAG")
                If TL.ContainsKey(LINE_TAG) Then
                    Dim LINE_NO As Integer = TL(LINE_TAG)
                    For k As Integer = 1 To PROFBn
                        t(LINE_NO, k) = Val(rowASTSRPT1.Item(k - 1 + 9 + 1) & "")
                    Next k
                End If
            Next

            For k As Integer = 1 To PROFBn

                ' NOT SURE WHY THESE ARE NECESSARY - MIGHT HAVE BEEN HANDLED DIFFERENTLY IN V1
                ' SEE ALSO SPECIAL UPDATE BELOW IN 1ST LOOP TO Write_Calculation
                t(RI("GSLS"), k) = t(RI("GSLSB"), k) + t(RI("GSLSP"), k)

                If ASCMAIN1.CLIENT = "AHA" Then
                    t(RI("TSRA"), k) = t(RI("RSLS"), k) + t(RI("RSLSAC"), k) - t(RI("PROMO"), k) - t(RI("COMM"), k) - t(RI("FIX"), k) - t(RI("DIF"), k) - t(RI("TST"), k) - t(RI("ARDED"), k) + t(RI("MISC"), k)
                    t(RI("NSLS"), k) = t(RI("GSLS"), k) + t(RI("TSRA"), k)
                Else
                    t(RI("NSLS"), k) = t(RI("GSLS"), k) + t(RI("RSLS"), k)
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    t(RI("GCGS"), k) = t(RI("GCGSB"), k) + t(RI("GCGSP"), k) + t(RI("SAMPLE"), k) + t(RI("GWP"), k) + t(RI("DSP"), k)
                Else
                    t(RI("GCGS"), k) = t(RI("GCGSB"), k) + t(RI("GCGSP"), k)

                    t(RI("ASP"), k) = t(RI("PROMO"), k) + t(RI("ADVERT"), k) + t(RI("COMM"), k)
                    t(RI("NC"), k) = t(RI("SAMPLE"), k) + t(RI("GWP"), k) + t(RI("DSP"), k)
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    t(RI("TCGS"), k) = t(RI("GCGS"), k) + t(RI("RCGS"), k) + t(RI("RCGSAC"), k)

                Else
                    t(RI("TCGS"), k) = t(RI("GCGS"), k) + t(RI("RCGS"), k)
                    ' t(RI("TCGS"), k) = t(RI("ROYALT"), k) + t(RI("GCGS"), k) ' NO ROYALTY
                End If


                t(RI("GP"), k) = t(RI("NSLS"), k) - t(RI("TCGS"), k)


                '  t(RI("TOTASP"), k) = t(RI("DEMFIX"), k) + t(RI("MODROT"), k) + t(RI("TAXFEE"), k)
                '  t(RI("ACTMAR"), k) = t(RI("GP"), k) - t(RI("ASP"), k) - t(RI("NC"), k) - t(RI("TOTASP"), k) - t(RI("DEMCOM"), k) - t(RI("AE"), k) - t(RI("RAM"), k)
                If ASCMAIN1.CLIENT = "AHA" Then
                Else
                    t(RI("ACTMAR"), k) = t(RI("GP"), k) - t(RI("ASP"), k) - t(RI("NC"), k) ' - t(RI("DEMCOM"), k)
                End If

                '   t(RI("OVRTOT"), k) = 0 '  t(RI("DIST"), k) + t(RI("MKTG"), k) + t(RI("SLSADM"), k) + t(RI("ADMIN"), k)

                If ASCMAIN1.CLIENT = "AHA" Then
                Else
                    t(RI("MIS"), k) = t(RI("ACTMAR"), k)  ' - t(RI("OVRTOT"), k)
                End If

                If t(RI("NSLS"), k) = 0 Then
                    t(RI("STHSIN"), k) = 0
                Else
                    t(RI("STHSIN"), k) = 100 * t(RI("RETAIL"), k) * 0.6 / t(RI("NSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("BASGRS"), k) = 0
                Else
                    t(RI("BASGRS"), k) = 100 * t(RI("GSLSB"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("PROGRS"), k) = 0
                Else
                    t(RI("PROGRS"), k) = 100 * t(RI("GSLSP"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("RTGSHP"), k) = 0
                Else
                    t(RI("RTGSHP"), k) = 100 * t(RI("RSLS"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("CGSSHP"), k) = 0
                Else
                    t(RI("CGSSHP"), k) = 100 * t(RI("GCGS"), k) / t(RI("GSLS"), k)
                End If
                If ASCMAIN1.CLIENT = "AHA" Then
                Else
                    If t(RI("NSLS"), k) = 0 Then
                        t(RI("ASPSHP"), k) = 0
                    Else
                        t(RI("ASPSHP"), k) = 100 * t(RI("TOTASP"), k) / t(RI("NSLS"), k)
                    End If
                    If t(RI("NSLS"), k) = 0 Then
                        t(RI("AMNSHP"), k) = 0
                    Else
                        t(RI("AMNSHP"), k) = 100 * t(RI("ACTMAR"), k) / t(RI("NSLS"), k)
                    End If
                End If

            Next k
            '{"GSLS", "NSLS", "GCGS", "TCGS", "GP", "ASP", "NC", "TOTASP", "ACTMAR", "OVRTOT", "MIS"}
            For Each LINE_TAG As String In New String() {"GSLS", "TSRA", "NSLS", "GCGS", "TCGS", "GP", "ASP", "NC", "ACTMAR", "MIS"}
                If ASCMAIN1.CLIENT = "AHA" And (LINE_TAG = "ASP" Or LINE_TAG = "NC" Or LINE_TAG = "ACTMAR" Or LINE_TAG = "MIS") Then
                Else
                    If Not (ASCMAIN1.CLIENT = "AHA") And LINE_TAG = "TSRA" Then
                    Else
                        Dim LINE_NO As Integer = RI(LINE_TAG) : Write_Calculation(LINE_TAG, LINE_NO, PROFBn, KY, t)
                    End If
                End If
            Next

            For Each LINE_TAG As String In New String() {"STHSIN", "BASGRS", "PROGRS", "RTGSHP", "CGSSHP", "ASPSHP", "AMNSHP"}
                If ASCMAIN1.CLIENT = "AHA" And (LINE_TAG = "ASPSHP" Or LINE_TAG = "AMNSHP") Then
                Else
                    Dim LINE_NO As Integer = RI(LINE_TAG) : Write_Calculation(LINE_TAG, LINE_NO, PROFBn, KY, t)
                End If
            Next
        Next

        ' Add Sub-Caption to Caption if Necessary

        If CAPTION_included And Not CAPTIONSUB_included Then
            For Each rowGLTPROF1 As DataRow In dst.Tables("GLTPROF1").Select("LINE_CAPTION_SUB is Not Null")
                rowGLTPROF1.Item("LINE_CAPTION") = rowGLTPROF1.Item("LINE_CAPTION") & " - " & rowGLTPROF1.Item("LINE_CAPTION_SUB")
            Next
        End If


        ' Output to Excel

        ASCMAIN1.Progress("Excel", "")

        For i As Integer = 1 To COLUMN_NAMEs.Count  ' cfmax
            For Each row As DataRow In dst.Tables("ASTSRPT1").Select("G" & CStr(i) & " = '" & aRC & "'") ' " LIKE 'Z~Recap *'")
                Dim GZ As String = row.Item("G" & CStr(i))
                Dim GI As Integer = Len(RCX) + 1
                Dim Z As String = RCX & COLUMN_CAPTIONs(i - 1) ' & GZ ' Mid(GZ, GI)
                row.Item("G" & CStr(i)) = Z
            Next
        Next i

        ASCMAIN1.sql = "Delete from " & ASTSRPT1
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("ASTSRPT1", "1=1")
        Load_Excel()

    End Sub

    Public Overrides Sub Print_Report()
        'SUBT = ""
        'CR_params.Add("SUBT", SUBT)
        'Generate_Report(RPT, , SUBT)

        ' KICK OUT THE EXCEL SPREADSHEET
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If Not Absx1.chkFor("RECAP_LAST_LEVEL").Checked Then
                Absx1.chkFor("RECAP_LAST_LEVEL").Checked = True
            End If

            If tblASTDSQLA.Select("SEQUENCE IS NOT NULL").Length < 2 Then
                EMsg &= vbCr & "2 Sort-Levels (minimum) must be selected"
            ElseIf tblASTDSQLA.Select("SEQUENCE IS NOT NULL").Length > 2 Then
                EMsg &= vbCr & "You may select only 2 Sort-Levels" ' more than 2 has not been tested
            End If

            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must pick at least 1 column to Sort by"
            End If

            Dim COL_LAYOUT_CODE As String = Absx1.cmbFor("COL_LAYOUT_CODE").Value & ""
            If COL_LAYOUT_CODE = "" Then
                EMsg &= vbCr & "You Must Select a Column Layout"
            Else
                Dim rowGLTPROF2 As DataRow = LookUp("GLTPROF2", COL_LAYOUT_CODE)
                If rowGLTPROF2 Is Nothing Then
                    EMsg &= vbCr & "You Must Select a Valid Column Layout"
                Else
                    Dim COL_CODE_RECAP As String = Absx1.cmbFor("COL_CODE_RECAP").Value & ""
                    If COL_CODE_RECAP = "" Then
                        EMsg &= vbCr & "You Must Select a Recap Column"
                    Else
                        For i As Integer = 1 To 12
                            Dim COL_CODE_XX As String = rowGLTPROF2.Item("COL_CODE_" & Format$(i, "00")) & ""
                            If COL_CODE_XX = COL_CODE_RECAP Then
                                Exit For
                            ElseIf i = 12 Then
                                EMsg &= vbCr & "You Must Select a Recap Column that exists in the Layout Selected"
                            End If
                        Next
                    End If
                End If

            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Private Sub cmbCOL_LAYOUT_CODE_ValueChanged(sender As Object, e As EventArgs) Handles cmbCOL_LAYOUT_CODE.ValueChanged

        Dim COL_LAYOUT_CODE As String = cmbCOL_LAYOUT_CODE.Value & ""
        Dim rowGLTPROF2 As DataRow = LookUp("GLTPROF2", COL_LAYOUT_CODE)
        Dim COL_CODE_RECAP As Integer = Val(rowGLTPROF2.Item("COL_CODE_RECAP") & "")
        If COL_CODE_RECAP = 0 Then COL_CODE_RECAP = 1

        Dim sqlCOL_CODE As String = ""
        For i As Integer = 1 To 12
            Dim COL_CODE_XX As String = rowGLTPROF2.Item("COL_CODE_" & Format$(i, "00")) & ""
            If COL_CODE_XX = "" Then
                Exit For
            Else
                If COL_CODE_XX <> "CAPTION" And COL_CODE_XX <> "CAPTIONSUB" Then
                    sqlCOL_CODE &= "OR COL_CODE = '" & COL_CODE_XX & "'"
                End If
            End If
        Next i

        Dim dvw As DataView = DirectCast(Absx1.cmbFor("COL_CODE_RECAP").DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlCOL_CODE, 4)
        Absx1.cmbFor("COL_CODE_RECAP").Value = rowGLTPROF2.Item("COL_CODE_" & Format$(COL_CODE_RECAP, "00"))

    End Sub

    Sub Write_Calculation(LINE_TAG As String, LINE_NO As Integer, PROFBn As Integer, KY() As String, t(,) As Decimal)

        Dim row As DataRow = dst.Tables("ASTSRPT1").Rows.Find(New String() {KY(1), KY(2), KY(3), KY(4), KY(5), KY(6), KY(7), KY(8), KY(9), LINE_TAG})
        If row Is Nothing Then
            row = dst.Tables("ASTSRPT1").NewRow
            For j As Integer = 1 To 9
                row.Item(j - 1) = KY(j)
            Next j
            row.Item(9) = LINE_TAG
            dst.Tables("ASTSRPT1").Rows.Add(row)
        End If
        ' If ASCMAIN1.Running_in_VS And KY(2) = "ULTA" Then Stop
        For k As Integer = 1 To PROFBn
            row.Item(k + 9) = t(LINE_NO, k)
        Next k
    End Sub

    Sub Load_Excel()

        ' Set up parameterized row and col settings

        XL_ROWS = dst.Tables("GLTPROF1").Rows.Count ' # of Rows in P&L - not counting space after
        XL_COLS = RCOLS_count   ' # of numeric columns in Layout Selected
        XL_RH = 0               ' # of Row Heading Cols, before 1st col of data
        XL_CH = 7               ' # of Col Heading Rows, before 1st row of data
        XL_DESC = 4             ' # OF Cols of line formatting data, these cols are deleted

        For i As Integer = 1 To XL_COLS
            If RCOLS(i, 0) = "CAPTION" Then
                XL_COL_CAPTION = i  ' Col where CAPTION is rendered in P&L
                Exit For
            End If
        Next i

        ' Create Workbook

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        ' Report Summary

        cfmax = COLUMN_NAMEs.Count  ' # of PBs selected

        Dim sql_Gs As String = ""
        For i As Integer = 1 To cfmax
            sql_Gs &= ", G" & CStr(i)
        Next i
        sql_Gs = Mid$(sql_Gs, 2)

        ' Get LINE_TAGs, in LINE_NO order, for Pivot Table columns

        Dim CODES As String = ""
        ASCMAIN1.sql = "Select GLTPROF1.LINE_TAG, GLTPROF1.LINE_NO from " & GLTPROF1 & " GLTPROF1"
        ASCMAIN1.sql &= " where NVL(LINE_SUPPRESS,'0') <> '1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "LINE_NO")
            CODES &= ",'" & row.Item(0) & "'"
        Next

        ASCMAIN1.sql = "Select * from" _
            & "(" _
            & "  Select " & sql_Gs & ", ASTSRPT1.LINE_TAG" & vbCrLf _
            & ", " & RCOLS(0, 3) & " " & RCOLS(0, 0) _
            & "  From " & ASTSRPT1 & " ASTSRPT1, " & GLTPROF1 & " GLTPROF1" & vbCrLf _
            & "  where GLTPROF1.LINE_TAG = ASTSRPT1.LINE_TAG" & vbCrLf _
            & ")" & vbCrLf _
            & " Pivot " & vbCrLf _
            & "(" & vbCrLf _
            & "  Sum(" & RCOLS(0, 0) & ")" & vbCrLf _
            & "  for LINE_TAG" & vbCrLf _
            & "  in (" & Mid(CODES, 2) & ")" & vbCrLf _
            & ")" '  where G1 Not Like '" & RCX & "%'"

        SI = 1
        oSheet = oWB.Sheets(SI - 1)
        oSheet.Name = "Report Summary"

        ' Load the DataTable into the Summary Sheet
        ' There will be data rows, as well as .Recap rows

        Dim tbl As DataTable = Load_DataTable(XL_CH + 0, XL_RH + XL_DESC + 1 + 1, sql_Gs)


        ' Format each column in the Summary Sheet

        Dim SHEET_COLS As Integer = tbl.Columns.Count
        For j As Integer = cfmax + 1 To SHEET_COLS
            Dim i As Integer = j - cfmax
            Dim rowGLTPROF1 As DataRow = dst.Tables("GLTPROF1").Rows.Find(i)
            With oSheet.Range(Excel_Cell(1, XL_RH + XL_DESC + cfmax + i + 1)).EntireColumn
                .ColumnWidth = 11
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                If rowGLTPROF1.Item("LINE_TYPE") & "" = "P" Then
                    .NumberFormat = "###,##0.0;(###,##0.0)"
                Else
                    If CHKTHOUSANDS = "Y" Then
                        .NumberFormat = "###,##0.0;(###,##0.0)"
                    Else
                        .NumberFormat = "###,##0;(###,##0)"
                    End If
                End If
            End With
            LINE_CAPTION = rowGLTPROF1.Item("LINE_CAPTION") & ""
            LINE_CAPTION_SUB = rowGLTPROF1.Item("LINE_CAPTION_SUB") & ""

            If LINE_CAPTION_SUB <> "" Then
                oSheet.Cells(Excel_Cell(XL_CH - 1, XL_RH + XL_DESC + cfmax + i + 1)).Value = LINE_CAPTION
                oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Value = LINE_CAPTION_SUB
            Else
                Dim z As String = LINE_CAPTION
                If Mid$(z, 1, 2) = "% " Then
                    oSheet.Cells(Excel_Cell(XL_CH - 2, XL_RH + XL_DESC + cfmax + i + 1)).Value = "%"
                    z = Mid$(z, 3)
                    Dim k As Integer = InStr(z, " vs ")
                    If k = 0 Then k = InStr(z, " to ")
                    If k = 0 Then k = InStr(z, "@")
                    If k = 0 Then k = InStr(z, " ")
                    If k = 0 Then
                        oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Value = z
                    Else
                        oSheet.Cells(Excel_Cell(XL_CH - 1, XL_RH + XL_DESC + cfmax + i + 1)).Value = Mid$(z, 1, k - 1)
                        oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Value = Mid$(z, k)
                    End If
                Else
                    If InStr(z, " ") = 0 Then
                        oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Value = z
                    Else
                        oSheet.Cells(Excel_Cell(XL_CH - 1, XL_RH + XL_DESC + cfmax + i + 1)).Value = Mid$(z, 1, InStr(z, " ") - 1)
                        oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Value = Mid$(z, InStr(z, " ") + 1)
                    End If
                End If
            End If

            If InStr("S" & "T", rowGLTPROF1.Item("LINE_TYPE") & "") <> 0 Then
                oSheet.Cells(Excel_Cell(XL_CH - 0, XL_RH + XL_DESC + cfmax + i + 1)).Font.Bold = True
                oSheet.Cells(Excel_Cell(XL_CH - 1, XL_RH + XL_DESC + cfmax + i + 1)).Font.Bold = True
            End If

            If rowGLTPROF1.Item("LINE_TYPE") & "" = "T" Then
                oSheet.Range(Excel_Cell(XL_CH + 1, XL_RH + XL_DESC + cfmax + i + 1) & ":" & Excel_Cell(XL_CH + tbl.Rows.Count, XL_RH + XL_DESC + cfmax + i + 1)).Interior.Color = SpreadsheetGear.Colors.LightGreen
            End If
        Next j

        Sheet_Heading()

       
        ' Show Recap Column

        With oSheet.Cells(XL_CH - 3, XL_RH + cfmax + 1)
            .Value = RCOLS(0, 1) & ", " & RCOLS(0, 2)
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        ' Column Captions for Sorted Columns

        For j As Integer = 1 To cfmax
            oSheet.Cells(XL_CH - 2, XL_RH + j).Value = COLUMN_CAPTIONs(j - 1)
            oSheet.Cells(XL_CH - 1, XL_RH + j).Value = "" ' to get rid of G1, G2, etc
        Next j

        ' Change the value from " .Recap *" to "Totals" for each column what was recapped
        ' Fix the value of the codes from Customer:MACYS to MACYS for non-recap values

        Dim i_counter As Integer = 1
        Dim jj As Integer = 0
        Do
            Dim CODE_VALUE As String = ""
            For j As Integer = 1 To cfmax
                CODE_VALUE = oSheet.Cells(Excel_Cell(XL_CH + i_counter, XL_RH + j + 1)).Value
                If InStr(CODE_VALUE, RCX) = 1 Then

                    CODE_VALUE = "Totals"
                    If jj = 0 And j <> cfmax Then CODE_VALUE = "Grand"
                    oSheet.Cells(Excel_Cell(XL_CH + i_counter, XL_RH + j + 1)).Value = CODE_VALUE
                    If j = cfmax Or i_counter = 1 Then
                        jj = jj + 1
                    End If
                Else
                    'Stop
                    CODE_VALUE = Mid$(CODE_VALUE, InStr(CODE_VALUE, ":") + 1)
                    oSheet.Cells(Excel_Cell(XL_CH + i_counter, XL_RH + j + 1)).Value = CODE_VALUE
                End If
            Next j

            Dim SHEET_NO_formatted As String = ""
            SHEET_NO_formatted = Format$(jj, "000")

            ' note that this sheet index value will change when we insert the re-caps for each level in the section below
            oSheet.Cells(Excel_Cell(XL_CH + i_counter, XL_RH + 1)).Value = "'" & SHEET_NO_formatted
            ' I THINK I SHOULD BE USING I_COUNTER AND NOT JJ-1 - PROB WILL SEE THE NEED WHEN i TEST 3 PBS
            i_counter += 1
            jj += 1
        Loop While oSheet.Cells(Excel_Cell(XL_CH + i_counter, XL_RH + 1 + 1)).Value <> ""


        ' Add a worksheet for each row in the Report Summary
        ' - which at this point has Data and .Recap rows, where a .Recap row is really a Total row
        ' - it does not have what will be known as Recap Sheets, 
        '   where (for a Brand/Customer Report) the Customer's are recapped as separate columns for the Brand

        ' The only worksheet that exists at this point is the Report Summary sheet

        '    Dim TBLX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), Split(Trim(sql_Gs), ", "))

        For Each rowASTSRPT1_D As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), Split(Trim(sql_Gs), ", ")).Select("", sql_Gs)
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("(" & CStr(SI) & ") " & rowASTSRPT1_D.Item("G1") & "-" & rowASTSRPT1_D.Item("G2") & "", "")

            Dim z As String = ""
            recap = False
            For cfi = cfmax To 1 Step -1
                z = rowASTSRPT1_D.Item("G" & CStr(cfi)) & ""
                If InStr(z, RCX) <> 1 Then
                    Exit For
                End If
            Next cfi

            If cfi = 0 Then z = ""

            ' cfi is the level that we are reporting data for
            ' we determine cfi by looking for the first field that does not say ".Recap", starting with cfmax and working towards 1
            ' need to test a report with 3 PBs so that we can see how we are handling .recaps both before and after real data

            Dim recap_format As Boolean  = False
            SHEET_NAME = Replace$(Replace$(Replace$(Replace$(z, "?", "Unk"), "*", "All"), ":", "-"), "/", "-")
            If InStr(SHEET_NAME, RCX) = 1 Then
                recap_format = True
                SHEET_NAME = Mid$(SHEET_NAME, Len(RCX) + 1)
            End If
            SHEET_NAME = Format$(SI - 1, "000") & " " & SHEET_NAME

            If cfi <> cfmax Then
                ' GIVING UP NOW 11/18/17 - THERE IS SOMETHING WRONG OR INCONGRUOUS WITH THE RECAP DATA IN ASTSRPT1 
                '  AND THE REQUIREMENTS OF THIS LOGIC TO PROVIDE RECAP SHEETS - WHICH ARE REALLY COOL PIVOTS
                ' - SO TOO BAD WE CANNOT GET THEM TO WORK NOW
                ' TO RESUME THIS WORK, SEE ASFSRPTM SEARCH FOR GLRPROF1 
                ' - CODE DISABLED FOR NOW, AND BE AWARE THAT AS A HACK i THREW THE TOTALS INTO THE TABLE USING ASTSRPT1_TOTALS
                '  - ALL OF WHICH PROBABLY SHOULD BE REMOVED IF AND WHEN WE GET THESE PIVOT RECAPS TO WORK
                ' REMMING OUT THE BELOW FOR NOW, BECAUSE IT JUST MAKES A BLANK SHEET
                ' If False Then
                recap = True
                Load_Sheet_Recap(rowASTSRPT1_D)
                SHEET_NAME = SHEET_NAME & " Totals"
                recap = False
                ' End If
            End If

            SI = SI + 1
            oSheet = oWB.Worksheets.Add
            oSheet.Cells.Font.Name = "Times New Roman"

            Mid$(SHEET_NAME, 1, 3) = Format$(SI - 1, "000")
            oSheet.Name = SHEET_NAME

            Load_Sheet(rowASTSRPT1_D)
            Format_Sheet(tbl)

            Dim WITHIN As String = ""
            If cfi > 1 Then
                For i As Integer = 1 To cfi - 1
                    WITHIN = WITHIN & "; " & rowASTSRPT1_D.Item("G" & CStr(i))
                Next i
                WITHIN = Mid$(WITHIN, 3)
            End If

            With oSheet.Cells(Excel_Cell(3, 1))
                .Value = WITHIN
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            End With

            If cfi <> cfmax And COL_CODE_RECAP <> 0 Then
                ' Copy the recap column from the Totals sheet to the last column of the recap sheet 
                ' - pretty sneaky, but the %'ages at the bottom of the report are a mess for each of the recapped entities - thats why we null out Ps in Pivot sql
                'Dim rangeCopyFrom As SpreadsheetGear.IRange = oSheet.Range(Excel_Cell(1, XL_RH + COL_CODE_RECAP) & ":" & Excel_Cell(XL_ROWS_TOTAL, XL_RH + COL_CODE_RECAP))
                'Dim rangePaste_To As SpreadsheetGear.IRange = oWB.Worksheets(SI - 2).Range(Excel_Cell(1, XL_RH + XL_COLS_D + 1) & ":" & Excel_Cell(XL_ROWS_TOTAL, XL_RH + XL_COLS_D + 1))

                ' switching to entire column because XL_ROWS_TOTAL didn't seem to copy to %s in the total column

                Dim rangeCopyFrom As SpreadsheetGear.IRange = oSheet.Range(Excel_Cell(1, XL_RH + COL_CODE_RECAP)).EntireColumn
                Dim rangePaste_To As SpreadsheetGear.IRange = oWB.Worksheets(SI - 2).Range(Excel_Cell(1, XL_RH + XL_COLS_D + 1)).EntireColumn

                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                oWB.Worksheets(SI - 2).Select()
                oWB.Worksheets(SI - 2).Cells(Excel_Cell(XL_CH - 2, XL_RH + XL_COLS_D + 1)).Value = "Totals"
                oWB.Worksheets(SI - 2).Range("A1:A1").Select()
            End If

            'If Not recap_format Then
            '    oWB.Worksheets(SI - 1).Range(Excel_Cell(XL_CH + 1, 1)).Select()
            'Else
            '    oWB.Worksheets(SI - 1).Range(Excel_Cell(XL_CH + 1, 3)).Select()
            'End If
            'oWB.Worksheets(SI - 1).WindowInfo.FreezePanes = True

            If cfi <> cfmax And COL_CODE_RECAP <> 0 Then
                oWB.Worksheets(SI - 1).Range("A1:A1").Select()
            End If
        Next

End_of_Process:

        ASCMAIN1.Progress("Now Hyperlinking", "")
        oSheet = oWB.Worksheets(0)
        Dim iSheet As Integer = 1
        If oWB.Worksheets.Count > 1 Then
            Do
                Dim j As Integer = oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 1)).Value + 1
                If j - 1 <> iSheet Then
                    oSheet.Range(Excel_Cell(XL_CH + iSheet, 1) & ":" & Excel_Cell(XL_CH + iSheet, XL_RH + 1 + SHEET_COLS + 1)).Insert(SpreadsheetGear.InsertShiftDirection.Down)
                    oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 1)).Value = "'" & Format$(iSheet, "000")
                    oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 2)).Value = oSheet.Cells(Excel_Cell(XL_CH + iSheet + 1, XL_RH + 2)).Value
                    oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 3)).Value = "Recap"
                    oSheet.Hyperlinks.Add(oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 1)), "", "'" & oWB.Worksheets(j - 2).Name & "'!" & Excel_Cell(4, XL_RH + 1), oWB.Worksheets(j - 2).Name, "'" & Format$(iSheet, "000"))

                    With oWB.Worksheets(j - 2)
                        .Hyperlinks.Add(.Cells(Excel_Cell(5, XL_RH + 1)), "", "'" & oSheet.Name & "'!" & Excel_Cell(XL_CH + iSheet, XL_RH + 1), oSheet.Name, "Back to Report Summary")
                    End With

                    iSheet += 1
                End If

                If oWB.Worksheets.Count >= j Then
                    oSheet.Hyperlinks.Add(oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 1)), "", "'" & oWB.Worksheets(j - 1).Name & "'!" & Excel_Cell(4, XL_RH + 1), oWB.Worksheets(j - 1).Name, "'" & Format$(iSheet, "000"))
                    With oWB.Worksheets(j - 1)
                        .Hyperlinks.Add(.Cells(Excel_Cell(5, XL_RH + 1)), "", "'" & oSheet.Name & "'!" & Excel_Cell(XL_CH + iSheet, XL_RH + 1), oSheet.Name, "Back to Report Summary")
                    End With

                End If
                iSheet += 1
            Loop While oSheet.Cells(Excel_Cell(XL_CH + iSheet, XL_RH + 1)).Value <> ""

            oSheet.Range(Excel_Cell(XL_CH - 1, XL_RH + 1) & ":" & Excel_Cell(XL_CH - 0, XL_RH + 1 + SHEET_COLS)).Interior.Color = SpreadsheetGear.Colors.Gold
            oSheet.Range(Excel_Cell(XL_CH, XL_RH + 1) & ":" & Excel_Cell(XL_CH, XL_RH + 1 + SHEET_COLS)).Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            oSheet.Range(Excel_Cell(XL_CH - 2, XL_RH + 1) & ":" & Excel_Cell(XL_CH - 2, XL_RH + 1 + SHEET_COLS)).Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous

        End If

        oWB.Worksheets(0).Select()
        oWB.Worksheets(0).Range("A1:A1").Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Sub Format_Sheet(tbl As DataTable)

        If recap Then
            oSheet.Range(XL_CH, XL_RH + XL_DESC + 1, XL_CH + XL_ROWS - 1, XL_RH + XL_DESC + 1).Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
        Else
            With oSheet.Range(XL_CH, XL_RH + XL_DESC + XL_COL_CAPTION - 1, XL_CH + XL_ROWS - 1, XL_RH + XL_DESC + XL_COL_CAPTION - 1)
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
            If UBound(RCOLS, 1) > XL_COL_CAPTION Then
                If RCOLS(XL_COL_CAPTION + 1, 0) = "CAPTIONSUB" Then
                    With oSheet.Range(XL_CH, XL_RH + XL_DESC + XL_COL_CAPTION, XL_CH + XL_ROWS - 1, XL_RH + XL_DESC + XL_COL_CAPTION)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With

                End If
            End If
        End If

        oSheet.Range(Excel_Cell(XL_CH, XL_RH + 1) & ":" & Excel_Cell(XL_CH, XL_RH + XL_DESC + XL_COLS)).Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        oSheet.Range(Excel_Cell(XL_CH - 2, XL_RH + 1) & ":" & Excel_Cell(XL_CH - 2, XL_RH + XL_DESC + XL_COLS)).Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous

        oSheet.Range(Excel_Cell(XL_CH - 1, XL_RH + 1) & ":" & Excel_Cell(XL_CH, XL_RH + XL_DESC + XL_COLS)).Interior.Color = SpreadsheetGear.Colors.LightBlue


        For i As Integer = 1 To XL_COLS + IIf(recap, 1, 0) ' Abs(recap)
            With oSheet.Range(Excel_Cell(1, XL_RH + XL_DESC + i)).EntireColumn
                If (recap And i = 1) Or (Not recap AndAlso RCOLS(i, 0) = "CAPTION") Then
                    .ColumnWidth = 20
                ElseIf (recap And i = 2) Or (Not recap AndAlso RCOLS(i, 0) = "CAPTIONSUB") Then
                    .ColumnWidth = 12
                Else
                    .ColumnWidth = 11
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    If CHKTHOUSANDS = "Y" Then
                        .NumberFormat = "###,##0.0;(###,##0.0)"
                    Else
                        .NumberFormat = "###,##0;(###,##0)"
                    End If
                End If
            End With
        Next i

        ' Add Spaces, Bolding, and Numeric Formatting

        Dim j As Integer = XL_CH + 1
        For i As Integer = 1 To XL_ROWS
            j = j + 1
            Dim LINE_TYPE As String = oSheet.Cells(Excel_Cell(j, XL_RH + 3)).Text
            If LINE_TYPE = "S" Or LINE_TYPE = "T" Then
                If recap Then
                    oSheet.Cells(Excel_Cell(j, XL_RH + XL_DESC + 1)).Font.Bold = True
                Else
                    oSheet.Cells(Excel_Cell(j, XL_RH + XL_DESC + XL_COL_CAPTION)).Font.Bold = True
                End If
            End If

            If LINE_TYPE = "P" Then
                For c As Integer = 1 To XL_COLS
                    If recap Then
                        oSheet.Cells(j - 1, XL_RH + XL_DESC + c - 1).NumberFormat = "###,##0.0;(###,##0.0)"
                        If c >= 3 Then oSheet.Cells(j - 1, XL_RH + XL_DESC + c - 1).Value = "" ' the %s on a recap sheet are sums on the code columns - need to be recalculated somehow - maybe when i have the onions to make this report work with formulas
                    Else
                        If (Mid$(RCOLS(c, 0), 1, 1) = "V" Or Mid$(RCOLS(c, 0), 1, 1) = "P") Then
                            oSheet.Cells(j - 1, XL_RH + XL_DESC + c - 1).Value = ""
                        Else
                            oSheet.Cells(j - 1, XL_RH + XL_DESC + c - 1).NumberFormat = "###,##0.0;(###,##0.0)"
                        End If
                    End If
                Next c
            End If

            Dim j_orig As Integer = j
            Dim LINE_SPACE_AFTER As String = oSheet.Cells(Excel_Cell(j, XL_RH + 4)).Text
            If LINE_SPACE_AFTER = "1" Then
                j = j + 1
                oSheet.Range(Excel_Cell(j, XL_RH + 1) & ":" & Excel_Cell(j, XL_RH + XL_DESC + XL_COLS)).Insert(SpreadsheetGear.InsertShiftDirection.Down)
            End If

            If LINE_TYPE = "T" Then
                oSheet.Range(Excel_Cell(j_orig, XL_RH + 1) & ":" & Excel_Cell(j_orig, XL_RH + XL_DESC + XL_COLS)).Interior.Color = SpreadsheetGear.Colors.LightGreen ' LightYellow
                oSheet.Range(Excel_Cell(j_orig, XL_RH + 1) & ":" & Excel_Cell(j_orig, XL_RH + XL_DESC + XL_COLS)).Font.Bold = True
            Else
                If LINE_TYPE = "S" Then
                    oSheet.Range(Excel_Cell(j_orig, XL_RH + 1) & ":" & Excel_Cell(j_orig, XL_RH + XL_DESC + XL_COLS)).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke ' LightGray
                End If
            End If
        Next i

        If XL_ROWS_TOTAL = 0 Then
            XL_ROWS_TOTAL = j
        End If

        ' Column Headings

        If recap Then
            For i As Integer = 3 To XL_COLS
                Dim z As String = tbl.Columns(XL_DESC + i - 1).ColumnName
                z = Mid(z, InStr(z, ":") + 1)
                oSheet.Cells(Excel_Cell(XL_CH - 2 + 0, XL_RH + XL_DESC + i)).Value = z
                oSheet.Cells(Excel_Cell(XL_CH - 2 + 1, XL_RH + XL_DESC + i)).Value = RCOLS(0, 1)
                oSheet.Cells(Excel_Cell(XL_CH - 2 + 2, XL_RH + XL_DESC + i)).Value = RCOLS(0, 2)
            Next i
        Else
            For i As Integer = 1 To XL_COLS
                If RCOLS(i, 0) <> "CAPTION" And RCOLS(i, 0) <> "CAPTIONSUB" Then
                    oSheet.Cells(Excel_Cell(XL_CH - 2 + 1, XL_RH + XL_DESC + i)).Value = RCOLS(i, 1)
                    oSheet.Cells(Excel_Cell(XL_CH - 2 + 2, XL_RH + XL_DESC + i)).Value = RCOLS(i, 2)
                End If
            Next i
        End If

        Sheet_Heading()

    End Sub
    Sub Load_Sheet(rowASTSRPT1_D As DataRow)

        ASCMAIN1.sql = "Select GLTPROF1.LINE_NO, GLTPROF1.LINE_TAG" & vbCrLf _
            & ", GLTPROF1.LINE_TYPE, GLTPROF1.LINE_SPACE_AFTER" & vbCrLf
        For i As Integer = 1 To RCOLS_count
            Dim EXP As String = RCOLS(i, 3)
            If COL_ORA.ContainsKey(RCOLS(i, 0)) Then
                EXP = COL_ORA(RCOLS(i, 0))
            End If
            ASCMAIN1.sql &= ", " & EXP & " as " & RCOLS(i, 0) & vbCrLf
        Next i
        ASCMAIN1.sql &= " From " & ASTSRPT1 & " ASTSRPT1, " & GLTPROF1 & " GLTPROF1" & vbCrLf
        ASCMAIN1.sql &= " Where GLTPROF1.LINE_TAG = ASTSRPT1.LINE_TAG" & vbCrLf
        ASCMAIN1.sql &= "   and NVL(GLTPROF1.LINE_SUPPRESS,'0') <> '1'" & vbCrLf
        For i As Integer = 1 To cfmax
            Dim z As String = "G" & CStr(i)
            ASCMAIN1.sql &= " and ASTSRPT1." & z & " = '" & rowASTSRPT1_D.Item(z) & "'" & vbCrLf
        Next i

        Load_DataTable(XL_CH, XL_RH + 1, "LINE_NO")
    End Sub

    Function Load_DataTable(Rx As Integer, Cx As Integer, Optional OrderBy As String = "") As DataTable

        Dim tbl As DataTable = ASCDATA1.GetDataTable

        If recap Then
            For Each dc As DataColumn In tbl.Columns
                If dc.ColumnName.StartsWith("'") And dc.ColumnName.EndsWith("'") Then
                    dc.ColumnName = Mid(dc.ColumnName, 2, dc.ColumnName.Length - 2)
                End If
            Next
        End If

        Dim dvw As DataView = tbl.DefaultView
        dvw.Sort = OrderBy
        tbl = dvw.ToTable

        range = oSheet.Range(Excel_Cell(Rx, Cx))
        range.CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.None)

        If recap Then
            oSheet.Range(Excel_Cell(Rx, Cx + 4)).Value = ""
            oSheet.Range(Excel_Cell(Rx, Cx + 5)).Value = ""

        Else
            oSheet.Range(Excel_Cell(Rx, Cx)).EntireRow.Clear()
        End If

        Return tbl
    End Function
    Sub Load_Sheet_Recap(rowASTSRPT1_D As DataRow)

        ASCMAIN1.Progress("-", "Recap")

        'Mid(SHEET_NAME, 1, 3) = Format$(SI - 1, "000")
        'Dim z As String = SHEET_NAME & " -Recap- " & COLUMN_CAPTIONs(cfi + 1 - 1) ' cf(2, cfi + 1)
        'If Len(z) > 31 Then
        '    z = Mid(z, 1, 31)
        'End If

        SI = SI + 1

        Mid(SHEET_NAME, 1, 3) = Format$(SI - 1, "000")
        Dim z As String = SHEET_NAME & " -Recap- " & COLUMN_CAPTIONs(cfi + 1 - 1) ' cf(2, cfi + 1)
        If Len(z) > 31 Then
            z = Mid(z, 1, 31)
        End If

        oSheet = oWB.Worksheets.Add
        oSheet.Name = z

        Dim CODES As String = ""
        ASCMAIN1.sql = "Select Distinct " & "G" & CStr(cfi + 1) & " from " & ASTSRPT1 & " ASTSRPT1" & vbCrLf _
            & " where G" & CStr(cfi + 1) & " Not Like '" & RCX & "%'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "G" & CStr(cfi + 1))
            CODES &= ",'" & row.Item(0) & "'"
        Next

        sql = ""
        If cfi > 0 Then
            For i As Integer = 1 To cfi
                If i <= cfi Then
                    z = "G" & CStr(i)
                    sql &= " and ASTSRPT1." & z & " = '" & rowASTSRPT1_D.Item(z) & "'"
                End If
            Next i
        End If

        sql &= " and ASTSRPT1.G" & CStr(cfi + 1) & " NOT LIKE '" & RCX & "%'"
        If cfi + 2 < cfmax Then '  If cfi + 2 <= cfmax   Then
            sql &= " and ASTSRPT1.G" & CStr(cfi + 2) & " LIKE '" & RCX & "%'"
        End If

        '& "  Sum(" & RCOLS(0, 0) & ")" & vbCrLf _
        '& "  Sum(Case when LINE_TYPE = 'P' then NULL else " & RCOLS(0, 0) & " End)" & vbCrLf _

        ASCMAIN1.sql = "Select * from" _
            & "(" _
            & "  Select GLTPROF1.LINE_NO, GLTPROF1.LINE_TAG" & vbCrLf _
            & ", GLTPROF1.LINE_TYPE, GLTPROF1.LINE_SPACE_AFTER" & vbCrLf _
            & ", GLTPROF1.LINE_CAPTION, GLTPROF1.LINE_CAPTION_SUB" & vbCrLf _
            & ", G" & CStr(cfi + 1) & vbCrLf _
            & ", " & RCOLS(0, 3) & " " & RCOLS(0, 0) & vbCrLf _
            & "  From " & ASTSRPT1 & " ASTSRPT1, " & GLTPROF1 & " GLTPROF1" & vbCrLf _
            & "  where GLTPROF1.LINE_TAG = ASTSRPT1.LINE_TAG" & vbCrLf _
            & "    and NVL(GLTPROF1.LINE_SUPPRESS,'0') <> '1'" & vbCrLf _
            & sql & vbCrLf _
            & ")" & vbCrLf _
            & " Pivot " & vbCrLf _
            & "(" & vbCrLf _
            & "  Sum(" & RCOLS(0, 0) & ")" & vbCrLf _
            & "  for " & "G" & CStr(cfi + 1) & vbCrLf _
            & "  in (" & Mid(CODES, 2) & ")" & vbCrLf _
            & ")"

        Dim tbl As DataTable = Load_DataTable(XL_CH, XL_RH + 1, "LINE_NO")

        XL_COLS_D = XL_COLS
        XL_COLS = tbl.Columns.Count - 4

        Format_Sheet(tbl)
        oSheet.Range(Excel_Cell(XL_CH + 1, XL_RH + XL_DESC + 2)).Select()

        XL_COLS = XL_COLS_D
        XL_COLS_D = tbl.Columns.Count - 4

    End Sub
    Sub Sheet_Heading(Optional rowASTSRPT1_D As DataRow = Nothing)

        Dim COFF As Integer
        If SI = 1 Then
            COFF = 4
        ElseIf recap Then
            COFF = 2
        Else
            COFF = XL_COL_CAPTION - 1
        End If

        With oSheet.Cells(Excel_Cell(1, XL_DESC + 1 + COFF))
            .Value = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        With oSheet.Cells(Excel_Cell(2, XL_DESC + 1 + COFF))
            .Value = "Profitability Report"
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With


        With oSheet.Cells(Excel_Cell(3, XL_DESC + 1 + COFF))
            .Value = COL_LAYOUT_DESC
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        If CHKTHOUSANDS = "Y" Then
            With oSheet.Cells(Excel_Cell(4, XL_DESC + 1 + COFF))
                .Value = "In $000's"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End With
        End If

        With oSheet.Cells(Excel_Cell(1, XL_DESC + 1))
            .Value = "Gen: " & Format(Now + ASCMAIN1.NowTSD, "MM/dd/yy HH:mm")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With oSheet.Cells(Excel_Cell(2, XL_DESC + 1))
            .Value = "As Of: " & RYPLEGEND
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With oSheet.Cells(Excel_Cell(4, XL_DESC + 1))
            .Value = oSheet.Name
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With oSheet.Cells(Excel_Cell(5, XL_DESC + 1))
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            If cfi > 0 Then
                Dim rowASTGROUP As DataRow = Nothing
                If rowASTSRPT1_D IsNot Nothing Then rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(rowASTSRPT1_D.Item("G" & CStr(cfi)))

                If rowASTGROUP Is Nothing Then
                    SHEET_ENTITY_DESC = ""
                Else
                    SHEET_ENTITY_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                End If
                .Value = SHEET_ENTITY_DESC
                .Font.Color = SpreadsheetGear.Colors.Blue ' System.Drawing.Color.Blue
            End If

        End With

        oSheet.Cells.Font.Name = "Times New Roman"

        ' Delete columns used to describe attributes of row
        For i As Integer = 1 To XL_DESC
            oSheet.Range(Excel_Cell(XL_CH + 1, XL_RH + 1)).EntireColumn.Delete()
        Next i

        If SI = 1 Then
            oSheet.Range(Excel_Cell(XL_CH + 1, XL_RH + 1 + cfmax + 1)).Select()
        ElseIf Not recap Then
            oSheet.Range(Excel_Cell(XL_CH + 1, 1)).Select()
        Else
            oSheet.Range(Excel_Cell(XL_CH + 1, 3)).Select()
        End If
        oSheet.WindowInfo.FreezePanes = True

        oSheet.Range("A1:A1").Select()
    End Sub
End Class