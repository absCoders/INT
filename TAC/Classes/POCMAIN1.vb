Public Class POCMAIN1

    Public Shared Sub Production_Commit(
        S As Integer, PO_ORDER_NO As String, Optional PO_ORDER_LNO As Integer = 0, Optional VEND_WHSE_CODE As String = "")

        Dim sqlw As String = ""
        If PO_ORDER_NO <> "" Then
            sqlw = " and POTORDR9.PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            If PO_ORDER_LNO <> 0 Then
                sqlw &= " and POTORDR9.PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            End If
        End If

        If sqlw = "" And S = -1 Then
            Stop
            ' truncate potordr9
            ' delete entire col from ictsat2
            ' exit sub
        End If

        Dim sql9 As String = "Select POTORDR9.*,POTORDR1.VEND_WHSE_CODE from POTORDR9,POTORDR1 where POTORDR1.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" & sqlw
        If VEND_WHSE_CODE <> "" Then
            sql9 = $"Select POTORDR9.*,'{VEND_WHSE_CODE}' VEND_WHSE_CODE from POTORDR9" & ASCMAIN1.SQL_Add_WHERE(sqlw)
        End If

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & sql9 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 Set WHSE_QTY_COMM = NVL(WHSE_QTY_COMM,0) " & IIf(S = 1, "+", "-") & "1 * NVL(R1.PO_QTY_COM,0)" & vbCrLf _
            & "    where WHSE_CODE = R1.VEND_WHSE_CODE and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into ICTSTAT2 (WHSE_CODE,ITEM_CODE,WHSE_QTY_COMM)" & vbCrLf _
            & "     Values (R1.VEND_WHSE_CODE,R1.ITEM_CODE," & IIf(S = 1, "+", "-") & "1 * NVL(R1.PO_QTY_COM,0));" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;" & vbCrLf
        ASCDATA1.ExecuteSQL()

        If S = -1 Then
            ASCDATA1.ExecuteSQL("Delete from POTORDR9" & ASCMAIN1.SQL_Add_WHERE(sqlw))
        End If

    End Sub

    Public Shared Sub ICTSTAT2_PO(S As Integer, PO_ORDER_NO As String)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select POTORDR2.WHSE_CODE, POTORDR2.ITEM_CODE, SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & "   from POTORDR2" & vbCrLf _
            & "   where POTORDR2.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
            & "   group by POTORDR2.WHSE_CODE, POTORDR2.ITEM_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 Set WHSE_QTY_ONPO = NVL(WHSE_QTY_ONPO,0) " & IIf(S = 1, "+", "-") & "1 * NVL(R1.PO_QTY_OPN,0)" & vbCrLf _
            & "    where WHSE_CODE = R1.WHSE_CODE and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into ICTSTAT2 (WHSE_CODE,ITEM_CODE,WHSE_QTY_ONPO)" & vbCrLf _
            & "     Values (R1.WHSE_CODE,R1.ITEM_CODE," & IIf(S = 1, "+", "-") & "1 * NVL(R1.PO_QTY_OPN,0));" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;" & vbCrLf
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Shared Sub Update_POTORDR9( _
                                     frm As ASFBASE0, _
                                     PO_ORDER_NO As String, _
                                     VEND_WHSE_CODE As String, _
                                     Optional write_plan As Boolean = False, _
                                     Optional write_POTORDR2 As Boolean = True, _
                                     Optional PO_ORDER_LNO_only As Integer = 0, _
                                     Optional P_Index As Integer = 0)

        ' Record PO Commitments

        Dim rowPOTORDR1 As DataRow = frm.dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)
        If rowPOTORDR1 Is Nothing Then
            rowPOTORDR1 = frm.Fill_Record("POTORDR1", PO_ORDER_NO)
        End If

        frm.dst.Tables("POTORDR9").Rows.Clear()
        'Stop ' WHOS TO SAY THAT POTORDR2 IS READY FOR THIS?
        Dim sql As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        If PO_ORDER_LNO_only <> 0 Then sql &= " and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO_only)

        For Each rowPOTORDR2 As DataRow In frm.dst.Tables("POTORDR2").Select(sql)
            Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
            Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE")
            Dim PO_QTY_OPN As Int64 = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
            If PO_QTY_OPN <> 0 Then
                Dim std_or_cur As String
                Dim BM_ISSUE_NO As String
                If rowPOTORDR2.Item("BM_ISSUE_SEL") & "" = "1" Then
                    std_or_cur = "C"
                    BM_ISSUE_NO = ""
                Else
                    std_or_cur = ""
                    BM_ISSUE_NO = rowPOTORDR2.Item("BM_ISSUE_NO") & ""
                End If
                If std_or_cur <> "" Or BM_ISSUE_NO <> "" Then

                    If ASCMAIN1.Running_in_VS AndAlso PO_ORDER_NO = "500726" Then Stop

                    Dim TBL As DataTable = TAC.POCMAIN1.Get_BM(frm, ITEM_CODE, std_or_cur, BM_ISSUE_NO, _
                             False, True, "C", PO_QTY_OPN, VEND_WHSE_CODE, "HOPCA")
                    For Each rowBMTMAIN3 As DataRow In TBL.Select("")

                        Dim rowPOTORDR9 As DataRow = frm.dst.Tables("POTORDR9").NewRow
                        rowPOTORDR9.Item("PO_ORDER_NO") = PO_ORDER_NO
                        rowPOTORDR9.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        rowPOTORDR9.Item("ITEM_CODE") = rowBMTMAIN3.Item("BM_COMP_ITEM")
                        rowPOTORDR9.Item("PO_QTY_COM") = rowBMTMAIN3.Item("QTY_COM")
                        frm.dst.Tables("POTORDR9").Rows.Add(rowPOTORDR9)

                        If write_plan Then
                            If frm.dst.Tables("DPTMUPDS").Rows.Find(New Object() {rowBMTMAIN3.Item("BM_COMP_ITEM"), rowPOTORDR2.Item("PO_DATE_COMPSDUE"), rowPOTORDR1.Item("VEND_WHSE_CODE")}) Is Nothing Then
                                frm.dst.Tables("DPTMUPDS").Rows.Add(New Object() {rowBMTMAIN3.Item("BM_COMP_ITEM"), rowPOTORDR2.Item("PO_DATE_COMPSDUE"), rowPOTORDR1.Item("VEND_WHSE_CODE")})
                            End If

                            Dim rowDPTMUPD0 As DataRow = frm.dst.Tables("DPTMUPD0").NewRow
                            With rowDPTMUPD0
                                .Item("ITEM_CODE") = rowBMTMAIN3.Item("BM_COMP_ITEM")
                                .Item("SD") = "C"
                                .Item("BM_ISSUE_NO") = BM_ISSUE_NO
                                .Item("DATE_REQ") = rowPOTORDR2.Item("PO_DATE_COMPSDUE")
                                ' .Item("QTY_PLN") = q
                                .Item("QTY_REQ") = rowBMTMAIN3.Item("QTY_COM")
                                .Item("PO_ORDER_NO") = PO_ORDER_NO
                                .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")
                                .Item("WHSE_CODE") = rowPOTORDR1.Item("VEND_WHSE_CODE")
                                .Item("P_INDEX") = P_Index ' Calc_P_Index(rowPOTORDR2.Item("PO_DATE_COMPSDUE"))
                            End With
                            frm.dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)
                        End If
                    Next

                    If (BM_ISSUE_NO <> "" And rowPOTORDR2.Item("BM_ISSUE_NO") & "" <> "") And _
                          BM_ISSUE_NO <> rowPOTORDR2.Item("BM_ISSUE_NO") & "" Then
                        'ASCDATA1.ExecuteSQL("Update POTORDR2 Set BM_ISSUE_NO = '" & BM_ISSUE_NO & "' where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO))
                        rowPOTORDR2.Item("BM_ISSUE_NO") = BM_ISSUE_NO
                        If write_POTORDR2 Then
                            frm.Update_Record_TDA("POTORDR2")
                        End If
                    End If
                End If
            End If
        Next
        frm.Update_Record_TDA("POTORDR9")
        TAC.POCMAIN1.Production_Commit(1, PO_ORDER_NO, PO_ORDER_LNO_only)
    End Sub

    Public Shared Function Get_BM(
        frmASFBASE0 As ASFBASE0,
        BM_PROD_ITEM As String,
        std_or_cur As String,
        ByRef BM_ISSUE_NO As String,
        include_VSM As Boolean,
        include_When_Exhausted As Boolean,
        cost_sfr As String,
        prod_qty As Int64,
        VEND_WHSE_CODE As String,
        ava_type As String) As DataTable

        ' when prod_qty <> 0: 1) be sure to populate prod_whse, 2) set xr_ONLY = ""
        ' ava_type: "H" = On Hand Only; "HO" = OnHand + On Order; "HC" = On Hand - Qty Com; "HOP" = On hand + On Order + Planned

        Dim tblBMTMAIN3 As DataTable

        ASCMAIN1.sql = "Select BMTMAIN3.*" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_ON_HAND" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_ONPO" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_PLAN" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_COMM" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_WASTE_PCT" & vbCrLf _
            & ", ICTITEM1.ITEM_PLAN_WASTE_PCT" & vbCrLf _
            & ", ICTCOSTC.ITEM_COST_TOTAL" & vbCrLf _
            & " from BMTMAIN3,BMTMAIN1,ICTSTAT2,ICTITEM1,ICTCOSTC" & vbCrLf _
            & " where BMTMAIN3.BM_PROD_ITEM = '" & BM_PROD_ITEM & "'" & vbCrLf _
            & "   and BMTMAIN1.BM_PROD_ITEM = BMTMAIN3.BM_PROD_ITEM" & vbCrLf _
            & "   and ICTSTAT2.WHSE_CODE (+) = '" & VEND_WHSE_CODE & "'" & vbCrLf _
            & "   and ICTSTAT2.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM" & vbCrLf _
            & "   and ICTCOSTC.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM"
        If Not include_VSM Then
            ASCMAIN1.sql &= "  and (NVL(BMTMAIN3.BM_VEND_SUPP_MATL,'0') <> '1')" & vbCrLf
        End If
        If Not include_When_Exhausted Then
            ASCMAIN1.sql &= "  and BMTMAIN3.BM_WHEN_EXHAUSTED is Null" & vbCrLf
        End If
        If BM_ISSUE_NO <> "" Then
            ASCMAIN1.sql &= "  and BMTMAIN3.BM_ISSUE_NO = '" & BM_ISSUE_NO & "'" & vbCrLf
        Else
            If std_or_cur = "S" Then
                ASCMAIN1.sql &= "  and BMTMAIN3.BM_ISSUE_NO = BMTMAIN1.BM_ISSUE_STD" & vbCrLf
            Else
                ASCMAIN1.sql &= "  and TO_NUMBER(BMTMAIN3.BM_ISSUE_NO) = BMTMAIN1.BM_ISSUE_COUNTER" & vbCrLf
            End If
        End If

        tblBMTMAIN3 = ASCDATA1.GetDataTable

        'If ASCMAIN1.Running_in_VS Then
        '    If BM_PROD_ITEM = "KS001V28USA" Then Stop
        'End If

        Dim F As String = ""
        If ava_type.Contains("H") Then F &= "+ISNULL(WHSE_QTY_ON_HAND,0)"
        If ava_type.Contains("O") Then F &= "+ISNULL(WHSE_QTY_ONPO,0)"
        If ava_type.Contains("P") Then F &= "+ISNULL(WHSE_QTY_PLAN,0)"
        If ava_type.Contains("C") Then F &= "-ISNULL(WHSE_QTY_OPEN,0)"
        If ava_type.Contains("A") Then F &= "-ISNULL(WHSE_QTY_PICK,0)"
        If ava_type.Contains("C") Then F &= "-ISNULL(WHSE_QTY_COMM,0)"
        tblBMTMAIN3.Columns.Add("QTY_AVA", GetType(System.Int64), Mid(F, 2))
        F = IIf(cost_sfr = "S", "ITEM_COST_WASTE_PCT", "ITEM_PLAN_WASTE_PCT")
        tblBMTMAIN3.Columns.Add("WASTE_PCT", GetType(System.Decimal), "ISNULL(" & F & ",0)")
        tblBMTMAIN3.Columns.Add("QTY_REQ", GetType(System.Int64), "0.499999 + " & CStr(prod_qty) & " * ISNULL(BM_QTY_PER_ASSY,0) * (1 + WASTE_PCT / 100)")
        tblBMTMAIN3.Columns.Add("PCT_AVA", GetType(System.Decimal), "IIF(QTY_REQ=0,0,QTY_AVA/QTY_REQ)")
        tblBMTMAIN3.Columns.Add("QTY_COM", GetType(System.Int64))

        If BM_ISSUE_NO = "" And tblBMTMAIN3.Rows.Count <> 0 Then
            BM_ISSUE_NO = tblBMTMAIN3.Rows(0).Item("BM_ISSUE_NO") & ""
        End If

        Dim XR As New List(Of String)

        For Each rowBMTMAIN3 As DataRow In tblBMTMAIN3.Select("", "BM_SEQ")
            Dim BM_WHEN_EXHAUSTED As String = rowBMTMAIN3.Item("BM_WHEN_EXHAUSTED") & ""
            If BM_WHEN_EXHAUSTED <> "" And Not XR.Contains(BM_WHEN_EXHAUSTED) Then
                XR.Add(BM_WHEN_EXHAUSTED)
            End If
            rowBMTMAIN3.Item("QTY_COM") = rowBMTMAIN3.Item("QTY_REQ")
        Next

        If XR.Count <> 0 Then
            For Each BM_WHEN_EXHAUSTED As String In XR
                Dim sqlw As String = "BM_WHEN_EXHAUSTED = '" & BM_WHEN_EXHAUSTED & "'"
                Dim PCT As Decimal = Val(tblBMTMAIN3.Compute("MIN(PCT_AVA)", sqlw) & "")
                If PCT > 1 Then PCT = 1
                If PCT < 0 Then PCT = 0
                For Each rowBMTMAIN3 As DataRow In tblBMTMAIN3.Select(sqlw)
                    Dim QTY_COM As Int64 = (0.499999 + Val(rowBMTMAIN3.Item("QTY_REQ") * PCT))
                    Dim QTY_AVA As Int64 = Val(rowBMTMAIN3.Item("QTY_AVA") & "")
                    If QTY_COM > QTY_AVA Then QTY_COM = QTY_AVA
                    rowBMTMAIN3.Item("QTY_COM") = QTY_COM
                Next
                sqlw = Replace(sqlw, "BM_WHEN_EXHAUSTED", "BM_REPLACE_WITH")
                PCT = 1 - PCT
                For Each rowBMTMAIN3 As DataRow In tblBMTMAIN3.Select(sqlw)
                    Dim QTY_COM As Int64 = 0.5 + Val(rowBMTMAIN3.Item("QTY_REQ") * PCT)
                    rowBMTMAIN3.Item("QTY_COM") = QTY_COM
                Next
            Next
        End If

        Return tblBMTMAIN3
    End Function


    Public Shared Function GetUserSignature(rowASTUSER1 As DataRow) As String

        Dim strHtml As String = ""

        strHtml &= $"<div><strong>{rowASTUSER1.Item("USER_NAME")}</strong></div>"
        strHtml &= $"<div>{rowASTUSER1.Item("USER_TITLE")}</div>"
        Dim USER_TELEPHONE As String = rowASTUSER1.Item("USER_TELEPHONE") & ""
        If USER_TELEPHONE <> "" Then USER_TELEPHONE = Mid(USER_TELEPHONE, 1, 3) & "." & Mid(USER_TELEPHONE, 4, 3) & "." & Mid(USER_TELEPHONE, 7, 4)
        Dim USER_EXT As String = rowASTUSER1.Item("USER_EXT") & ""
        If USER_EXT <> "" And USER_TELEPHONE <> "" Then USER_TELEPHONE &= " x " & USER_EXT
        Dim USER_FAX As String = rowASTUSER1.Item("USER_FAX") & ""
        If USER_FAX <> "" Then USER_FAX = Mid(USER_FAX, 1, 3) & "." & Mid(USER_FAX, 4, 3) & "." & Mid(USER_FAX, 7, 4)

        If USER_TELEPHONE <> "" Then strHtml &= $"<div><strong>T:</strong> {USER_TELEPHONE}</div>"
        If USER_FAX <> "" Then strHtml &= $"<div><strong>F:</strong> {USER_FAX}</div>"

        Return strHtml

    End Function
End Class