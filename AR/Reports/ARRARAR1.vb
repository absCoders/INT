Public Class ARRARAR1

    Dim GLTCREC3 As String

    Dim chkInd As Boolean

    Private xRYP0_legend As String
    Private xRYP1_legend As String
    Private xRYP0 As String
    Private xRYP1 As String

    Private pRYP0 As String
    Private pRYP1 As String
    Private sqlc As String
    Private sqlwc As String
    Private sqlwo As String
    Dim fc As Integer
    Dim fi(6) As String
    Dim fa(6) As Decimal
    Dim z0 As String


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -12, 0, 0)
        'Set_cmbYP("RYP1", ASCMAIN1.CYP, -12, 0, 0)

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        chkInd = Absx1.chkFor("CHKIND").Checked

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        sqlw = ""
        'RYP
        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
        'LYP
        xRYP1_legend = Absx1.cmbFor("RYP1").Value
        xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
        'LYP1
        pRYP0 = ASCMAIN1.Period_Calc(xRYP0, -1)
        'RYP1
        pRYP1 = ASCMAIN1.Period_Calc(xRYP1, -1)


        If xRYP0 = xRYP1 Then
            SUBT = "A/R Activity in " & xRYP0_legend
        Else
            SUBT = "A/R Activity From " & xRYP0_legend & " To " & xRYP1_legend
        End If
        sqlw &= " where GLTCREC3.OPS_YYYYPP between '" & pRYP0 & "' and '" & xRYP1 & "'"
        sqlw &= " and GLTCREC3.CREC_TYPE_CODE = 'AR'"

        Prepare_dst(True, New String() {sqlw})

        Check_if_Empty("ARTARAR1")

    End Sub

    Overrides Sub Print_Report()
        Dim chkIP As String = ""
        If chkInd Then
            CHKIP = "1"
        Else
            CHKIP = "0"
        End If
        CR_params.Add("CHKIND", CHKIP)
        Generate_Report(RPT, "A/R Activity Report", SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

            GLTCREC3 = ASCMAIN1.Temp_Table("Select GLTCREC3.OPS_YYYYPP, GLTCREC3.DETL_CVX_NO CUST_CODE, SUM(CREC_AMT) AMT " _
                                            & " from GLTCREC3 where ROWNUM < 1 Group By GLTCREC3.OPS_YYYYPP, GLTCREC3.DETL_CVX_NO")
            'Create_TDA(.Tables.Add("ARTCRDR1"), ARTSTMT1, "*", 0, False, , 2)


            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTSTMT1.OPS_YYYYPP from ARTCUST1, ARTSTMT1 where ROWNUM < 1 "
            Create_TDA(.Tables.Add, "ARTARAR1", "**", 0, False, , 2)
            With dst.Tables("ARTARAR1").Columns
                .Add("AR_LYP", GetType(System.Double))
                .Add("AR_RYP", GetType(System.Double))
                .Add("INV_NEW", GetType(System.Double))
                .Add("CRM_NEW", GetType(System.Double))
                .Add("DRM_NEW", GetType(System.Double))
                .Add("ONA_NEW", GetType(System.Double))
                .Add("CHB_NEW", GetType(System.Double))
                .Add("RTN_NEW", GetType(System.Double))
                .Add("INV_OFF", GetType(System.Double))
                .Add("CRM_OFF", GetType(System.Double))
                .Add("DRM_OFF", GetType(System.Double))
                .Add("ONA_OFF", GetType(System.Double))
                .Add("CHB_OFF", GetType(System.Double))
                .Add("RTN_OFF", GetType(System.Double))
                .Add("CASH", GetType(System.Double))
                .Add("PMTS", GetType(System.Double))
                .Add("DISC", GetType(System.Double))
                .Add("WOFF", GetType(System.Double))
            End With

            For Each TABLE_NAME As String In New String() _
                {"ARTCUST1"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

        End With

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTCREC3)
        ASCDATA1.ExecuteSQL("Insert into " & GLTCREC3 _
                            & " Select GLTCREC3.OPS_YYYYPP, GLTCREC3.DETL_CVX_NO CUST_CODE, SUM(CREC_AMT) AMT from GLTCREC3 " & sqlw _
                            & " Group By GLTCREC3.OPS_YYYYPP, GLTCREC3.DETL_CVX_NO ")


        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ARTARAR1"}
            Fill_Records(TABLE_NAME)
        Next

        Dim tbl As DataTable
        'Opening A/R
        z0 = "1"
        Build_SQL(z0, "ARTCUST9", "OPS_YYYYPP")
        ASCMAIN1.sql = sqlc & " SUM(AMT) AMT " & vbCrLf _
        & " from ARTCUST1, " & GLTCREC3 & " ARTCUST9" & sqlwc & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 1
            fi(1) = "AR_LYP"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If
        z0 = ""

        'Closing A/R
        If xRYP0 = ASCMAIN1.CYP Then
            Build_SQL(z0, "ARTOPEN1", "")
            ASCMAIN1.sql = sqlc & "SUM (INV_BALANCE)" _
            & " from ARTOPEN1,ARTCUST1 " & sqlwc & sqlwo
            tbl = ASCDATA1.GetDataTable
            If tbl.Rows.Count <> 0 Then
                fc = 1
                fi(1) = "AR_RYP"
                For Each row As DataRow In tbl.Rows
                    Build_ARTARAR1(row)
                Next
            End If
        End If
        If xRYP0 <> ASCMAIN1.CYP Or xRYP0 <> xRYP1 Then
            Build_SQL(z0, "ARTCUST9", "OPS_YYYYPP")
            ASCMAIN1.sql = sqlc & "SUM (AMT)" _
            & " from ARTCUST1, " & GLTCREC3 & " ARTCUST9" & sqlwc & sqlwo
            tbl = ASCDATA1.GetDataTable
            If tbl.Rows.Count <> 0 Then
                fc = 1
                fi(1) = "AR_RYP"
                For Each row As DataRow In tbl.Rows
                    Build_ARTARAR1(row)
                Next
            End If
        End If

        'New Shipments & Returns
        Build_SQL(z0, "SOTINVH1", "ORDR_YYYYPP_UPDATED")
        ASCMAIN1.sql = sqlc & "SUM (DECODE (INV_TYPE, 'I', INV_TOTAL_AMOUNT,0))" _
        & ", SUM (DECODE (INV_TYPE, 'C', INV_TOTAL_AMOUNT,'D',INV_TOTAL_AMOUNT,0)) " _
        & " from SOTINVH1,ARTCUST1 " & sqlwc & sqlwo
        ' EMP EXCLUDES ORDR_TYPE_CODE = 'P' - MAYBE WE SHOULD HAVE A FLAG IN SOTTYPE1 WHERE WE SAY AR OR AP
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 2
            fi(1) = "INV_NEW"
            fi(2) = "RTN_NEW"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'New Chargebacks & On/Accounts
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        Dim sqlA As String
        sqlA = " and NVL(CHARGEBACK_IND,'0') = '1'" _
        & " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " and ARTPYMT2.CUST_CODE is Not Null"
        ASCMAIN1.sql = sqlc & " SUM (DECODE (ABS(GL_DIST_AMT),GL_DIST_AMT,GL_DIST_AMT,0)) " _
        & ", SUM (DECODE (ABS(GL_DIST_AMT),GL_DIST_AMT,0,GL_DIST_AMT))" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 2
            fi(1) = "CHB_NEW"
            fi(2) = "ONA_NEW"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'Deductions
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        sqlA = " and (CHARGEBACK_IND is Null or CHARGEBACK_IND <> -1)" _
        & " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " and ARTPYMT2.CUST_CODE is Not Null" _
        & " and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1')"
        ASCMAIN1.sql = sqlc & " SUM (GL_DIST_AMT) " _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 1
            fi(1) = "WOFF"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'Payments and Write-Offs
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        sqlA = " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " and (ARTPYMT2.PYMT_DELETED IS NULL OR ARTPYMT2.PYMT_DELETED <> '1')" _
        & " and ARTPYMT2.CUST_CODE IS NOT NULL "
        ASCMAIN1.sql = sqlc & " Sum (INV_PMT) " _
        & ", Sum (INV_DISC_TAKEN) " _
        & ", Sum (INV_WRITE_OFF) " _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 3
            fi(1) = "PMTS"
            fi(2) = "DISC"
            fi(3) = "WOFF"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'Direct Write-Offs to G/L
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        sqlA = " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " and ARTPYMT2.CUST_CODE is Not Null" _
        & " and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1')"
        ASCMAIN1.sql = sqlc & " SUM (GL_DIST_AMT) " _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT4, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 1
            fi(1) = "WOFF"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'Cash
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        sqlA = " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT2.CUST_CODE is Not Null" _
        & " and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1')"
        ASCMAIN1.sql = sqlc & " SUM (CUST_PYMT_AMT) " _
        & " from ARTPYMT1, ARTPYMT2, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 1
            fi(1) = "CASH"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        'Paid Items
        Build_SQL(z0, "ARTPYMT2", "OPS_YYYYPP")
        sqlA = " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " and ARTPYMT2.CUST_CODE is Not Null"
        ASCMAIN1.sql = sqlc & " Sum (Decode (INV_TYPE, 'I', INV_PMT, 0)) " _
        & ", Sum (Decode (INV_TYPE, 'C', INV_PMT, 0)) " _
        & ", Sum (Decode (INV_TYPE, 'D', INV_PMT, 0)) " _
        & ", Sum (Decode (INV_TYPE, 'O', INV_PMT, 0)) " _
        & ", Sum (Decode (INV_TYPE, 'R', INV_PMT, 0)) " _
        & ", Sum (Decode (INV_TYPE, 'B', INV_PMT, 0)) " _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1 " & sqlwc & sqlA & sqlwo
        tbl = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            fc = 6
            fi(1) = "INV_OFF"
            fi(2) = "CRM_OFF"
            fi(3) = "DRM_OFF"
            fi(4) = "ONA_OFF"
            fi(5) = "RTN_OFF"
            fi(6) = "CHB_OFF"
            For Each row As DataRow In tbl.Rows
                Build_ARTARAR1(row)
            Next
        End If

        EnforceConstraints(True)
    End Sub

    Private Sub AbsCheckBox0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'chkHISTORY = AbsCheckBox0.Checked
        'grpPERIOD_RANGE.Visible = chkHISTORY
        'grpREMITVIA.Visible = chkHISTORY
    End Sub

    Sub Build_SQL(ByVal z As String, ByVal z1 As String, ByVal z2 As String)
        Dim zc As String = ""
        sqlc = ""
        sqlwc = ""
        sqlwo = ""
        z0 = z

        If z1 = "SOTINVH1" Or z1 = "ARTCRDR1" Then
            If z1 = "SOTINVH1" Then
                zc = "Decode(" & z1 & ".CUST_BILL_TO_CUST,Null," & z1 & ".CUST_CODE," & z1 & ".CUST_BILL_TO_CUST)"
            Else
                zc = "Decode(ARTCUST1.CUST_BILL_TO_CUST,Null,ARTCUST1.CUST_CODE,ARTCUST1.CUST_BILL_TO_CUST)"
            End If
        ElseIf z1 = "SOTINVHC" Then
            zc = "CONT_PYMT_PAYEE_CODE"
        Else
            zc = "ARTCUST1.CUST_CODE"
        End If
        If z2 = "" Then
            z2 = "'" & xRYP0 & "'"
        End If
        sqlc = "Select " & zc & ", " & z2 & ","
        If z1 = "SOTINVHC" Then
            sqlwc = " where ARTCUST1.CUST_CODE = SOTINVHC.CONT_PYMT_PAYEE_CODE "
        Else
            sqlwc = " where ARTCUST1.CUST_CODE = " & z1 & ".CUST_CODE"
        End If
        If z2 <> "" And z2 <> "'" & xRYP0 & "'" Then
            If z0 = "1" Then
                If Not chkInd Then
                    sqlwc = sqlwc & " and " & z2 & " = '" & pRYP0 & "'"
                Else
                    If pRYP1 = pRYP0 Then
                        sqlwc = sqlwc & " and " & z2 & " = '" & pRYP1 & "'"
                    Else
                        sqlwc = sqlwc & " and " & z2 & " >= '" & pRYP0 & "'"
                        sqlwc = sqlwc & " and " & z2 & " <= '" & pRYP1 & "'"
                    End If
                End If
            Else
                If Not chkInd And z1 = "ARTCUST9" Then
                    sqlwc = sqlwc & " and OPS_YYYYPP = '" & xRYP0 & "'"
                Else
                    If xRYP0 = xRYP1 Then
                        sqlwc = sqlwc & " and " & z2 & " = '" & xRYP0 & "'"
                    Else
                        sqlwc = sqlwc & " and " & z2 & " >= '" & xRYP1 & "'"
                        sqlwc = sqlwc & " and " & z2 & " <= '" & xRYP0 & "'"
                    End If
                End If
            End If
        End If

        sqlwc &= SQL_in("CUST_CODE", "ARTCUST1.CUST_CODE")
        sqlwc &= SQL_in("COLLECTOR_CODE", "ARTCUST1.COLLECTOR_CODE")

        sqlwo = " group by " & zc & ", " & z2

    End Sub

    Sub Build_ARTARAR1(ByVal row As DataRow)
        Dim cust As String = row.Item(0) & ""
        Dim YP As String = ""
        Dim i As Integer
        If chkInd Then
            YP = row.Item("OPS_YYYYPP") & ""
            If z0 = "1" Then
                YP = ASCMAIN1.Period_Calc(YP, 1)
            End If
        Else
            YP = "000000"
        End If
        fa(0) = 0
        For i = 1 To fc
            fa(I) = Val(row.Item(I + 1) & "")
            fa(0) = fa(0) + System.Math.Abs(fa(i))
        Next i


        Dim rowARTARAR1 As DataRow = dst.Tables("ARTARAR1").Rows.Find _
    (New Object() {row.Item(0), YP})
        If rowARTARAR1 Is Nothing Then
            rowARTARAR1 = dst.Tables("ARTARAR1").NewRow
            With rowARTARAR1
                'ADD ROW
                .Item("CUST_CODE") = row.Item(0)
                .Item("OPS_YYYYPP") = YP
                For i = 1 To fc
                    .Item(fi(i)) = Val(.Item(fi(i)) & "") + Val(fa(i) & "")
                Next i
            End With
            dst.Tables("ARTARAR1").Rows.Add(rowARTARAR1)
        Else
            'EDIT ROW
            With rowARTARAR1
                For i = 1 To fc
                    .Item(fi(i)) = Val(.Item(fi(i)) & "") + Val(fa(i) & "")
                Next i
            End With
            dst.Tables("ARTARAR1").AcceptChanges()
        End If

    End Sub

End Class