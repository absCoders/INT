Public Class DPCMAIN1


    Public Shared Function Set_sqlYP(RYP As String, sqlDPTMRPGO As String)

        Dim sqlYP As String = sqlDPTMRPGO

        Dim RYP_diff As Integer = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYP) + 1
        Dim mx As String = Format(RYP_diff, "00")
        sqlYP = Replace(sqlYP, $"QTY_00 EOM", $"QTY_{mx} EOM")
        sqlYP = Replace(sqlYP, $"QTY_00 POS", $"QTY_{mx} POS")
        sqlYP = Replace(sqlYP, $"'000000' OPS_YYYYPP", $"'{RYP}' OPS_YYYYPP")
        sqlYP = Replace(sqlYP, $"OPS_YYYYPP = '000000'", $"OPS_YYYYPP = '{RYP}'")

        'For m As Integer = 25 To 1 Step -1
        '    Dim m0 As String = Format(m, "00")
        '    Dim mx As String = Format(m + X, "00")
        '    sqlYP = Replace(sqlYP, $"QTY_{m0}", $"QTY_{mx}")
        'Next

        Return sqlYP
    End Function


    Public Shared Sub Create_Worktables_DPTMRPGO(initialize As Boolean, frm As ASFBASE0, ByRef sqlDPTMRPGO As String, ByRef DPTMRPGO As String, RYP_end As String, chkShowAllMonths As Boolean)


        If initialize Then

            sqlDPTMRPGO = "SELECT '000000' OPS_YYYYPP, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_COST_STD
, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_STATUS
, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO
, ICTITEM1.ITEM_POS_MAX, ICTITEM1.ITEM_POS_MIN
, 0 FC
, DEM00, DEM01, DEM02, DEM03, DEM04, DEM05, DEM06, DEM07, DEM08, DEM09, DEM10, DEM11, DEM12
, DEM13, DEM14, DEM15, DEM16, DEM17, DEM18, DEM19, DEM20, DEM21, DEM22, DEM23, DEM24, DEM25
, OPO00, OPO01, OPO02, OPO03, OPO04, OPO05, OPO06, OPO07, OPO08, OPO09, OPO10, OPO11, OPO12
, OPO13, OPO14, OPO15, OPO16, OPO17, OPO18, OPO19, OPO20, OPO21, OPO22, OPO23, OPO24, OPO25
, OPP00, OPP01, OPP02, OPP03, OPP04, OPP05, OPP06, OPP07, OPP08, OPP09, OPP10, OPP11, OPP12
, OPP13, OPP14, OPP15, OPP16, OPP17, OPP18, OPP19, OPP20, OPP21, OPP22, OPP23, OPP24, OPP25
, EOM.EOM, EOM.EOM * ICTITEM1.ITEM_COST_STD EOM_EXT_COST, POS.POS
 FROM ICTITEM1, ICTCOLL1
, (Select ITEM_CODE, QTY_00 EOM from DPTMRPG1 WHERE MRP_TYPE = '5') EOM
, (Select ITEM_CODE, QTY_00 POS from DPTMRPG1 WHERE MRP_TYPE = '6') POS
, (Select ITEM_CODE, QTY_00 DEM00
, QTY_01 DEM01, QTY_02 DEM02, QTY_03 DEM03, QTY_04 DEM04, QTY_05 DEM05, QTY_06 DEM06
, QTY_07 DEM07, QTY_08 DEM08, QTY_09 DEM09, QTY_10 DEM10, QTY_11 DEM11, QTY_12 DEM12
, QTY_13 DEM13, QTY_14 DEM14, QTY_15 DEM15, QTY_16 DEM16, QTY_17 DEM17, QTY_18 DEM18
, QTY_19 DEM19, QTY_20 DEM20, QTY_21 DEM21, QTY_22 DEM22, QTY_23 DEM23, QTY_24 DEM24, QTY_25 DEM25
 from DPTMRPG1 WHERE MRP_TYPE = '1') FCS
, (Select ITEM_CODE, QTY_00 OPO00
, QTY_01 OPO01, QTY_02 OPO02, QTY_03 OPO03, QTY_04 OPO04, QTY_05 OPO05, QTY_06 OPO06
, QTY_07 OPO07, QTY_08 OPO08, QTY_09 OPO09, QTY_10 OPO10, QTY_11 OPO11, QTY_12 OPO12
, QTY_13 OPO13, QTY_14 OPO14, QTY_15 OPO15, QTY_16 OPO16, QTY_17 OPO17, QTY_18 OPO18
, QTY_19 OPO19, QTY_20 OPO20, QTY_21 OPO21, QTY_22 OPO22, QTY_23 OPO23, QTY_24 OPO24, QTY_25 OPO25
 from DPTMRPG1 WHERE MRP_TYPE = '3') OPO
, (Select ITEM_CODE, QTY_00 OPP00
, QTY_01 OPP01, QTY_02 OPP02, QTY_03 OPP03, QTY_04 OPP04, QTY_05 OPP05, QTY_06 OPP06
, QTY_07 OPP07, QTY_08 OPP08, QTY_09 OPP09, QTY_10 OPP10, QTY_11 OPP11, QTY_12 OPP12
, QTY_13 OPP13, QTY_14 OPP14, QTY_15 OPP15, QTY_16 OPP16, QTY_17 OPP17, QTY_18 OPP18
, QTY_19 OPP19, QTY_20 OPP20, QTY_21 OPP21, QTY_22 OPP22, QTY_23 OPP23, QTY_24 OPP24, QTY_25 OPP25
 from DPTMRPG1 WHERE MRP_TYPE = '4') OPP
where EOM.ITEM_CODE = ICTITEM1.ITEM_CODE
  AND POS.ITEM_CODE = ICTITEM1.ITEM_CODE
  AND FCS.ITEM_CODE = ICTITEM1.ITEM_CODE
  AND OPO.ITEM_CODE = ICTITEM1.ITEM_CODE
  AND OPP.ITEM_CODE = ICTITEM1.ITEM_CODE
  AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"

            sqlDPTMRPGO = "SELECT X.*
, ST1.QTY_BEG, ST1.QTY_BEG * X.ITEM_COST_STD AMT_BEG
, ST1.QTY_REC, ST1.QTY_REC * X.ITEM_COST_STD AMT_REC
, ST1.QTY_RTN, ST1.QTY_RTN * X.ITEM_COST_STD AMT_RTN
, ST1.QTY_ADJ, ST1.QTY_ADJ * X.ITEM_COST_STD AMT_ADJ
FROM (
SELECT ITEM_CODE
, SUM (WHSE_QTY_BEG) QTY_BEG
, SUM (WHSE_QTY_REC) QTY_REC
, SUM (WHSE_QTY_RTN) QTY_RTN
, SUM (WHSE_QTY_ADJ) QTY_ADJ
FROM ICTSTAT1 WHERE OPS_YYYYPP = '000000' GROUP BY ITEM_CODE
) ST1, (" & sqlDPTMRPGO & ") X WHERE ST1.ITEM_CODE (+) = X.ITEM_CODE"
            DPTMRPGO = ASCMAIN1.Temp_Table(sqlDPTMRPGO)

            ASCMAIN1.sql = $"Select * from {DPTMRPGO}"
            frm.Create_TDA(frm.dst.Tables.Add, "DPTMRPGO", "**", 0, False, "", 2)
            With frm.dst.Tables("DPTMRPGO").Columns
                .Add("OVER_QTY", GetType(System.Int64), "ISNULL(EOM,0)-ISNULL(FC,0)")
                .Add("OVER_EXT_COST", GetType(System.Decimal), "ISNULL(OVER_QTY,0)*ISNULL(ITEM_COST_STD,0)")
                .Add("FC_CUR", GetType(System.Int64))
                .Add("FC_FUT", GetType(System.Int64))
                .Add("PO_CUR", GetType(System.Int64))
                .Add("PO_FUT", GetType(System.Int64))
                .Add("PP_CUR", GetType(System.Int64))
                .Add("PP_FUT", GetType(System.Int64))
                .Add("FC_EXT_COST", GetType(System.Decimal), "ISNULL(FC,0)*ISNULL(ITEM_COST_STD,0)")
                .Add("FCTM", GetType(System.Int64))
                .Add("FCTM_EXT_COST", GetType(System.Decimal), "ISNULL(FCTM,0)*ISNULL(ITEM_COST_STD,0)")
                .Add("POTM", GetType(System.Int64))
                .Add("POTM_EXT_COST", GetType(System.Decimal), "ISNULL(POTM,0)*ISNULL(ITEM_COST_STD,0)")
                .Add("PPTM", GetType(System.Int64))
                .Add("PPTM_EXT_COST", GetType(System.Decimal), "ISNULL(PPTM,0)*ISNULL(ITEM_COST_STD,0)")

                .Add("ZERO", GetType(System.String), "IIF(ISNULL(EOM,0)=0 AND ISNULL(FC_CUR,0)=0 AND ISNULL(FC_FUT,0)=0 AND ISNULL(PO_CUR,0)=0 AND ISNULL(PO_FUT,0)=0 AND ISNULL(OVER_QTY,0)=0 AND ISNULL(FC,0)=0,'0','1')")
            End With



        Else
            Dim mIndex As Integer = 1
            If chkShowAllMonths Then

                Dim yp As String = ASCMAIN1.CYP
                Do While yp <= RYP_end

                    ASCMAIN1.sql = $"Truncate Table {DPTMRPGO}"
                    ASCDATA1.ExecuteSQL()

                    Dim sqlYP As String = Set_sqlYP(yp, sqlDPTMRPGO)
                    ASCMAIN1.sql = $"Insert into {DPTMRPGO} " & sqlYP
                    ASCDATA1.ExecuteSQL()
                    Fill_DPTMRPGO(yp, frm, mIndex)

                    yp = ASCMAIN1.Period_Calc(yp, 1)
                    mIndex += 1
                Loop

            Else
                mIndex = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYP_end) + 1

                ASCMAIN1.sql = $"Truncate Table {DPTMRPGO}"
                ASCDATA1.ExecuteSQL()

                Dim sqlYP As String = Set_sqlYP(RYP_end, sqlDPTMRPGO)
                ASCMAIN1.sql = $"Insert into {DPTMRPGO} " & sqlYP
                ASCDATA1.ExecuteSQL()
                Fill_DPTMRPGO(RYP_end, frm, mIndex)
            End If

        End If
    End Sub

    Public Shared Sub Fill_DPTMRPGO(RYPx As String, frm As ASFBASE0, mIndex As Integer)

        Dim X As Integer = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYPx) + 1

        frm.Fill_Records("DPTMRPGO", , False)

        For Each rowDPTMRPGO As DataRow In frm.dst.Tables("DPTMRPGO").Select($"OPS_YYYYPP = '{RYPx}'")
            Dim ITEM_CODE As String = rowDPTMRPGO.Item("ITEM_CODE") & ""
            Dim ITEM_POS_MAX As Decimal = Val(rowDPTMRPGO.Item("ITEM_POS_MAX") & "")
            If ITEM_POS_MAX = 0 Then
                rowDPTMRPGO.Item("FC") = 0
            Else
                Dim FC As Int32 = 0 ' Val(rowDPTMRPGO.Item($"DEM{Format(0, "00")}") & "") ' PD
                Dim P As Decimal = Math.Truncate(ITEM_POS_MAX)
                'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "BN001A01" Then Stop
                If P >= 1 Then
                    For I As Integer = 1 To P
                        If I + X < 25 Then
                            FC += Val(rowDPTMRPGO.Item($"DEM{Format(I + X, "00")}") & "")
                        End If
                    Next
                End If
                'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH017A06" Then Stop
                If ITEM_POS_MAX - P > 0 Then
                    If P + X + 1 < 25 Then
                        FC += Val(rowDPTMRPGO.Item($"DEM{Format(P + X + 1, "00")}") & "") * (ITEM_POS_MAX - P)
                    End If
                End If
                rowDPTMRPGO.Item("FC") = FC
            End If

            Dim FC_CUR As Int32 = 0 ' Val(rowDPTMRPGO.Item($"DEM{Format(0, "00")}") & "") ' PD
            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH017A06" Then Stop
            Dim FC_FUT As Int32 = 0
            For I As Integer = 0 To 24
                Dim DEM As Int32 = Val(rowDPTMRPGO.Item($"DEM{Format(I, "00")}") & "")
                If I <= X Then
                    FC_CUR += DEM
                Else
                    FC_FUT += DEM
                    'Debug.Print(DEM)
                End If
            Next
            rowDPTMRPGO.Item("FC_CUR") = FC_CUR
            rowDPTMRPGO.Item("FC_FUT") = FC_FUT

            Dim PO_CUR As Int32 = 0
            Dim PO_FUT As Int32 = 0
            For I As Integer = 0 To 24
                Dim OPO As Int32 = Val(rowDPTMRPGO.Item($"OPO{Format(I, "00")}") & "")
                If I <= X Then
                    PO_CUR += OPO
                Else
                    PO_FUT += OPO
                End If
            Next
            rowDPTMRPGO.Item("PO_CUR") = PO_CUR
            rowDPTMRPGO.Item("PO_FUT") = PO_FUT

            Dim PP_CUR As Int32 = 0
            Dim PP_FUT As Int32 = 0
            For I As Integer = 0 To 24
                Dim OPP As Int32 = Val(rowDPTMRPGO.Item($"OPP{Format(I, "00")}") & "")
                If I <= X Then
                    PP_CUR += OPP
                Else
                    PP_FUT += OPP
                End If
            Next
            rowDPTMRPGO.Item("PP_CUR") = PP_CUR
            rowDPTMRPGO.Item("PP_FUT") = PP_FUT

            rowDPTMRPGO.Item("FCTM") = Val(rowDPTMRPGO.Item($"DEM{Format(mIndex, "00")}") & "")
            rowDPTMRPGO.Item("POTM") = Val(rowDPTMRPGO.Item($"OPO{Format(mIndex, "00")}") & "")
            rowDPTMRPGO.Item("PPTM") = Val(rowDPTMRPGO.Item($"OPP{Format(mIndex, "00")}") & "")
            If mIndex = 1 Then
                rowDPTMRPGO.Item("FCTM") += Val(rowDPTMRPGO.Item($"DEM00") & "")
                rowDPTMRPGO.Item("POTM") += Val(rowDPTMRPGO.Item($"OPO00") & "")
                rowDPTMRPGO.Item("PPTM") += Val(rowDPTMRPGO.Item($"OPP00") & "")
            End If
        Next



    End Sub


End Class
