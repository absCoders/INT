Imports System.Text
Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Shared.Styles
Imports Infragistics.Win.UltraWinGrid

Public Class POFTRACK

    Private POTORDRX As String = String.Empty

    Dim VEND_CODE As String

    Dim sqlPOTORDRX As String
    Dim ems As List(Of ASCNOTE1)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")

        Create_WorkTable(True)

        With dst

            ASCMAIN1.sql = "Select POTORDR1.VEND_CODE" & vbCrLf _
                & ", Count (*) POS" & vbCrLf _
                & " , Sum (POTORDRX.PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
                & " , Sum (POTORDRX.PO_QTY_REC) PO_QTY_REC" & vbCrLf _
                & " , Sum (POTORDRX.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
                & " , Sum (POTORDRX.PO_AMT_ORD) PO_AMT_ORD" & vbCrLf _
                & " , Sum (POTORDRX.PO_AMT_REC) PO_AMT_REC" & vbCrLf _
                & " , Sum (POTORDRX.PO_AMT_OPN) PO_AMT_OPN" & vbCrLf _
                & $" from POTORDR1, {POTORDRX} POTORDRX" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDRX.PO_ORDER_NO" & vbCrLf _
                & " group by POTORDR1.VEND_CODE"
            ASCMAIN1.sql = $"Select X.*, APTVEND1.VEND_NAME from ({ASCMAIN1.sql}) X, APTVEND1 where APTVEND1.VEND_CODE = X.VEND_CODE"
            Create_TDA(.Tables.Add, "POTORDRV", "**", 0, False,, 1)
            With .Tables("POTORDRV")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & ", POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST PO_AMT_ORD, POTORDR2.PO_QTY_REC * POTORDR2.PO_COST PO_AMT_REC" & vbCrLf _
                & ", POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST PO_AMT_OPN" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.VEND_NAME, POTORDR1.VEND_ACCT_NO" & vbCrLf _
                & ", POTORDR1.VEND_CONF_NO, POTORDR1.VEND_CONF_DATE" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE" & vbCrLf _
                & ", POTORDR1.PO_DATE_ORDERED" & vbCrLf _
                & $" from POTORDR1,POTORDR2,{POTORDRX} POTORDRX" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO And POTORDRX.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDRD", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select POTORDR1.*" & vbCrLf _
            & ", POTORDRX.PO_QTY_ORD, POTORDRX.PO_QTY_REC, POTORDRX.PO_QTY_OPN, POTORDRX.PO_AMT_ORD, POTORDRX.PO_AMT_REC, POTORDRX.PO_AMT_OPN" & vbCrLf _
            & $" from POTORDR1, {POTORDRX} POTORDRX" & vbCrLf _
            & " where POTORDR1.VEND_CODE = :PARM1 And POTORDR1.PO_ORDER_NO = POTORDRX.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, True, "V", , "VEND_CONF_DATE")

            ASCMAIN1.sql = "Select POTORDR1.*" & vbCrLf _
            & ", POTORDRX.PO_QTY_ORD, POTORDRX.PO_QTY_REC, POTORDRX.PO_QTY_OPN, POTORDRX.PO_AMT_ORD, POTORDRX.PO_AMT_REC, POTORDRX.PO_AMT_OPN" & vbCrLf _
            & $" from POTORDR1, {POTORDRX} POTORDRX" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDRX.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDRA", "**", 0, False,, 1)

            ASCMAIN1.sql = "Select POTORDR2.* from POTORDR2 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V",, "PO_DATE_ETD,PO_DATE_ETD_NOTES,PO_DATE_REQUIRED")

            ASCMAIN1.sql = "Select POTORDR2.* from POTORDR2 where PO_ORDER_NO = :PARM1 and PO_ORDER_LNO = :PARM2"
            Create_TDA(.Tables.Add, "POTORDR2_REFRESH", "**", 0, False, "VN")

            ASCMAIN1.sql = "Select POTORDR8.* from POTORDR8 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR8", "**", 0, True, "V")

            Create_Relation("POTORDR2", "POTORDR8", "PO_ORDER_NO,PO_ORDER_LNO")

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTIREC1.RECEIPT_DATE, ICTIREC1.SOURCE_DOC_NO, ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER" & vbCrLf _
            & " from ICTIREC2,ICTIREC1" & vbCrLf _
            & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO  And ICTIREC2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "V")

            Create_Relation("POTORDR2", "ICTIRECX", "PO_ORDER_NO,PO_ORDER_LNO")

            ASCMAIN1.sql = "SELECT X.*, X.AMT_REC - X.AMT_INV VAR_AMT FROM (
                SELECT ICTIREC1.VEND_CODE, POTORDR1.PO_ORDER_NO, APTINVH1.INV_NUM
                , COUNT (*) RECS
                , SUM (ICTIREC2.QTY_REC) QTY_REC, SUM (APTINVH5.INV_QTY) INV_QTY
                , SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) AMT_REC
                , SUM (APTINVH5.INV_QTY * APTINVH5.INV_COST) AMT_INV
                FROM APTINVH5,APTINVH1,ICTIREC2, ICTIREC1,POTORDR1
                WHERE APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO
                AND ICTIREC1.RECEIPT_NO= APTINVH5.RECEIPT_NO
                AND ICTIREC2.RECEIPT_NO= APTINVH5.RECEIPT_NO
                AND ICTIREC2.RECEIPT_LNO= APTINVH5.RECEIPT_LNO
                AND APTINVH1.OPS_YYYYPP = :PARM1
                AND POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO
                AND APTINVH5.INV_COST <> ICTIREC2.PO_COST
                AND APTINVH1.INV_STATUS = 'P'
                group by ICTIREC1.VEND_CODE, POTORDR1.PO_ORDER_NO, APTINVH1.INV_NUM) X
                "
            Create_TDA(.Tables.Add, "APTINVHS", "**", 0, False, "V", 0)


            ASCMAIN1.sql = "SELECT X.*, X.AMT_REC - X.AMT_INV VAR_AMT FROM (
                SELECT ICTIREC1.VEND_CODE, POTORDR1.PO_ORDER_NO, APTINVH1.INV_NUM, ICTIREC2.ITEM_CODE, ICTIREC2.PO_COST, APTINVH5.INV_COST
                , COUNT (*) RECS
                , SUM (ICTIREC2.QTY_REC) QTY_REC, SUM (APTINVH5.INV_QTY) INV_QTY
                , SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) AMT_REC
                , SUM (APTINVH5.INV_QTY * APTINVH5.INV_COST) AMT_INV
                FROM APTINVH5,APTINVH1,ICTIREC2, ICTIREC1,POTORDR1
                WHERE APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO
                AND ICTIREC1.RECEIPT_NO= APTINVH5.RECEIPT_NO
                AND ICTIREC2.RECEIPT_NO= APTINVH5.RECEIPT_NO
                AND ICTIREC2.RECEIPT_LNO= APTINVH5.RECEIPT_LNO
                AND APTINVH1.OPS_YYYYPP = :PARM1
                AND POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO
                AND APTINVH5.INV_COST <> ICTIREC2.PO_COST
                AND APTINVH1.INV_STATUS = 'P'
                GROUP BY  ICTIREC1.VEND_CODE, POTORDR1.PO_ORDER_NO, APTINVH1.INV_NUM
                , ICTIREC2.ITEM_CODE, ICTIREC2.PO_COST, APTINVH5.INV_COST) X
                "
            Create_TDA(.Tables.Add, "APTINVHD", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0, False)
            Fill_Records("ICTWHSE1")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            ASCMAIN1.sql = "Select * from SOTCARR1 where carrier_code in ('DHL', 'FEDEX', 'UPS')"
            Fill_Records("SOTCARR1", String.Empty, True, ASCMAIN1.sql)


            ASCMAIN1.sql = "Select WHSE_CODE CODE_VALUE, WHSE_DESC DESC_VALUE, '1' SEL from ICTWHSE1 where WHSE_STATUS = 'A'"
            Create_TDA(.Tables.Add, "ICTWHSES", "**", 0, False)
            Fill_Records("ICTWHSES", String.Empty, True, ASCMAIN1.sql)
            .Tables("ICTWHSES").Columns("SEL").DefaultValue = "0"
        End With

        grdAPTINVHS.DataSource = dst.Tables("APTINVHS")
        grdAPTINVHD.DataSource = dst.Tables("APTINVHD")

        grdICTWHSES.DataSource = dst.Tables("ICTWHSES")
        Sort_grdColumns(grdICTWHSES, "CODE_VALUE")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)
        cbeYP_Invoiced.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP_Invoiced.SelectedItem = cbeYP_Invoiced.Items(0)



        grdPOTORDRV.DataSource = dst.Tables("POTORDRV")
        grdPOTORDRD.DataSource = dst.Tables("POTORDRD")

        grdPOTORDRA.DataSource = dst.Tables("POTORDRA")

        grdPOTORDR1.DataSource = dst.Tables("POTORDR1")
        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")

        Create_Summary(grdAPTINVHS, "INV_NUM", "Count")
        Create_Summary(grdAPTINVHS, New String() {"RECS", "QTY_REC", "INV_QTY", "AMT_REC", "AMT_INV", "VAR_AMT"})

        Create_Summary(grdAPTINVHD, "ITEM_CODE", "Count")
        Create_Summary(grdAPTINVHD, New String() {"RECS", "QTY_REC", "INV_QTY", "AMT_REC", "AMT_INV", "VAR_AMT"})

        Create_Summary(grdPOTORDRV, "VEND_CODE", "Count")
        Create_Summary(grdPOTORDRV, "SEL")
        Create_Summary(grdPOTORDRV, New String() {"POS", "PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN", "PO_AMT_ORD", "PO_AMT_REC", "PO_AMT_OPN"})

        Create_Summary(grdPOTORDRD, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRD, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN", "PO_AMT_ORD", "PO_AMT_REC", "PO_AMT_OPN"})

        Create_Summary(grdPOTORDR1, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDR1, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN", "PO_AMT_ORD", "PO_AMT_REC", "PO_AMT_OPN"})

        Create_Summary(grdPOTORDRA, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRA, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN", "PO_AMT_ORD", "PO_AMT_REC", "PO_AMT_OPN"})

        Create_Summary(grdPOTORDR2, "PO_ORDER_LNO", "Count")
        Create_Summary(grdPOTORDR2, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN"})

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTORDRV, grdPOTORDR1, grdPOTORDRD, grdPOTORDRA}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            If grd.Name = "grdPOTORDRV" Then
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    c.CellActivation = Activation.NoEdit
                    If grd.Name = "grdPOTORDRV" And c.Key = "SEL" Then
                        c.CellActivation = Activation.AllowEdit
                    End If
                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If c.Key.Contains("PO_QTY_") Then c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    If c.Key.Contains("PO_AMT_") Then c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    If c.Key = "POS" Then c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                Next
            End With

            If grd.Name = "grdPOTORDR1" Then
                With grd.DisplayLayout.Bands(0)
                    .Columns("PO_ORDER_NO").Header.Fixed = True
                    .Columns("PO_DATE_ORDERED").Header.Fixed = True
                End With
            End If

            If grd.Name = "grdPOTORDRD" Then
                With grd.DisplayLayout.Bands(0)
                    .Columns("PO_ORDER_NO").Header.Fixed = True
                    .Columns("PO_ORDER_LNO").Header.Fixed = True
                    .Columns("ITEM_CODE").Header.Fixed = True
                    .Columns("ITEM_DESC").Header.Fixed = True
                End With
            End If
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTORDR2, grdAPTINVHS, grdAPTINVHD}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    c.CellActivation = Activation.NoEdit
                Next
            End With

            If grd.Name = "grdPOTORDR2" Then
                With grd.DisplayLayout.Bands(0)
                    .Columns("PO_ORDER_LNO").Header.Fixed = True
                    .Columns("ITEM_CODE").Header.Fixed = True
                End With
            End If
            If grd.Name = "grdPOTORDR2" Then
                grd.DisplayLayout.ViewStyleBand = ViewStyleBand.Horizontal
                'grd.Rows.ExpandAll(True)
            End If
        Next

        Show_Filter(grdPOTORDRV)
        Show_Filter(grdPOTORDRD)

        Show_Filter(grdAPTINVHS)
        Show_Filter(grdAPTINVHD)

        With chtICTIRECD
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            .EnableCrossHair = True

            .ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With


        ASCMAIN1.Add_Value_List(grdPOTORDR1, "PO_ORDER_TYPE")
        ASCMAIN1.Add_Value_List(grdPOTORDR2, "PO_STATUS")
        ASCMAIN1.Add_Value_List(grdPOTORDRD, "PO_STATUS")

        splPOTORDRD.Panel2Collapsed = True

        MakeTransparent(chkShowAll)

        MakeTransparent(chkToDC)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""
        Dim zMsg As String = String.Empty

        Select Case eItemKey

            Case "Load"

                If EMsg = "" Then
                    If MyBase.Absx1.txtFor("VEND_CODE").Text.Trim.Length > 0 Then
                        If Not Validate_Code("VEND_CODE", False, False) Then
                            EMsg &= vbCr & "You must provide a valid Vendor Code."
                            Exit Select
                        Else
                            VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                            If Not ASCMAIN1.Logical_Lock("POTORDR1", VEND_CODE) Then
                                Exit Sub
                            End If
                        End If
                    Else
                        EMsg &= vbCr & "You must provide a valid Vendor Code."
                    End If

                End If

            Case "Done"

            Case "Status Request"
                If dst.Tables("POTORDRV").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Vendors Selected for the Status Request Report"
                End If

                If EMsg = "" Then
                    If MsgBox($"OK to generate Status Request Reports & emails" & vbCrLf _
                              & $" for all Open POs for {dst.Tables("POTORDRV").Select("SEL = '1'").Length } Vendors", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Upload XLS"


                If EMsg = "" Then
                    Dim openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an XLSX file to Import"
                    openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        Dim FILENAME As String = openFileDialog1.FileName
                        If FILENAME <> "" And FILENAME.ToUpper.EndsWith(".XLSX") Then
                            Import_Status_Response(FILENAME)
                            If FILENAME = "" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Summary()

            Case "Status Request"
                EntryMode = "S"
                Status_Request()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1.Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                ' .Items("Done").Settings.Enabled = iScreenMode
                .Items("Done").Visible = ScreenMode And Not (EntryMode = "E")
                .Items("Update").Visible = ScreenMode And (EntryMode = "E")
                .Items("Cancel").Visible = ScreenMode And (EntryMode = "E")

                .Items("Refresh").Visible = Not ScreenMode
                .Items("Status Request").Visible = False ' Not ScreenMode

                .Items("Upload XLS").Visible = ScreenMode And (EntryMode = "E")

            End With

            UltraExplorerBar1.Groups("PO Summary").Visible = Not ScreenMode
            UltraExplorerBar1.Groups("PO Scope").Visible = False ' Not ScreenMode
        End If

        tabSummary.SelectedTab = tabSummary.Tabs(0)

        tabSummary.Visible = Not tf
        splPOTORDRV.Visible = tf

        MyBase.Set_Read_Only(grpHeader, tf)

        If ScreenMode Then
            grdPOTORDR1.Parent = splPOTORDR1.Panel1
            grdPOTORDR2.Parent = splPOTORDR1.Panel2

            grdPOTORDRD.Parent = tabPOTORDRV.Tabs("Items").TabPage

            If (EntryMode = "E") Then
                grdPOTORDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                With grdPOTORDR1.DisplayLayout.Bands(0)
                    For Each C As String In New String() {"PO_ORDR_NOTES_EXTERNAL", "VEND_CONF_NO", "VEND_CONF_DATE", "PO_DATE_REQUIRED", "PO_DATE_CANCEL"}
                        .Columns(C).CellActivation = Activation.AllowEdit
                        .Columns(C).CellAppearance.BackColor = System.Drawing.Color.LightBlue
                    Next
                End With

                grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                With grdPOTORDR2.DisplayLayout.Bands(0)
                    For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
                        .Columns(C).CellActivation = Activation.AllowEdit
                        .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Yellow
                    Next
                End With

                grdPOTORDR2.Rows.CollapseAll(True)
                ' grdPOTORDR2.DisplayLayout.ViewStyle = ViewStyle.SingleBand

                grdPOTORDRD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                With grdPOTORDRD.DisplayLayout.Bands(0)
                    For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
                        .Columns(C).CellActivation = Activation.AllowEdit
                        .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Yellow
                    Next
                End With
            End If

        Else
            grdPOTORDR1.Parent = splPOTORDRV_POS.Panel1
            grdPOTORDR2.Parent = splPOTORDRV_POS.Panel2

            'grdPOTORDR2.DisplayLayout.Bands(1).Hidden = False
            'grdPOTORDR2.DisplayLayout.ViewStyle = ViewStyle.SingleBand

            grdPOTORDRD.Parent = splPOTORDRD.Panel1

            grdPOTORDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdPOTORDR1.DisplayLayout.Bands(0)
                For Each C As String In New String() {"PO_ORDR_NOTES_EXTERNAL", "VEND_CONF_NO", "VEND_CONF_DATE", "PO_DATE_REQUIRED", "PO_DATE_CANCEL"}
                    .Columns(C).CellActivation = Activation.NoEdit
                    .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Empty
                Next
            End With

            grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdPOTORDR2.DisplayLayout.Bands(0)
                For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
                    .Columns(C).CellActivation = Activation.NoEdit
                    .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Empty
                Next

            End With

            grdPOTORDRD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdPOTORDRD.DisplayLayout.Bands(0)
                For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
                    .Columns(C).CellActivation = Activation.NoEdit
                    .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Empty
                Next

            End With

            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("VEND_CODE").Clear()
        Absx1.txtFor("VEND_NAME").Clear()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTORDRV", "POTORDR1", "POTORDR2", "POTORDR8", "ICTIRECX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        MyBase.EnforceConstraints(True)

        Absx1.txtFor("VEND_CODE").Focus()

        Refresh_Summary()
    End Sub

    Sub Load_Record()
        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        'Create_WorkTable(False)
        Refresh_Summary()

        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating POs")
        Dim PO_PARM_PINV_LT As Integer = Val(ROWs("POTPARM1").Item("PO_PARM_PINV_LT") & "")

        Try
            BeginTrans()
            If EntryMode = "C" Then

            ElseIf EntryMode = "E" Then

                dst.Tables("POTORDR8").Rows.Clear()

                For Each row As DataRow In dst.Tables("POTORDR2").Select("")
                    Dim PO_DATE_ETD As Date = Nothing
                    If row.Item("PO_DATE_ETD") & "" <> "" Then
                        PO_DATE_ETD = row.Item("PO_DATE_ETD") & ""
                    End If
                    Dim PO_DATE_ETD_NOTES As String = row.Item("PO_DATE_ETD_NOTES") & ""
                    Dim PO_DATE_REQUIRED As Date = row.Item("PO_DATE_REQUIRED") & ""

                    If row.Item("PO_DATE_ETD") & "" <> "" Then 'If PO_DATE_ETD <> "" Then
                        If VEND_CODE = "IPSA" Then
                            PO_DATE_REQUIRED = Format(CDate(PO_DATE_ETD).AddDays(PO_PARM_PINV_LT), "MM/dd/yyyy") '.Cells(r, PO_ORDER_NO_c + 15).Text ' T Date Shipped
                            row.Item("PO_DATE_REQUIRED") = PO_DATE_REQUIRED
                        End If

                    End If

                    Dim PO_DATE_ETD_ORIG As Date = Nothing
                    If row.Item("PO_DATE_ETD", DataRowVersion.Original) & "" <> "" Then
                        PO_DATE_ETD_ORIG = row.Item("PO_DATE_ETD", DataRowVersion.Original) & ""
                    End If
                    Dim PO_DATE_ETD_NOTES_ORIG As String = row.Item("PO_DATE_ETD_NOTES", DataRowVersion.Original) & ""
                    Dim PO_DATE_REQUIRED_ORIG As Date = Nothing
                    If row.Item("PO_DATE_REQUIRED", DataRowVersion.Original) & "" <> "" Then
                        PO_DATE_REQUIRED_ORIG = row.Item("PO_DATE_REQUIRED", DataRowVersion.Original) & ""
                    End If

                    If Format(PO_DATE_ETD, "yyyyMMdd") <> Format(PO_DATE_ETD_ORIG, "yyyyMMdd") _
                    Or PO_DATE_ETD_NOTES <> PO_DATE_ETD_NOTES_ORIG _
                    Or Format(PO_DATE_REQUIRED, "yyyyMMdd") <> Format(PO_DATE_REQUIRED_ORIG, "yyyyMMdd") Then
                        Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                        Dim PO_ORDER_LNO As Int32 = Val(row.Item("PO_ORDER_LNO") & "")

                        Dim rowPOTORDR8 As DataRow = dst.Tables("POTORDR8").NewRow
                        With rowPOTORDR8
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("PO_DATE_ETD") = PO_DATE_ETD
                            .Item("PO_DATE_ETD_NOTES") = PO_DATE_ETD_NOTES
                            .Item("PO_DATE_REQUIRED") = PO_DATE_REQUIRED
                        End With
                        dst.Tables("POTORDR8").Rows.Add(rowPOTORDR8)
                    End If
                Next

                For Each rowPO_ORDER_NO As DataRow In ASCDATA1.SelectDistinct("POTORDR8", New String() {"PO_ORDER_NO"}).Select("")
                    Dim PO_ORDER_NO As String = rowPO_ORDER_NO.Item("PO_ORDER_NO")
                    Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)
                    rowPOTORDR1.Item("VEND_CONF_DATE") = DATETIME_STAMP
                Next

                Update_Record_TDA("POTORDR1")
                Update_Record_TDA("POTORDR2")
                Update_Record_TDA("POTORDR8")
            End If

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
        ByVal ctl As Control, ByVal COLUMN_NAME As String,
        Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "VEND_CODE"
                'sql_where &= "VEND_CODE IN (SELECT VEND_CODE FROM POTVENDZ)"
                sql_where &= "VEND_CODE IN (SELECT Distinct VEND_CODE FROM POTORDR1)"
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTORDRV, "SSBBB", "Show Filter", "Show GroupBox", "Hide POs", "Select All", "De-Select All")
        'Load_Popup_Menu(grdPOTORDRD, "SSBBBBPBBBPB", "Show Filter", "Show GroupBox", "Cancel Open Qty", "Re-Open PO", "PO Inquiry", "NPI Order Processing", "Track UPS", "Track FedEx", "Track DHL", "Copy ETA Date")
        Load_Popup_Menu(grdPOTORDRD, "SSBBPB", "Show Filter", "Show GroupBox", "PO Inquiry")
        Load_Popup_Menu(grdPOTORDR1, "SSB", "Show Filter", "Show GroupBox", "PO Inquiry", "Hide PO Details")
        Load_Popup_Menu(grdPOTORDRA, "SSB", "Show Filter", "Show GroupBox", "PO Inquiry")
        Load_Popup_Menu(grdPOTORDR2, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

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
            Select Case e.SourceControl.Name

            End Select
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = Nothing ' DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Hide POs"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splPOTORDRV_PO.Panel2Collapsed = (tlb_sbt.Checked)

            Case "Hide PO Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splPOTORDRV_POS.Panel2Collapsed = (tlb_sbt.Checked)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("Load", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

        End Select
    End Sub

#End Region

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "VEND_CODE" And Not ScreenMode Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Click_Command("Load", e)
            End If
        End If
    End Sub

    Sub Create_WorkTable(initialize As Boolean)

        '                & IIf(optPROD.Value = "*", "", $"   And POTORDR1.PO_TYPE = '{optPROD.Value}'") & vbCrLf _
        ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR1.VEND_CODE" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_REC) PO_QTY_REC" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST) PO_AMT_ORD" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_REC * POTORDR2.PO_COST) PO_AMT_REC" & vbCrLf _
                & " , Sum(POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) PO_AMT_OPN" & vbCrLf _
                & " from POTORDR1, POTORDR2" & vbCrLf _
                & " Where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.PO_STATUS = 'O'" & vbCrLf _
                & " group by POTORDR2.PO_ORDER_NO, POTORDR1.VEND_CODE"
        sqlPOTORDRX = ASCMAIN1.sql

        If initialize Then

            Dim sqlw As String = ""

            POTORDRX = ASCMAIN1.Temp_Table($"Select X.* from ({ASCMAIN1.sql}) X where ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"Alter Table {POTORDRX} Add Primary Key (PO_ORDER_NO)")

        Else
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Creating Worktables")

            ASCMAIN1.sql = sqlPOTORDRX

            Dim JAN_01_TY As String = $"01-JAN-{Mid(ASCMAIN1.CYP, 1, 4)}"
            Dim JAN_01_LY As String = $"01-JAN-{Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) - 1, "0000")}"


            If EntryMode = "E" Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, " Where ", $" Where POTORDR2.PO_STATUS = 'O' and ")
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, " Where ", $" Where NVL(POTORDR1.VEND_CODE,'?') = '{VEND_CODE}' and ")

            ElseIf EntryMode = "S" Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, " Where ", $" Where POTORDR2.PO_STATUS = 'O' and ")
            Else
                If optPOs.Value = "Y" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, " and POTORDR1.PO_STATUS_CODE = 'O'", $" and POTORDR1.INIT_DATE > '{JAN_01_TY}'")
                ElseIf optPOs.Value = "L" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, " and POTORDR1.PO_STATUS_CODE = 'O'", $" and POTORDR1.INIT_DATE > '{JAN_01_LY}' and POTORDR1.INIT_DATE < '{JAN_01_TY}'")
                End If
            End If

            ASCDATA1.ExecuteSQL($"Truncate Table {POTORDRX}")
            ASCDATA1.ExecuteSQL($"Insert into {POTORDRX} {ASCMAIN1.sql }")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub


    Private Sub grdPOTORDRV_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTORDRV.AfterRowActivate
        Setup_Vendor()
    End Sub

    Sub Setup_Vendor()

        grdPOTORDR2.Visible = False

        If grdPOTORDRV.ActiveRow Is Nothing OrElse Not grdPOTORDRV.ActiveRow.IsDataRow Then
            grdPOTORDR1.Visible = False
        Else
            Dim VEND_CODE As String = grdPOTORDRV.ActiveRow.Cells("VEND_CODE").Value

            Fill_Records("POTORDR1", VEND_CODE)
            Sort_grdColumns(grdPOTORDR1, "PO_ORDER_NO")


            grdPOTORDR1.Visible = True
            grdPOTORDR1.Text = optPOs.Text & " - " & VEND_CODE
        End If

    End Sub

    Private Sub grdPOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTORDR1.AfterRowActivate
        Setup_PO()
    End Sub

    Sub Setup_PO()
        If grdPOTORDR1.ActiveRow Is Nothing OrElse Not grdPOTORDR1.ActiveRow.IsDataRow Then
            grdPOTORDR2.Visible = False
        Else
            Dim PO_ORDER_NO As String = grdPOTORDR1.ActiveRow.Cells("PO_ORDER_NO").Value

            If EntryMode = "E" Then
                Dim dvw As DataView = dst.Tables("POTORDR2").DefaultView
                dvw.RowFilter = $"PO_ORDER_NO = '{PO_ORDER_NO}'"
            Else
                EnforceConstraints(False)
                Fill_Records("POTORDR2", PO_ORDER_NO)
                Fill_Records("POTORDR8", PO_ORDER_NO)
                Sort_grdColumns(grdPOTORDR2, "INIT_DATE".ToLower,, 1)
                Fill_Records("ICTIRECX", PO_ORDER_NO)
                EnforceConstraints(False)
            End If
            Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")
            'grdPOTORDR2.Rows.ExpandAll(True)

            grdPOTORDR2.Text = "PO Details for PO " & PO_ORDER_NO & " - All Lines, whether Open or Closed"
            grdPOTORDR2.Visible = True
        End If
    End Sub

    Private Sub tabPOTORDRV_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPOTORDRV.SelectedTabChanged
        'If tabPOTVENDZ.SelectedTab.Key = "Shipments" Then

        'End If
    End Sub


    Sub Refresh_Summary()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating Work Tables")

        Create_WorkTable(False)

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTORDRV", "POTORDR1", "POTORDR2", "POTORDR8", "ICTIRECX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress("Now Creating Vendor Summary")

        If EntryMode = "E" Then
            dst.Tables("POTORDR8").Rows.Clear()
            ASCMAIN1.sql = $"Select POTORDR2.* from POTORDR2, {POTORDRX} POTORDRX, POTORDR1
where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO and POTORDRX.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
  and POTORDRX.VEND_CODE = '{VEND_CODE}'"
            Fill_Records("POTORDR2",,, ASCMAIN1.sql)
        End If

        If EntryMode = "E" Then
            grdPOTORDRV.Text = "All Open POs"
        ElseIf EntryMode = "S" Then
            grdPOTORDRV.Text = "All Open POs - DC Only"
        Else
            If optPOs.Value = "Y" Then
                grdPOTORDRV.Text = "POs Placed YTD, by Vendor"
            ElseIf optPOs.Value = "L" Then
                grdPOTORDRV.Text = "POs Placed Last Year, by Vendor"
            Else
                grdPOTORDRV.Text = "All Open POs, by Vendor"
            End If
        End If
        Fill_Records("POTORDRV")
        Sort_grdColumns(grdPOTORDRV, "VEND_CODE")

        If EntryMode = "E" Then
            grdPOTORDRD.Text = "All Open POs"
        ElseIf EntryMode = "S" Then
            grdPOTORDRD.Text = "All Open POs - DC Only"
        Else
            If optPOs.Value = "Y" Then
                grdPOTORDRD.Text = "POs Placed YTD, Details"
            ElseIf optPOs.Value = "L" Then
                grdPOTORDRD.Text = "POs Placed Last Year, Details"
            Else
                grdPOTORDRD.Text = "All Open POs, Details"
            End If
        End If
        Fill_Records("POTORDRD")
        Sort_grdColumns(grdPOTORDRD, "VEND_CODE")


        If EntryMode = "E" Then
            grdPOTORDRA.Text = "All Open POs"
        ElseIf EntryMode = "S" Then
            grdPOTORDRA.Text = "All Open POs - DC Only"
        Else
            If optPOs.Value = "Y" Then
                grdPOTORDRA.Text = "POs Placed YTD"
            ElseIf optPOs.Value = "L" Then
                grdPOTORDRA.Text = "POs Placed Last Year"
            Else
                grdPOTORDRA.Text = "All Open POs"
            End If
        End If
        Fill_Records("POTORDRA")
        Sort_grdColumns(grdPOTORDRA, "VEND_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTORDRV_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTORDRV.InitializeLayout

    End Sub

    Private Sub grdPOTORDRV_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTORDRV.DoubleClickRow

        Absx1.txtFor("VEND_CODE").Text = e.Row.Cells("VEND_CODE").Value & ""

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Vendor Data")

        Click_Command("Load")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optPOs_ValueChanged(sender As Object, e As EventArgs) Handles optPOs.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Summary()
    End Sub

    Private Sub chkExclLabs_CheckedChanged(sender As Object, e As EventArgs)
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Summary()
    End Sub

    Sub Status_Request()

        Dim VEND_CODEs As New List(Of String)
        For Each rowPOTORDRV As DataRow In dst.Tables("POTORDRV").Select("SEL='1'", "VEND_CODE")
            Dim VEND_CODE As String = rowPOTORDRV.Item("VEND_CODE")
            VEND_CODEs.Add(VEND_CODE)
        Next

        Refresh_Summary()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating PO Status Request Reports")

        ems = New List(Of ASCNOTE1)

        For Each VEND_CODE As String In VEND_CODEs
            email_PO_Status_Request_XLS(VEND_CODE)
        Next

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now emailing PO Status Request Reports")

        For Each em As ASCNOTE1 In ems
            em.EmailDocument()
            em = Nothing
        Next
        ems.Clear()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox("Status Request Reports have been emailed to you" & vbCrLf & " to forward to each Vendor", MsgBoxStyle.OkOnly, "Success")
    End Sub

    Sub email_PO_Status_Request_XLS(VEND_CODE As String)

        Dim PO_ORDER_NO As String = ""
        Dim rowPOTORDR1 As DataRow = Nothing

        Dim EMAIL_SUBJECT As String = "Request for PO Status"

        Dim rowPOTORDRV As DataRow = dst.Tables("POTORDRV").Rows.Find(VEND_CODE)

        ' XLS Worksheet

        Me.Cursor = Cursors.WaitCursor

        Dim FILENAME_CTL_NO As String = ASCMAIN1.Next_Control_No("POTORDRV.XLS_NO")
        FILENAME_CTL_NO = Mid(FILENAME_CTL_NO, Len(FILENAME_CTL_NO) - 5, 6)
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & VEND_CODE & "_" & FILENAME_CTL_NO & ".xlsx"

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim pageSetup As SpreadsheetGear.IPageSetup = worksheet.PageSetup

        Dim ws As Integer = -1

        Dim sqlPOs As String = $"VEND_CODE = '{VEND_CODE}'"
        Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("POTORDRD").Select(sqlPOs), "CUST_ACCT_CODE")

        For Each rowVA As DataRow In tbl.Select("")

            ws += 1
            If ws > 0 Then
                worksheet = workbook.Worksheets.Add()
            End If

            Dim CUST_ACCT_CODE As String = rowVA.Item("CUST_ACCT_CODE") & ""
            If CUST_ACCT_CODE = "" Then CUST_ACCT_CODE = CStr(workbook.Worksheets.Count + 1)

            With worksheet

                Dim sheetName As String = CUST_ACCT_CODE
                For Each chr As Char In "[]*/?\:'"
                    sheetName = sheetName.Replace(chr, "_")
                Next

                If sheetName.Length > 30 Then
                    sheetName = sheetName.Substring(0, 30).Trim
                End If

                .Name = sheetName
                Dim r As Integer = -1

                r += 1 : .Cells(r, 0).Value = "Vendor"
                r += 0 : .Cells(r, 1).Value = VEND_CODE
                r += 0 : .Cells(r, 2).Value = CUST_ACCT_CODE
                r += 0 : .Cells(r, 3).Value = rowPOTORDRV.Item("VEND_NAME") & ""

                .Range(0, 0, r, 0).Interior.Color = SpreadsheetGear.Colors.LightGray

                pageSetup.PrintTitleRows = "A1: " & .Cells(r, 2).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing)

                r += 2

                ' create a new sheet for each account

                Dim c As Integer = -1

                c += 1 : .Cells(r, c).Value = "PO Date" : .Cells(r, c).EntireColumn.NumberFormat = "MM/dd/yyyy"
                'c += 1 : .Cells(r, c).Value = "PO Ref (Prev PO)" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Whse" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Notes" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Type" : .Cells(r, c).EntireColumn.NumberFormat = "@"

                c += 1 : .Cells(r, c).Value = "PO No" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Line" : .Cells(r, c).EntireColumn.NumberFormat = "##0" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                c += 1 : .Cells(r, c).Value = "UPC Code" : .Cells(r, c).EntireColumn.NumberFormat = "@"

                c += 1 : .Cells(r, c).Value = "Description" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "ABB Item Code" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Qty Ordered" : .Cells(r, c).EntireColumn.NumberFormat = "#,###" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Qty Received" : .Cells(r, c).EntireColumn.NumberFormat = "#,###" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Qty Open" : .Cells(r, c).EntireColumn.NumberFormat = "#,###" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Unit Cost" : .Cells(r, c).EntireColumn.NumberFormat = "#,##0.00" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Amount" : .Cells(r, c).EntireColumn.NumberFormat = "#,##0.00" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Confirmation No" : .Cells(r, c).EntireColumn.NumberFormat = "@" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

                c += 1
                c += 1 : .Cells(r, c).Value = "Exp Ship Date" : .Cells(r, c).EntireColumn.NumberFormat = "mm/dd/yy" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                c += 1 : .Cells(r, c).Value = "Notes" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Qty Shipped" : .Cells(r, c).EntireColumn.NumberFormat = "#,###" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                c += 1 : .Cells(r, c).Value = "Date Shipped" : .Cells(r, c).EntireColumn.NumberFormat = "mm/dd/yy" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                c += 1 : .Cells(r, c).Value = "Tracking Number" : .Cells(r, c).EntireColumn.NumberFormat = "@"
                c += 1 : .Cells(r, c).Value = "Whse ETA Date" : .Cells(r, c).EntireColumn.NumberFormat = "mm/dd/yy" : .Cells(r, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

                .Range(r, 0, r, c - 7).Interior.Color = SpreadsheetGear.Colors.LightGray
                .Range(r, c - 5, r, c - 0).Interior.Color = SpreadsheetGear.Colors.Yellow

                Dim r0 As Integer = r + 1
                ' Dim cQty As Integer = 0

                For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD") _
                    .Select($"VEND_CODE = '{VEND_CODE}' and CUST_ACCT_CODE = '{CUST_ACCT_CODE}' AND PO_QTY_OPN <> 0", "PO_ORDER_NO, PO_ORDER_LNO")

                    If PO_ORDER_NO <> rowPOTORDRD.Item("PO_ORDER_NO") Then
                        PO_ORDER_NO = rowPOTORDRD.Item("PO_ORDER_NO")
                        rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
                    End If

                    'Dim VEND_ACCT_NO As String = rowPOTORDR1.Item("CUST_ACCT_CODE")
                    'Dim WHSE_CODE As String = rowPOTORDR1.Item("WHSE_CODE")

                    Dim PO_ORDER_TYPE As String = rowPOTORDRD.Item("PO_ORDER_TYPE") & ""
                    Dim PO_ORDER_TYPE_DESC As String = ""
                    Dim rowPOTORDR1_PO_ORDER_TYPE As DataRow = dst.Tables("POTORDR1_PO_ORDER_TYPE").Rows.Find(PO_ORDER_TYPE)
                    If rowPOTORDR1_PO_ORDER_TYPE IsNot Nothing Then
                        PO_ORDER_TYPE_DESC = rowPOTORDR1_PO_ORDER_TYPE.Item("T_DESC") & ""
                    End If


                    r += 1 : c = -1
                    c += 1 : .Cells(r, c).Value = $"'{Format(rowPOTORDRD.Item("PO_DATE_ORDERED"), "MM/dd/yyyy")}"
                    'c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("PO_REFERENCE")}"
     
                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("PO_NOTES")}"
                    c += 1 : .Cells(r, c).Value = $"'{PO_ORDER_TYPE_DESC}"

                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("PO_ORDER_NO")}"
                    c += 1 : .Cells(r, c).Value = Val(rowPOTORDRD.Item("PO_ORDER_LNO"))
                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("VEND_ITEM_NO") & ""}"
                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("ITEM_DESC") & ""}"
                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDRD.Item("ITEM_CODE") & ""}"
                    c += 1 : .Cells(r, c).Value = Val(rowPOTORDRD.Item("PO_QTY_ORD") & "")
                    c += 1 : .Cells(r, c).Value = Val(rowPOTORDRD.Item("PO_QTY_REC") & "")
                    c += 1 : .Cells(r, c).Value = Val(rowPOTORDRD.Item("PO_QTY_OPN") & "") ' : cQty = c
                    c += 1 : .Cells(r, c).Value = Val(rowPOTORDRD.Item("PO_COST") & "")
                    c += 1 : .Cells(r, c).Formula = "=" & .Cells(r, c - 2).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing) & " * " & .Cells(r, c - 1).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing)
                    c += 1 : .Cells(r, c).Value = $"'{rowPOTORDR1.Item("VEND_CONF_NO") & ""}"
                    c += 1

                    Dim ETD As String = ""
                    If rowPOTORDRD.Item("PO_DATE_ETD") & "" <> "" Then ETD = Format(rowPOTORDRD.Item("PO_DATE_ETD"), "MM/dd/yy")

                    Dim PO_DATE_ETD_NOTES As String = rowPOTORDRD.Item("PO_DATE_ETD_NOTES") & ""
                    Dim PO_DATE_ETD As String = rowPOTORDRD.Item("PO_DATE_ETD") & ""

                    c += 1 : .Cells(r, c).Value = ETD
                    c += 1 : .Cells(r, c).Value = PO_DATE_ETD_NOTES
                    c += 1 ' Qty Shipped
                    c += 1 ' Date Shipped
                    c += 1 ' Tracking Number
                    c += 1 : .Cells(r, c).Value = PO_DATE_ETD ' Whse ETA Date
                Next

                '.Cells(r + 2, cQty).Formula = "=SUM(" & .Cells(r0, cQty).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing) _
                '                                            & ":" & .Cells(r, cQty).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing) & ")"

                '.Cells(r + 2, c).Formula = "=SUM(" & .Cells(r0, c).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing) _
                '                                        & ":" & .Cells(r, c).GetAddress(False, False, SpreadsheetGear.ReferenceStyle.A1, False, Nothing) & ")"
                '.Cells(r + 2, 0).Value = "Total"

                .Cells.Columns.AutoFit()
                For CX As Integer = 0 To c - 3
                    .Cells(0, CX).EntireColumn.ColumnWidth *= 1.25
                Next

                .Cells(0, c - 6).EntireColumn.ColumnWidth = 3
                .Cells(0, c - 5).EntireColumn.ColumnWidth = 15
                .Cells(0, c - 4).EntireColumn.ColumnWidth = 30
                .Cells(0, c - 3).EntireColumn.ColumnWidth = 15
                .Cells(0, c - 2).EntireColumn.ColumnWidth = 15
                .Cells(0, c - 1).EntireColumn.ColumnWidth = 30
                .Cells(0, c - 0).EntireColumn.ColumnWidth = 15
            End With

        Next

        workbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Dim em As New ASCNOTE1("POFTRACK", dst)
        em.CreateComponents()
        em.SetEmailSubject(EMAIL_SUBJECT)

        Dim strHtml As String = "" '  Build_Html_Email(emailParms)

        Dim EMAILS As String = ""

        'For Each rowTATCONT1 As DataRow In ASCDATA1.GetDataTable.Select("", "CONTACT_NO")
        '    Dim CONTACT_EMAIL As String = rowTATCONT1.Item("CONTACT_EMAIL")
        '    EMAILS &= ";" & rowTATCONT1.Item("CONTACT_EMAIL")
        'Next

        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
        Dim VEND_PURCH_EMAIL As String = rowAPTVEND1.Item("VEND_PURCH_EMAIL") & ""
        Dim VEND_ACCT_NO As String = rowAPTVEND1.Item("VEND_ACCT_NO") & ""

        If VEND_PURCH_EMAIL <> "" Then EMAILS &= ";" & VEND_PURCH_EMAIL

        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)


        Dim yellow As String = " style='background-color:yellow'"
        Dim blue As String = " style='color:blue'"
        Dim margin As String = " style='margin-left:20px'"


        strHtml = "<h1>Request for Status</h1>"


        strHtml &= $"<br/><div>Vendor Contacts who should receive this email<br/>{Replace(Mid(EMAILS, 2), ";", "<br/>")}</div><br/>"

        If Now.Hour < 12 Then
            strHtml &= $"<br/><div>Good Morning</div>"
        Else
            strHtml &= $"<br/><div>Good Afternoon</div>"
        End If

        strHtml &= $"<br/><div>Instructions:</div>"
        strHtml &= $"<ul>"
        strHtml &= $"<li>Please fill out Ship Date And Notes (if any) in the Yellow Columns of the Attachex XLS file.</strong>.</li>"
        strHtml &= $"<li>Please make sure your Dates are formatted as dates (MM/DD/YYYY) and do not put messages into Date Columns.</li>"
        If tbl.Rows.Count > 1 Then strHtml &= $"<li>We have multiple Accounts with you with Open Purchase Orders. POs for each Account are listed on separate XLS Sheets</li>"
        strHtml &= $"</ul>"


        strHtml &= $"<br/><div>Thank You,</div><br/>" & TAC.POCMAIN1.GetUserSignature(rowASTUSER1)

        em.SetDocumentBody(strHtml)
        em.SetEmailTo(ASCMAIN1.USER_EMAIL) '  & ";wjz@absolution.com")
        em.Attachments.Add(FILENAME)

        ems.Add(em)

        worksheet = Nothing
        workbook = Nothing
    End Sub

    Sub Import_Status_Response(FILENAME As String)
        Try
            Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim PO_ORDER_NO_c As Int32 = 2
            Dim PO_ORDER_LNO_c As Int32 = PO_ORDER_NO_c + 1
            Dim PO_DATE_ETD_c As Int32 = PO_ORDER_NO_c + 10
            Dim PO_DATE_ETD_NOTES_c As Int32 = PO_DATE_ETD_c + 1

            Dim PO_PARM_PINV_LT As Integer = Val(ROWs("POTPARM1").Item("PO_PARM_PINV_LT") & "")

            Dim vendor_was_verified As Boolean = False

            For Each worksheet In workbook.Worksheets
                Dim r As Integer = 2
                With worksheet
                    Do Until .Cells(r, PO_ORDER_NO_c).Value & "" = ""
                        Dim PO_ORDER_NO As String = .Cells(r, PO_ORDER_NO_c).Value & "" ' F PO No
                        Dim PO_ORDER_LNO As Integer = Val(.Cells(r, PO_ORDER_LNO_c).Value & "") ' G PO Line
                        Dim PO_DATE_ETD As String = .Cells(r, PO_DATE_ETD_c).Text ' Q Ship Date
                        Dim PO_DATE_ETD_NOTES As String = .Cells(r, PO_DATE_ETD_NOTES_c).Value & "" ' R Notes

                        If Not vendor_was_verified Then
                            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                            If rowPOTORDR1 Is Nothing Then
                                MsgBox($"Could not find Lead PO {PO_ORDER_NO}", MsgBoxStyle.OkOnly, $"Cannot Upload Status for Vendor {VEND_CODE} using this XLS file")
                                Exit Sub
                            Else
                                If rowPOTORDR1.Item("VEND_CODE") & "" <> VEND_CODE Then
                                    MsgBox($"Lead PO {PO_ORDER_NO} is on file with Vendor {rowPOTORDR1.Item("VEND_CODE")}", MsgBoxStyle.OkOnly, $"Cannot Upload Status for Vendor {VEND_CODE} using this XLS file")
                                    Exit Sub
                                End If
                            End If
                            vendor_was_verified = True
                        End If

                        Dim PO_DATE_REQUIRED As String = ""
                        If PO_DATE_ETD <> "" Then
                            PO_DATE_REQUIRED = Format(CDate(PO_DATE_ETD).AddDays(PO_PARM_PINV_LT), "MM/dd/yyyy") '.Cells(r, PO_ORDER_NO_c + 15).Text ' T Date Shipped
                        End If

                        If PO_DATE_ETD <> "" Or PO_DATE_ETD_NOTES <> "" Then
                            If PO_DATE_ETD_NOTES.Length > 60 Then PO_DATE_ETD_NOTES = Mid(PO_DATE_ETD_NOTES, 1, 60)
                            If Not IsDate(PO_DATE_ETD) Then PO_DATE_ETD = ""
                            If Not IsDate(PO_DATE_REQUIRED) Then PO_DATE_REQUIRED = String.Empty

                            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})

                            If rowPOTORDR2 IsNot Nothing Then
                                If PO_DATE_ETD <> "" Then rowPOTORDR2.Item("PO_DATE_ETD") = CDate(PO_DATE_ETD)
                                If PO_DATE_ETD <> "" Or PO_DATE_ETD_NOTES <> "" Then rowPOTORDR2.Item("PO_DATE_ETD_NOTES") = PO_DATE_ETD_NOTES
                                'If PO_DATE_ETD <> "" Then rowPOTORDR2.Item("PO_DATE_ETD_NOTES") = PO_DATE_ETD_NOTES
                                If PO_DATE_REQUIRED <> String.Empty Then rowPOTORDR2.Item("PO_DATE_REQUIRED") = CDate(PO_DATE_REQUIRED)
                            End If

                            Dim rowPOTORDRD As DataRow = dst.Tables("POTORDRD").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
                            If rowPOTORDRD IsNot Nothing Then
                                If PO_DATE_ETD <> "" Then rowPOTORDRD.Item("PO_DATE_ETD") = CDate(PO_DATE_ETD)
                                If PO_DATE_ETD_NOTES <> "" Then rowPOTORDRD.Item("PO_DATE_ETD_NOTES") = PO_DATE_ETD_NOTES
                                If PO_DATE_REQUIRED <> String.Empty Then rowPOTORDR2.Item("PO_DATE_REQUIRED") = CDate(PO_DATE_REQUIRED)
                            End If
                        End If

                        r += 1
                    Loop
                End With
            Next

            MsgBox("Upload Successful" & vbCrLf & vbCrLf & "***** IMPORTANT *****" & vbCrLf & vbCrLf & "Remember to click Update", MsgBoxStyle.OkOnly, "Success")

        Catch ex As Exception
            MsgBox("Error processing XLSX file:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Cannot Process PO Status Reponse file Selected")
        End Try

    End Sub


    Function GetBase64Logo() As String
        Return "data:image/png;base64,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"
    End Function

    Private Sub grdPOTORDRD_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTORDRD.DoubleClickRow
        If Not ScreenMode AndAlso e.Row.IsDataRow Then
            Dim VEND_CODE As String = e.Row.Cells("VEND_CODE").Value
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE
            Click_Command("Load")
        End If
    End Sub

    Private Sub grdPOTORDRA_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTORDRA.DoubleClickRow
        If Not ScreenMode AndAlso e.Row.IsDataRow Then
            Dim VEND_CODE As String = e.Row.Cells("VEND_CODE").Value
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE
            Click_Command("Load")
        End If
    End Sub

    Private Sub grdPOTORDRD_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTORDRD.AfterRowUpdate
        Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
        Dim PO_ORDER_LNO As Int32 = Val(e.Row.Cells("PO_ORDER_LNO").Value & "")
        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
        For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
            rowPOTORDR2.Item(C) = e.Row.Cells(C).Value
        Next
        Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(New Object() {PO_ORDER_NO})
        For Each C As String In New String() {"PO_NOTES", "VEND_CONF_NO", "VEND_CONF_DATE"}
            rowPOTORDR1.Item(C) = e.Row.Cells(C).Value
        Next
    End Sub

    Private Sub grdPOTORDR2_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTORDR2.AfterRowUpdate
        Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
        Dim PO_ORDER_LNO As Int32 = Val(e.Row.Cells("PO_ORDER_LNO").Value & "")
        Dim rowPOTORDRD As DataRow = dst.Tables("POTORDRD").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
        For Each C As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
            rowPOTORDRD.Item(C) = e.Row.Cells(C).Value
        Next
    End Sub

    Private Sub grdPOTORDR1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTORDR1.AfterRowUpdate
        Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
        Dim rowPOTORDRDs() As DataRow = dst.Tables("POTORDRD").Select($"PO_ORDER_NO = '{PO_ORDER_NO}'")
        For Each rowPOTORDRD As DataRow In rowPOTORDRDs
            For Each C As String In New String() {"PO_NOTES", "VEND_CONF_NO", "VEND_CONF_DATE"}
                rowPOTORDRD.Item(C) = e.Row.Cells(C).Value
            Next
        Next
    End Sub

    Private Sub grdPOTORDR1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTORDR1.InitializeRow
        If EntryMode <> "E" Then

        End If

    End Sub


    Private Sub grdPOTORDRD_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTORDRD.InitializeRow
        grdPOTORDRD.DisplayLayout.Bands(0).Columns("PO_ORDER_NO").Header.ToolTipText = "Red means ETA (or FBO Date) is in the past, Yellow means it is within 10 days"

        If EntryMode <> "E" Then

        End If
    End Sub

    Private Sub tabSummary_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSummary.SelectedTabChanged

        If Me.SELECTION_NO = 0 Then Exit Sub

        UltraExplorerBar1.Groups("PO Receipts").Visible = False ' (tabSummary.SelectedTab.Key = "PO Receipts")
        UltraExplorerBar1.Groups("Invoiced Receipts").Visible = False '  (tabSummary.SelectedTab.Key = "Invoiced Pricing Discrepancies")
        UltraExplorerBar1.Groups("PO Summary").Visible = Not ((tabSummary.SelectedTab.Key = "PO Receipts") Or (tabSummary.SelectedTab.Key = "Invoiced Pricing Discrepancies"))

        If (tabSummary.SelectedTab.Key = "PO Receipts") Then
            Fetch_Receipts_Data()
        End If
        If (tabSummary.SelectedTab.Key = "Invoiced Pricing Discrepancies") Then
            Fetch_Invoiced_Data()
        End If

    End Sub

    Sub Fetch_Receipts_Data()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering PO Receipts Data")

        Setup_Receipts()

        'CreateGraph()
        chtICTIRECD.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_Receipts()

        Dim YP As String = cbeYP.Value

        ' CreateGraph()
    End Sub

    Sub Transfer_Rows(TABLE_NAME_from As String, TABLE_NAME_to As String)

        dst.Tables(TABLE_NAME_to).Rows.Clear()

        For Each row As DataRow In dst.Tables(TABLE_NAME_from).Select("")
            Dim row2 As DataRow = dst.Tables(TABLE_NAME_to).NewRow
            For I As Integer = 0 To dst.Tables(TABLE_NAME_from).Columns.Count - 1
                row2.Item(I) = row.Item(I)
            Next
            dst.Tables(TABLE_NAME_to).Rows.Add(row2)
        Next
    End Sub

    Sub Fetch_Invoiced_Data()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Invoiced Data")

        Dim YP_Invoiced As String = cbeYP_Invoiced.Value
        Fill_Records("APTINVHD", YP_Invoiced)
        Sort_grdColumns(grdAPTINVHD, "ITEM_CODE")

        Fill_Records("APTINVHS", YP_Invoiced)
        Sort_grdColumns(grdAPTINVHS, "VEND_CODE,PO_ORDER_NO,INV_NUM")
        grdAPTINVHS.Text = $"Invoiced Receipts Summary for {cbeYP_Invoiced.Text}"

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Fetch_Receipts_Data()
    End Sub


    Private Sub cbeYP_Invoiced_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP_Invoiced.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Fetch_Invoiced_Data()
    End Sub

    Sub CreateGraph()

        Dim chtIsVisible As Boolean = chtICTIRECD.Visible
        chtICTIRECD.Visible = False

        chtICTIRECD.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String

        ReDim CL(1)
        If optRecType.Value = "P" Then
            CL(1) = "NPI"
            CL(0) = "STK"
        Else
            CL(1) = "Returns"
            CL(0) = "Shelf"
        End If


        chtICTIRECD.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtICTIRECD.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtICTIRECD.LabelHash = labelHash

        chtICTIRECD.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtICTIRECD.Tooltips.FormatString = "<HIGHLOW>"

        Dim RLi As Integer = 0

        ReDim RL(dst.Tables("ICTIRECD").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("ICTIRECD").Select("", "RECEIPT_DATE")
            RL(RLi) = row(0)
            RLi += 1
        Next


        chtICTIRECD.Data.SetRowLabels(RL)
        chtICTIRECD.Data.SetColumnLabels(CL)

        chtICTIRECD.ChartType = ChartType.StackColumnChart
        Dim DVW As DataView = dst.Tables("ICTIRECD").DefaultView
        DVW.Sort = "RECEIPT_DATE"
        chtICTIRECD.DataSource = DVW.ToTable
        'chtICTIRECD.Data.IncludeColumn("REC_TYPE", False)
        'chtICTIRECD.ChartType = ChartType.ColumnChart

        chtICTIRECD.TitleTop.Text = $"Receipts History"

        chtICTIRECD.Legend.Visible = True
        chtICTIRECD.Legend.Location = LegendLocation.Right
        chtICTIRECD.Legend.Margins.Left = 5
        chtICTIRECD.Legend.Margins.Right = 10
        chtICTIRECD.Legend.Margins.Top = 15
        chtICTIRECD.Legend.Margins.Bottom = 90
        chtICTIRECD.Legend.SpanPercentage = 10

        chtICTIRECD.DataBind()

        chtICTIRECD.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub


    Public Class MyCustomTooltip
        Implements IRenderLabel

        Public Sub New()

        End Sub 'New

        Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
            Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

        End Function 'ToString 
    End Class 'MyCustomTooltip


    Private Sub grdAPTINVHS_AfterRowActivate(sender As Object, e As EventArgs) Handles grdAPTINVHS.AfterRowActivate
        Setup_APTINVHD()


    End Sub

    Sub Setup_APTINVHD()


        For Each C As String In New String() {"VEND_CODE", "PO_ORDER_NO", "INV_NUM"}
            grdAPTINVHD.DisplayLayout.Bands(0).Columns(C).Hidden = Not chkShowAll.Checked
        Next

        If chkShowAll.Checked Then
            Dim dvw As DataView = DirectCast(grdAPTINVHD.DataSource, DataTable).DefaultView
            dvw.RowFilter = ""
            grdAPTINVHD.Text = $"Invoiced Receipts Details for {cbeYP_Invoiced.Text}"
            grdAPTINVHD.Visible = True

        Else
            With grdAPTINVHS.ActiveRow
                If grdAPTINVHS.ActiveRow Is Nothing OrElse .IsFilterRow OrElse Not .IsDataRow Then
                    grdAPTINVHD.Visible = False
                Else
                    Dim VEND_CODE As String = .Cells("VEND_CODE").Value
                    Dim PO_ORDER_NO As String = .Cells("PO_ORDER_NO").Value
                    Dim INV_NUM As String = .Cells("INV_NUM").Value
                    Dim sqlw As String = $"VEND_CODE = '{VEND_CODE}'  and PO_ORDER_NO = '{PO_ORDER_NO}' and INV_NUM = '{INV_NUM}'"
                    Dim dvw As DataView = DirectCast(grdAPTINVHD.DataSource, DataTable).DefaultView
                    dvw.RowFilter = sqlw
                    grdAPTINVHD.Text = $"Invoiced Receipts Details for {VEND_CODE} {PO_ORDER_NO} {INV_NUM}"
                    grdAPTINVHD.Visible = True
                End If
            End With
        End If
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        Setup_APTINVHD()
    End Sub

    Private Sub optRecType_ValueChanged(sender As Object, e As EventArgs) Handles optRecType.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_Receipts()
    End Sub

End Class