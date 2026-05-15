Imports System.Math

Public Class ICRIADJ1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim ICTIADJ1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 2

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        SUBT = ""

        Dim sqlw As String = IIf(MENU_ITEM_OBJECT = "ICRIADJ1", "ICTIADJ1.JOURNAL_IND = '0'", "ICTIADJ1.REGISTER_IND = '0'")

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Adjustments Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Adjustments Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "ICTIADJ1.ADJ_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Adjustments Posted in " & xRYP0_legend
            Else
                SUBT = "Adjustments Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "ICTIADJ1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        If MENU_ITEM_OBJECT = "ICRIADJ1" And ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If

        Prepare_dst(True, sqlw)

        Check_if_Empty("ICTIADJ1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = IIf(MENU_ITEM_FORM = "", MENU_ITEM_OBJECT, MENU_ITEM_FORM)
        Generate_Report(RPT, , SUBT)
        If MENU_ITEM_OBJECT = "ICRIADJ1" Then
            Print_GL()
        End If

        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("ICTIADJ2")
        grdASTEXPT1.Text = "Inventory Adjustments Details - " & Absx1.optFor("RANGE").Text & " : " & SUBT
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "ADJ_NO", "Adj No", 100)
        Set_DX_Column(grdASTEXPT1, "ADJ_LNO", "Adj Lno", 50, "#,##0")
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 60, "", "", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "Cost Catgy", 70, "", "", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70, "", "", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_CODE", "Whse", 60, "", "", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "REASON_CODE", "Reason", 60, "", "", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item", 100)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)

        Set_DX_Column(grdASTEXPT1, "ADJ_QTY", "Adj Qty", 90, "#,##0", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_STD", "Std Cost", 90, "#,##0.0000", "", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "EXT_STD", "Ext Std", 90, "#,##0.00", "Sum", Color.Gold)

        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "YP", 60)
        Set_DX_Column(grdASTEXPT1, "ADJ_REF", "Ref", 60)
        Set_DX_Column(grdASTEXPT1, "ADJ_DATE", "Adj Date", 90)

        Set_DX_Column(grdASTEXPT1, "ADJ_NOTE", "Adj Note", 130)
        Set_DX_Column(grdASTEXPT1, "INIT_OPER", "Entered By", 100)
        Set_DX_Column(grdASTEXPT1, "INIT_DATE", "Entered", 90)
        Set_DX_Column(grdASTEXPT1, "RTRN_NO", "Return No", 130)
        Set_DX_Column(grdASTEXPT1, "REVERSED_BY_ADJ_NO", "Reversed By", 70)
        Set_DX_Column(grdASTEXPT1, "REVERSES_ADJ_NO", "Reverses", 70)

        '  grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "ADJ_NO,ADJ_LNO")

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String = "Update ICTIADJ1 " _
            & IIf(MENU_ITEM_OBJECT = "ICRIADJ1", _
                  " Set JOURNAL_IND = :PARM1, JOURNAL_XNO = :PARM2", _
                  " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2") _
            & " where ADJ_NO in (Select ADJ_NO from " & ICTIADJ1 & " )"
        ASCDATA1.ExecuteSQL(sql, "VV", New Object() {"1", MyBase.XNO})

        If MENU_ITEM_OBJECT = "ICRIADJ1" Then
            GL_Update()
        End If

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"
        ASCMAIN1.sql = "Select * from ICTIADJ1 where " & sqlw
        ICTIADJ1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTIADJ1 & " Add Primary Key (ADJ_NO)")

        ASCMAIN1.sql = "Select ICTIADJ1.*,ICTREAS1.REASON_DESC " & vbCrLf _
            & " from " & ICTIADJ1 & " ICTIADJ1,ICTREAS1 " & vbCrLf _
            & " where ICTREAS1.REASON_CODE (+) = ICTIADJ1.REASON_CODE " & vbCrLf _
            & "   and " & sqlw
        Create_TDA(dst.Tables.Add, "ICTIADJ1", "**", 0)

        ASCMAIN1.sql = "Select ICTIADJ2.*, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE " & vbCrLf _
            & ", ICTIADJ1.ADJ_DATE, ICTIADJ1.WHSE_CODE, ICTIADJ1.REASON_CODE, ICTIADJ1.ADJ_NOTE, ICTIADJ1.INIT_OPER, ICTIADJ1.INIT_DATE, ICTIADJ1.RTRN_NO, ICTIADJ1.REVERSED_BY_ADJ_NO, ICTIADJ1.REVERSES_ADJ_NO" & vbCrLf _
            & " from ICTIADJ2," & ICTIADJ1 & " ICTIADJ1, ICTITEM1, ICTCOLL1 " & vbCrLf _
            & " where ICTIADJ2.ADJ_NO = ICTIADJ1.ADJ_NO" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = ICTIADJ2.ITEM_CODE"
        Create_TDA(dst.Tables.Add, "ICTIADJ2", "**", 0)
        dst.Tables("ICTIADJ2").Columns.Add("EXT_STD", GetType(System.Decimal), "ADJ_QTY * ITEM_COST_STD")


        ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC " & vbCrLf _
            & " from ICTIADJ3," & ICTIADJ1 & " ICTIADJ1, GLTACCT1 " & vbCrLf _
            & " where ICTIADJ3.ADJ_NO = ICTIADJ1.ADJ_NO" & vbCrLf _
            & " and GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE"
        Create_TDA(dst.Tables.Add, "ICTIADJ3", "**", 0)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("ICTIADJ1")
        Fill_Records("ICTIADJ2")
        Fill_Records("ICTIADJ3")
        If RWU = "R" Then
            TAC.ICCMAIN1.Prepare_GL_Interface("ICIA", ICTIADJ1)
        End If
        EnforceConstraints(True)
    End Sub
End Class