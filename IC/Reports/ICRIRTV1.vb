Imports System.Math

Public Class ICRIRTV1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim ICTIRTV1 As String

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

        Dim sqlw As String = "ICTIRTV1.REGISTER_IND = '0' and nvl(ICTIRTV1.RTV_STATUS,0) <> 'H'"

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "RTVs Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "RTVs Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "ICTIRTV1.RTV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "RTVs Posted in " & xRYP0_legend
            Else
                SUBT = "RTVs Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "ICTIRTV1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        If ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If

        sqlw &= SQL_in("RTV_NO", "ICTIRTV1.RTV_NO")
        sqlw &= SQL_in("VEND_CODE_S", "ICTIRTV1.VEND_CODE")

        Prepare_dst(True, sqlw)

        Check_if_Empty("ICTIRTV1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SUMMARY", "0")
        Generate_Report(RPT, , "Detail")
        'CR_params.Add("SUMMARY", "1")
        'Generate_Report(RPT, , "Summary")


        If Absx1.optFor("RANGE").Value = "N" Then
            Call Print_GL()
        End If

        'Generate_Report("ICRIRTVV", "Return to Vendor", "Vendor Copy")

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

        Dim sql As String = "Update ICTIRTV1 " _
        & " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2" _
        & " where RTV_NO in (Select RTV_NO from " & ICTIRTV1 & " )"
        ASCDATA1.ExecuteSQL(sql, "VV", New Object() {"1", MyBase.XNO})

        GL_Update()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        ASCMAIN1.sql = "Select * from ICTIRTV1 where " & sqlw

        ICTIRTV1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTIRTV1 & " Add Primary Key (RTV_NO)")

        ASCMAIN1.sql = "Select ICTIRTV1.*, APTVEND1.VEND_NAME" _
        & " from " & ICTIRTV1 & " ICTIRTV1,APTVEND1 " _
        & "   where APTVEND1.VEND_CODE (+) = ICTIRTV1.VEND_CODE " _
        & "   and " & sqlw
        Call Create_TDA(dst.Tables.Add, "ICTIRTV1", "**", 0)

        ASCMAIN1.sql = "Select ICTIRTV2.*, ICTITEM1.ITEM_DESC " _
        & " from ICTIRTV2," & ICTIRTV1 & " ICTIRTV1, ICTITEM1 " _
        & " where ICTIRTV2.RTV_NO = ICTIRTV1.RTV_NO" _
        & " and ICTITEM1.ITEM_CODE = ICTIRTV2.ITEM_CODE"
        Call Create_TDA(dst.Tables.Add, "ICTIRTV2", "**", 0)

        ASCMAIN1.sql = "Select ICTIRTV3.*, GLTACCT1.ACCT_DESC " _
        & " from ICTIRTV3," & ICTIRTV1 & " ICTIRTV1, GLTACCT1 " _
        & " where ICTIRTV3.RTV_NO = ICTIRTV1.RTV_NO" _
        & " and GLTACCT1.ACCT_CODE = ICTIRTV3.ACCT_CODE"
        Call Create_TDA(dst.Tables.Add, "ICTIRTV3", "**", 0)

        ASCMAIN1.sql = "Select * " _
        & " from APTVEND1 where VEND_CODE IN ( Select Distinct VEND_CODE from " & ICTIRTV1 & " )" 
        Call Create_TDA(dst.Tables.Add, "APTVEND1", "**", 0)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("ICTIRTV1")
        Fill_Records("ICTIRTV2")
        Fill_Records("ICTIRTV3")
        Fill_Records("APTVEND1")
        TAC.ICCMAIN1.Prepare_GL_Interface("ICIV", ICTIRTV1)
        EnforceConstraints(True)
    End Sub
End Class