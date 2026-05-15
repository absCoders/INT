Imports System.Math

Public Class SORRTRN1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim SOTRTRN1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("ICTPARM1")
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

        Dim sqlw As String = "SOTRTRN1.REGISTER_IND = '0'"

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Returns Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Returns Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "SOTRTRN1.RTRN_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
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
            sqlw = "SOTRTRN1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        If ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTRTRN1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
        If Absx1.optFor("RANGE").Value = "N" Then
            Call Print_GL()
        End If
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

        Dim sql As String = "Update SOTRTRN1 " _
        & " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2" _
        & " where ADJ_NO in (Select ADJ_NO from " & SOTRTRN1 & " )"
        ASCDATA1.ExecuteSQL(sql, "VV", New Object() {"1", MyBase.XNO})

        Call GL_Update()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"
        ASCMAIN1.sql = "Select * from SOTRTRN1 where " & sqlw
        SOTRTRN1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRTRN1 & " Add Primary Key (RTRN_NO)")

        ASCMAIN1.sql = "Select SOTRTRN1.*,ICTREAS1.REASON_DESC " _
        & " from " & SOTRTRN1 & " SOTRTRN1,ICTREAS1 " _
        & " where ICTREAS1.REASON_CODE (+) = SOTRTRN1.REASON_CODE " _
        & "   and " & sqlw
        Call Create_TDA(dst.Tables.Add, "SOTRTRN1", "**", 0)

        ASCMAIN1.sql = "Select SOTRTRN2.*, ICTITEM1.ITEM_DESC " _
        & " from SOTRTRN2," & SOTRTRN1 & " SOTRTRN1, ICTITEM1 " _
        & " where SOTRTRN2.RTRN_NO = SOTRTRN1.RTRN_NO" _
        & " and ICTITEM1.ITEM_CODE = SOTRTRN2.ITEM_CODE"
        Call Create_TDA(dst.Tables.Add, "SOTRTRN2", "**", 0)

        ASCMAIN1.sql = "Select SOTRTRN3.*, GLTACCT1.ACCT_DESC " _
        & " from SOTRTRN3," & SOTRTRN1 & " SOTRTRN1, GLTACCT1 " _
        & " where SOTRTRN3.RTRN_NO = SOTRTRN1.RTRN_NO" _
        & " and GLTACCT1.ACCT_CODE = SOTRTRN3.ACCT_CODE"
        Call Create_TDA(dst.Tables.Add, "SOTRTRN3", "**", 0)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("SOTRTRN1")
        Fill_Records("SOTRTRN2")
        Fill_Records("SOTRTRN3")
        If RWU = "R" Then
            TAC.ICCMAIN1.Prepare_GL_Interface("ICIR", SOTRTRN1)
        End If
        EnforceConstraints(True)
    End Sub
End Class