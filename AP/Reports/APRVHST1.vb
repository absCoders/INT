Imports System.Math

Public Class APRVHST1
    Dim Report_Subt As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Call Get_PARM("GLTPARM1")
        'Call Get_PARM("APTPARM1")

        Range_Events(grpINV_DATE_RANGE)
        Range_Events(grpCHK_DATE_RANGE)
    End Sub

    Protected Overrides Sub Build_Workfile()

        'xINV_DTE0 = Nothing
        'If Absx1.dteFor("INVDTE0").Value & "" <> "" Then
        '    xINV_DTE0 = Absx1.dteFor("INVDTE0").Value
        'End If


        With dst
            Dim sql As String = ""

            sql &= " From APTINVH1,APTVEND1"
            sql &= " where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE"

            If Not Absx1.chkFor("CHKCHK_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
                sql &= "   and APTINVH1.CHECK_DATE  >= '" & z & "' "
                Report_Subt &= " Check Date From " & z

            End If
            If Not Absx1.chkFor("CHKCHK_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
                sql &= "   and APTINVH1.CHECK_DATE <= '" & z & "'"
                Report_Subt &= " Check Date to " & z
            End If

            If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
                sql &= "   and APTINVH1.INV_DATE  >= '" & z & "' "
                Report_Subt &= " Invoice Date From " & z
            End If
            If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
                sql &= "   and APTINVH1.INV_DATE <= '" & z & "'"
                Report_Subt &= " Invoice Date to " & z
            End If

            sql &= SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")

            ASCMAIN1.sql = "SELECT APTINVH1.* " & sql
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTINVH1"))

            ASCMAIN1.sql = "SELECT DISTINCT APTVEND1.*" & sql
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTVEND1"))
        End With

        Check_if_Empty("APTINVH1")

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , Report_Subt)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub


End Class