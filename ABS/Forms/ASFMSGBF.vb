Imports Infragistics.Win.UltraWinGrid

Public Class ASFMSGBF
    Public user_option As Integer = 0
    Public grow As UltraWinGrid.UltraGridRow

    Public grdFilter As Boolean = False
    Public grdGroupBy As Boolean = False

    Private f_Calling_Form As ASFBASE0
    Dim tblX As DataTable
    Public gridColumnCaptions As String = String.Empty

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If optmsg.Visible Then
            user_option = optmsg.CheckedIndex
            If user_option = -1 Then
                Exit Sub
            End If
        End If
        If grdmsg.Visible Then
            If grdmsg.ActiveRow IsNot Nothing Then
                grow = grdmsg.ActiveRow
            End If
        End If
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        user_option = -1
        txtmsg.Text = ""
        Me.Close()
    End Sub

    Public Sub Show_grd(
    ByVal tbl As DataTable,
    ByVal ff As ASFBASE1,
    Optional ByVal grdcaption As String = "",
    Optional ByVal grdCode As String = "")

        grow = Nothing

        grdmsg.DataSource = tbl
        tblX = tbl

        grdmsg.Visible = True
        cmdOK.Visible = True
        lblmsg.Visible = False

        Me.Height = Me.Height * 4

        grdmsg.Top = 0
        grdmsg.Height = cmdOK.Top * 0.95
        grdmsg.Dock = DockStyle.Top


        f_Calling_Form = ff

        grdmsg.DataBind()

        'f_Calling_Form = ff

        If grdFilter Then
            f_Calling_Form.Show_Filter(grdmsg, True)
        End If
        If grdGroupBy Then
            grdmsg.DisplayLayout.GroupByBox.Hidden = False
        End If

        Dim W As Int32 = 50
        For Each C As UltraWinGrid.UltraGridColumn In grdmsg.DisplayLayout.Bands(0).Columns
            W += C.Width
        Next
        If W > ff.Width * 0.8 Then
            W = ff.Width * 0.8
        End If
        If Me.Width < W Then
            Me.Width = W
        End If

        Me.Text = grdcaption

        f_Calling_Form.Format_grd_ASFMSGBF(grdmsg, grdCode)
        f_Calling_Form.Sort_grdColumns(grdmsg)

        If gridColumnCaptions.Length > 0 Then
            Dim icol As Int16 = 0
            For Each gridColumnCaption As String In gridColumnCaptions.Split(",")
                gridColumnCaption = gridColumnCaption.Trim
                If gridColumnCaption.Length > 0 Then
                    If grdmsg.DisplayLayout.Bands(0).Columns.Count > icol Then
                        grdmsg.DisplayLayout.Bands(0).Columns(icol).Header.Caption = gridColumnCaption
                    End If
                End If

                icol += 1
            Next
        End If

        Me.ShowDialog()

    End Sub

    Public Sub Show_Formatted_txt(
    ByVal Form_Caption As String,
    ByVal Formatted_Text As String,
    ByVal ff As Form)

        'f_Calling_Form = ff

        Me.Text = Form_Caption
        Me.Height = 0.6 * ff.Height
        Me.Width = 0.6 * ff.Width

        'fmttxt.Value = Formated_Text
        'fmttxt.ReadOnly = True
        'fmttxt.Appearance.BackColor = Color.White ' Color.Empty ' Color.White
        'fmttxt.Visible = True

        WebBrowser1.DocumentText = Formatted_Text
        WebBrowser1.Visible = True

        cmdOK.Visible = False
        cmdCancel.Text = "Done"
        cmdPrint.Visible = True

        'fmttxt.Appearance.BackColorDisabled = Color.Empty ' Color.White
        'fmttxt.Appearance.ForeColorDisabled = Color.Empty ' Color.DodgerBlue
        'fmttxt.
        lblmsg.Visible = False


        Me.ShowDialog()
    End Sub

    Public Function Get_txtblock_from_User(
    ByVal Label_Text As String,
    ByVal Form_Caption As String,
    Optional ByVal default_value As String = "",
    Optional ByVal read_only As Boolean = False,
    Optional ByVal maxlength As Int16 = 0,
    Optional ByVal spellchecker As Boolean = False) As String

        txtmsg.MaxLength = maxlength
        txtmsg.Visible = True
        txtmsg.Text = ""

        If spellchecker Then
            txtmsg.SpellChecker = ASFMAIN1.UltraSpellChecker1
        Else
            txtmsg.SpellChecker = Nothing
        End If

        txtmsg.ReadOnly = read_only

        If default_value <> "" Then
            txtmsg.Text = default_value
        End If
        txtmsg.Multiline = True
        Me.Height += txtmsg.Height * 5
        txtmsg.Height = txtmsg.Height * 6
        'Me.AcceptButton = Nothing

        Dim lblHeight As Long = lblmsg.Height
        lblmsg.Text = Label_Text

        lblHeight = lblmsg.Height - lblHeight
        If lblHeight > 0 Then
            Me.txtmsg.Top += lblHeight
            'Me.cmdCancel.Top += lblHeight
            'Me.cmdOK.Top += lblHeight
            Me.Height += (lblHeight * 2)
        End If

        Me.Text = Form_Caption
        Me.ShowDialog()

        Return txtmsg.Text
    End Function

    Public Function Get_numint_from_User(
    ByVal Label_Text As String,
    ByVal Form_Caption As String,
    Optional ByVal maxValue As Long = 2147483647,
    Optional ByVal minValue As Long = -2147483648,
    Optional ByVal defaultValue As Integer = 0) As Integer

        nummsg.Visible = True
        nummsg.Value = defaultValue
        nummsg.NumericType = UltraWinEditors.NumericType.Integer

        If maxValue <> 0 Then
            nummsg.MaxValue = maxValue
        End If
        If minValue <> 0 Then
            nummsg.MinValue = minValue
        End If

        Dim lblHeight As Long = lblmsg.Height
        lblmsg.Text = Label_Text

        lblHeight = lblmsg.Height - lblHeight
        If lblHeight > 0 Then
            Me.txtmsg.Top += lblHeight
            Me.Height += (lblHeight * 2)
        End If

        Me.Text = Form_Caption
        Me.ShowDialog()

        Return nummsg.Value

    End Function

    Public Function Get_numdec_from_User(
    ByVal Label_Text As String,
    ByVal Form_Caption As String,
    Optional ByVal maxValue As Long = 2147483647,
    Optional ByVal minValue As Long = -2147483648,
    Optional ByVal defaultValue As Integer = 0) As Decimal

        nummsg.Visible = True
        nummsg.Value = defaultValue
        nummsg.NumericType = UltraWinEditors.NumericType.Decimal

        If maxValue <> 0 Then
            nummsg.MaxValue = maxValue
        End If
        If minValue <> 0 Then
            nummsg.MinValue = minValue
        End If

        Dim lblHeight As Long = lblmsg.Height
        lblmsg.Text = Label_Text

        lblHeight = lblmsg.Height - lblHeight
        If lblHeight > 0 Then
            Me.txtmsg.Top += lblHeight
            Me.Height += (lblHeight * 2)
        End If

        Me.Text = Form_Caption
        Me.ShowDialog()

        Return nummsg.Value

    End Function

    Public Function Get_numdouble_from_User(
        ByVal Label_Text As String,
        ByVal Form_Caption As String,
        Optional ByVal maxValue As Long = 2147483647,
        Optional ByVal minValue As Long = -2147483648,
        Optional ByVal defaultValue As Integer = 0) As Decimal

        nummsg.Visible = True
        nummsg.Value = defaultValue
        nummsg.NumericType = UltraWinEditors.NumericType.Double

        If maxValue <> 0 Then
            nummsg.MaxValue = maxValue
        End If
        If minValue <> 0 Then
            nummsg.MinValue = minValue
        End If

        Dim lblHeight As Long = lblmsg.Height
        lblmsg.Text = Label_Text

        lblHeight = lblmsg.Height - lblHeight
        If lblHeight > 0 Then
            Me.txtmsg.Top += lblHeight
            Me.Height += (lblHeight * 2)
        End If

        Me.Text = Form_Caption
        Me.ShowDialog()

        Return nummsg.Value

    End Function

    Public Function Get_txt_from_User(
    ByVal Label_Text As String,
    ByVal Form_Caption As String,
    Optional ByVal password As Boolean = False,
    Optional ByVal maxLength As Long = 0,
    Optional ByVal defaultValue As String = "") As String

        txtmsg.Visible = True
        txtmsg.Text = ""

        If maxLength > 0 Then
            txtmsg.MaxLength = maxLength
        End If

        If defaultValue.Trim.Length > 0 Then
            txtmsg.Text = defaultValue.Trim
        End If

        If password Then
            txtmsg.PasswordChar = "*"
        End If

        Dim lblHeight As Long = lblmsg.Height
        lblmsg.Text = Label_Text

        lblHeight = lblmsg.Height - lblHeight
        If lblHeight > 0 Then
            Me.txtmsg.Top += lblHeight
            'Me.cmdCancel.Top += lblHeight
            'Me.cmdOK.Top += lblHeight
            Me.Height += (lblHeight * 2)
        End If

        Me.Text = Form_Caption
        Me.ShowDialog()

        Return txtmsg.Text
    End Function

    Public Function Get_opt_from_User(
    ByVal Label_Text As String, ByVal options() As String,
    ByVal default_option As Integer,
    ByVal Form_Caption As String) As Integer

        Dim opt_height As Double = optmsg.Height
        optmsg.Visible = True

        lblmsg.AutoSize = True

        lblmsg.Text = Label_Text
        optmsg.Top = lblmsg.Top + lblmsg.Height + 10

        Me.Text = Form_Caption
        For i As Integer = 0 To UBound(options)
            Dim optitem As New Infragistics.Win.ValueListItem
            If i > 0 Then
                optmsg.Items.Add(optitem)
            End If
            optmsg.Items(i).DisplayText = options(i)
            '            optitem.DisplayText = options(i)
        Next

        optmsg.CheckedIndex = default_option
        If default_option = -1 Then
            cmdOK.Enabled = False
        End If

        Me.Height += opt_height * (optmsg.Items.Count)
        Me.ShowDialog()

        Return user_option
    End Function

    Private Sub ASFMSGBF_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ASCMAIN1.Center(Me)
        grdmsg.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        grdmsg.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        ASCMAIN1.grdInitializeLayout(grdmsg)

        If txtmsg.Multiline Then
            Me.AcceptButton = Nothing
        End If
        'txtmsg.SpellChecker = ASFMAIN1.UltraSpellChecker1
        ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(txtmsg, "txtMenu")


        ' fmttxt.Visible = False
        grdmsg.Parent = SplitContainer1.Panel1
        grdmsg.Dock = DockStyle.Fill
        SplitContainer1.FixedPanel = FixedPanel.Panel2
        If img1.Image Is Nothing Then
            img1.Dock = DockStyle.None
            img1.Visible = False
        Else
            img1.Dock = DockStyle.Fill
            img1.Visible = True
            cmdPrint.Visible = True
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        End If
    End Sub

    Private Sub optmsg_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optmsg.ValueChanged
        If optmsg.CheckedIndex <> -1 Then
            cmdOK.Enabled = True
        End If
    End Sub

    Private Sub txtmsg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmsg.KeyPress
        If e.KeyChar = vbCr Then
            If Not txtmsg.Multiline Then
                cmdOK.PerformClick()
            End If
        End If
    End Sub

    Private Sub grdmsg_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdmsg.DoubleClickRow
        grow = grdmsg.ActiveRow
        Me.Close()
    End Sub

    Private Sub grdmsg_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdmsg.MouseDown

        Try
            Dim pt As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
            Dim elem As Infragistics.Win.UIElement
            elem = grdmsg.DisplayLayout.UIElement.ElementFromPoint(pt)
            If elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.RowSelectorHeaderUIElement)) Then
                'f_Calling_Form.Excel_Export(grdmsg)
                'dup export              f_Calling_Form.Export_to_Excel(grdmsg, True, False, Me.Text, "A")
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Dim PrintDialog1 As New System.Windows.Forms.PrintDialog

        Dim result As DialogResult = PrintDialog1.ShowDialog()

        If result = Windows.Forms.DialogResult.OK Then
            If img1 Is Nothing Then
                WebBrowser1.Print()
            Else
                Try
                    SplitContainer1.Panel2Collapsed = True

                    img = CaptureForm1()
                    pd = New System.Drawing.Printing.PrintDocument

                    'pd.Print()

                    Dim ppDialog As PrintPreviewDialog = New PrintPreviewDialog()
                    ppDialog.ClientSize = New Size(400, 500)
                    ppDialog.Document = pd
                    pd.Print()
                    '  ppDialog.ShowDialog()

                    pd = Nothing
                    img = Nothing

                    SplitContainer1.Panel2Collapsed = False
                Catch ex As Exception
                    MsgBox("Please verify that the Default Printer is available", MsgBoxStyle.OkOnly, "Cannot Select Default Printer")
                End Try
            End If

        End If
    End Sub

    Private Sub nummsg_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles nummsg.MouseUp
        nummsg.SelectAll()
    End Sub

    Public Sub Show_img(
    ByVal img As Image,
    ByVal ff As ASFBASE1,
    Optional ByVal frmcaption As String = "")

        img1.Image = img
        img1.Visible = True
        Me.Height = img.Height + SplitContainer1.Height - SplitContainer1.Panel1.Height + 100
        Me.Width = img.Width

        If Me.Height > ASFMAIN1.Height * 0.8 Then
            Me.Height = Me.Height * 0.8
            Me.Width = Me.Width * 0.8
        End If

        If Me.Width > ASFMAIN1.Width * 0.8 Then
            Me.Height = Me.Height * 0.8
            Me.Width = Me.Width * 0.8
        End If

        img1.Dock = DockStyle.Fill
        img1.ScaleImage = ScaleImage.Always

        cmdOK.Visible = True
        lblmsg.Visible = False
        'lblmsg.Visible = True
        'lblmsg.Text = frmcaption
        'lblmsg.Dock = DockStyle.Top
        '   Me.Height = Me.Height * 4
        f_Calling_Form = ff
        Me.Text = frmcaption
        Me.ShowDialog()

    End Sub

    Private Sub nummsg_ValidationError(sender As Object, e As Infragistics.Win.UltraWinEditors.ValidationErrorEventArgs) Handles nummsg.ValidationError
        e.RetainFocus = False
    End Sub


    'Imports System.Runtime.InteropServices
    ' Global Variables 

    Dim img As Bitmap
    Dim WithEvents pd As System.Drawing.Printing.PrintDocument

    'Returns the Form as a bitmap
    Public Function CaptureForm1() As Bitmap

        Dim g1 As Graphics = Me.CreateGraphics()
        Dim MyImage = New Bitmap(Me.ClientRectangle.Width, Me.ClientRectangle.Height, g1)

        Dim g2 As Graphics = Graphics.FromImage(MyImage)
        Dim dc1 As IntPtr = g1.GetHdc()
        Dim dc2 As IntPtr = g2.GetHdc()
        BitBlt(dc2, 0, 0, Me.ClientRectangle.Width, (Me.ClientRectangle.Height), dc1, 0, 0, 13369376)
        g1.ReleaseHdc(dc1)
        g2.ReleaseHdc(dc2)
        'saves image to c drive just, u can comment it also
        'MyImage.Save("c:\abc.bmp")
        Return MyImage
    End Function

    <System.Runtime.InteropServices.DllImport("gdi32.DLL", EntryPoint:="BitBlt",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode,
    ExactSpelling:=True,
    CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)>
    Private Shared Function BitBlt(ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As System.Int32) As Boolean

        ' Leave function empty - DLLImport attribute forwards calls to MoveFile to
        ' MoveFileW in KERNEL32.DLL.
    End Function

    Private Sub pd_QueryPageSettings(ByVal sender _
    As Object, ByVal e As _
    System.Drawing.Printing.QueryPageSettingsEventArgs) _
    Handles pd.QueryPageSettings
        e.PageSettings.Landscape = True
    End Sub

    'this method will be called each time when pd.printpage event occurs
    Sub pd_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles pd.PrintPage

        Dim x As Integer = e.MarginBounds.X '/ 2
        Dim y As Integer = e.MarginBounds.Y '/ 2
        'e.Graphics.DrawImage(img, x, y)

        'e.Graphics.DrawImage(img, 0, 0)

        'e.HasMorePages = False



        'local scope
        Dim mySource As Rectangle
        'Dim myDestination As Rectangle

        'define a rectangle as the size of the original image (source)
        mySource = New Rectangle(x:=x, y:=y, width:=e.MarginBounds.Width, height:=e.MarginBounds.Height)

        'draw the original bitmap to the source rectangle
        e.Graphics.DrawImage(image:=img, rect:=mySource)

        Dim ABS_logo As Bitmap = New System.Drawing.Bitmap(filename:=ASCMAIN1.Folders("Images") & "abs\abs_logo.jpg")

        e.Graphics.DrawImage(image:=ABS_logo, rect:=New Rectangle(x:=0, y:=0, width:=80, height:=40))
        Dim p As New Pen(Color.Blue, 1)
        e.Graphics.DrawLine(p, 0, 45, e.PageBounds.Width, 45)


        Dim printFont As Font = Me.Font
        Dim myBrush As New SolidBrush(Color.Black)

        e.Graphics.DrawString(Me.Text & " " & Format(Now, "MM/dd/yyyy hh:mm tt") & " " & ASCMAIN1.USER_ID, printFont, myBrush, 0, e.PageBounds.Height - e.MarginBounds.Y + 20, New StringFormat())

    End Sub

    Private Sub grdmsg_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdmsg.InitializeRow
        If e.Row.IsDataRow Then
            If tblX.Columns.Contains("TABLE_NAME") And tblX.Columns.Contains("COLUMN_NAME") Then
                Dim TABLE_NAME As String = e.Row.Cells("TABLE_NAME").Value
                Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Value
                Dim rowASTAESC1 As DataRow = ASCMAIN1.tblASTAESC1.Rows.Find(New String() {TABLE_NAME, COLUMN_NAME})
                If rowASTAESC1 IsNot Nothing Then
                    Dim TXTEDITOR As New Infragistics.Win.UltraWinEditors.UltraTextEditor
                    TXTEDITOR.PasswordChar = "*"
                    e.Row.Cells("OLD_VALUE").EditorComponent = TXTEDITOR
                    e.Row.Cells("NEW_VALUE").EditorComponent = TXTEDITOR
                End If

            Else
                If f_Calling_Form IsNot Nothing AndAlso tblX.TableName = "ASTAUDT1" Then
                    Dim TABLE_NAME As String = f_Calling_Form.TABLE_NAME
                    Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Value
                    Dim rowASTAESC1 As DataRow = ASCMAIN1.tblASTAESC1.Rows.Find(New String() {TABLE_NAME, COLUMN_NAME})
                    If rowASTAESC1 IsNot Nothing Then
                        Dim TXTEDITOR As New Infragistics.Win.UltraWinEditors.UltraTextEditor
                        TXTEDITOR.PasswordChar = "*"
                        e.Row.Cells("OLD_VALUE").EditorComponent = TXTEDITOR
                        e.Row.Cells("NEW_VALUE").EditorComponent = TXTEDITOR
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub Show_Xml_File(ByVal Form_Caption As String, ByVal XMLFilename As String, ByVal ff As Form)
        Try
            Me.Text = Form_Caption
            Me.Height = 0.6 * ff.Height
            Me.Width = 0.6 * ff.Width

            WebBrowser1.Navigate(XMLFilename)
            WebBrowser1.Visible = True

            cmdOK.Visible = False
            cmdPrint.Visible = True
            lblmsg.Visible = False
            Me.ShowDialog()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Show File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class