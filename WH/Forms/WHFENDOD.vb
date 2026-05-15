Imports ABSolution

Public Class WHFENDOD

    Private wktable As String = String.Empty
    Private ShippingLabelDirectory As String
    Private clsShip As TAC.WHCSHIP1
    Private manifestData As String = String.Empty
    Private linesPerPage As Int16 = 8

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            wktable = ASCMAIN1.Temp_Table("Select Ship_cntl_no from whtship1 where rownum < 1")

            ASCMAIN1.sql = "SELECT WHTSHIP1.*, SOTCARR2.CARRIER_PROD_DESC, " _
                & " WHTSHIP5.SHIP_FIRST_NAME, WHTSHIP5.SHIP_LAST_NAME, WHTSHIP5.SHIP_ADDR1, WHTSHIP5.SHIP_ADDR2, WHTSHIP5.SHIP_CITY, WHTSHIP5.SHIP_STATE, WHTSHIP5.SHIP_ZIP_CODE, WHTSHIP5.SHIP_COMPANY" _
                & " FROM WHTSHIP1, SOTCARR2, " & wktable & " WHTSHIPX, WHTSHIP5" _
                & " WHERE WHTSHIP1.SHIP_CNTL_NO  = WHTSHIPX.SHIP_CNTL_NO" _
                & " AND TO_CHAR(WHTSHIP1.CARRIER_PROD_CODE)  = SOTCARR2.CARRIER_PROD_CODE (+)" _
                & " AND WHTSHIP1.CARRIER_CODE  = SOTCARR2.CARRIER_CODE (+)" _
                & " AND WHTSHIP1.SHIP_CNTL_NO  = WHTSHIP5.SHIP_CNTL_NO (+)" _
                & " AND WHTSHIP5.SHIP_ADDR_TYPE = 'ST'"

            Create_TDA(.Tables.Add, "WHTSHIP1", ASCMAIN1.sql, 0, False, "", 0)

            ASCMAIN1.sql = "SELECT WHTSHIP2.SHIP_CNTL_NO, WHTSHIP2.WEIGHT, WHTSHIP2.TRACKING_NO, WHTSHIP2.INV_NO, NVL(WHTSHIP2.CART_NO, WHTSHIP2.SHIP_PACKAGE_NO) CART_NO," _
                & "  DECODE(NVL(WHTSHIP1.INSURED_SHIPMENT, '0'), '0', 0, NVL(WHTSHIP2.INSURED_VALUE, 0)) INSURED_VALUE, WHTSHIP2.INV_BOL_NO," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.NET_CHARGE, 0), 0) NET_CHARGE," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.BASE_CHARGE, 0), 0) BASE_CHARGE," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.TOTAL_DISCOUNT, 0), 0) TOTAL_DISCOUNT," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.TOTAL_SURCHARGES, 0), 0) TOTAL_SURCHARGES," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.LIST_PRICE, 0), 0) LIST_PRICE," _
                & "  DECODE(WHTSHIP1.STATUS, 'P', NVL(WHTSHIP2.COD_AMOUNT, 0), 0) COD_AMOUNT" _
                & "  FROM WHTSHIP1, WHTSHIP2, " & wktable & " WHTSHIPX" _
                & "  WHERE WHTSHIP1.SHIP_CNTL_NO = WHTSHIP2.SHIP_CNTL_NO" _
                & "  AND WHTSHIP2.SHIP_CNTL_NO  = WHTSHIPX.SHIP_CNTL_NO"
            Create_TDA(.Tables.Add, "WHTSHIP2", ASCMAIN1.sql, 0, False, "", 0)

            .Relations.Add("WHTSHIP1_WHTSHIP2", dst.Tables("WHTSHIP1").Columns("SHIP_CNTL_NO"), dst.Tables("WHTSHIP2").Columns("SHIP_CNTL_NO"))
        End With

        grdWHTSHIPX.DataSource = dst.Tables("WHTSHIP1")
        Create_Summary(grdWHTSHIPX, "SHIP_CNTL_NO", "Count")
        Create_Summary(grdWHTSHIPX, "CART_NO", "Count", "WHTSHIP1_WHTSHIP2")

        For Each field As String In New String() {"BASE_CHARGE", "INSURED_VALUE", "NET_CHARGE", "TOTAL_DISCOUNT", "TOTAL_SURCHARGES", "LIST_PRICE"}
            Create_Summary(grdWHTSHIPX, field, "Sum", "WHTSHIP1_WHTSHIP2")
        Next

        dteShipDate.MaxDate = DateTime.Now
        dteShipDate.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)
        dteShipDate.Value = dteShipDate.MaxDate

        numHH.Value = DateTime.Now.ToString("HH")
        numMM.Value = DateTime.Now.ToString("mm")

        tabOptions_SelectedTabChanged(Nothing, Nothing)

        dteClose.MinDate = DateAdd(DateInterval.Day, -7, DateTime.Now)
        dteClose.MaxDate = DateAdd(DateInterval.Day, 7, DateTime.Now)

        Get_PARM("SOTPARM1")

        'If IsDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) AndAlso (ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) >= DateTime.Now Then
        'dteClose.Value = ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE")
        'Else
        dteClose.Value = DateTime.Now
        'End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CARRIER_CODE", False, True)
                If Not IsDate(dteShipDate.DateTime) Then
                    EMsg &= vbCr & "Invalid ship date"
                End If

            Case "Done"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Done"
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

                tabOptions.Tabs("Fedex Close").Visible = Not ScreenMode
                If Not ScreenMode Then
                    .Groups("Close Ground Shipments").Visible = False
                Else

                    .Groups("Close Ground Shipments").Visible = tabOptions.SelectedTab.Key = "Fedex Close"
                End If

            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)
        dst.Tables("WHTSHIP1").Rows.Clear()
        dst.Tables("WHTSHIP2").Rows.Clear()
        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        Try
            ASCMAIN1.Progress("Now Loading")

            MyBase.EnforceConstraints(False)

            ASCDATA1.ExecuteSQL("Truncate Table " & wktable)
            ASCMAIN1.sql = "Insert Into " & wktable & " SELECT SHIP_CNTL_NO FROM WHTSHIP1 WHERE TRUNC(SHIP_DATE) = '" & dteShipDate.DateTime.ToString("dd-MMM-yyyy") & "'"

            If Absx1.txtFor("CARRIER_CODE").TextLength > 0 Then
                ASCMAIN1.sql &= " and WHTSHIP1.CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'"
            End If

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            Fill_Records("WHTSHIP1")
            Fill_Records("WHTSHIP2")

            Clear_All_Filters(grdWHTSHIPX)
            Sort_grdColumns(grdWHTSHIPX, "SHIP_CNTL_NO", False, 0)
            Sort_grdColumns(grdWHTSHIPX, "CART_NO", False, 1)

            MyBase.EnforceConstraints(True)
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message, "Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("")
        End Try
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()


            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Form Controls"

    Private Sub btnOpenManifest_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenManifest.Click
        DisplayManifest("")
    End Sub

    Private Sub btnPrintManifest_Click(sender As System.Object, e As System.EventArgs) Handles btnPrintManifest.Click
        If txtManifest.Text.Length = 0 Then
            MessageBox.Show("There is nothing to print.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim PrinterName As String = String.Empty

        Dim printDialog1 As PrintDialog = New PrintDialog
        Dim result As DialogResult = printDialog1.ShowDialog(Me)
        If result = DialogResult.OK Then
            PrinterName = printDialog1.PrinterSettings.PrinterName
        Else
            Exit Sub
        End If

        manifestData = txtManifest.Text

        ' Set to our selected printer
        PrintDocument1.PrinterSettings.PrinterName = PrinterName
        ' This hides the print progress dialog
        PrintDocument1.PrintController = New System.Drawing.Printing.StandardPrintController()

        PrintDocument1.Print()

    End Sub

    Private Sub btnUpdateShipDate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdateShipDate.Click
        Get_PARM("SOTPARM1")

        If IsDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) Then
            If CDate(dteClose.Value) <= CDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) Then
                MessageBox.Show("You may not set the New Ship Date to a date less equal than the current setting of " & CDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) _
                                , "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        If MessageBox.Show("Do you want to set the current Ship Date to " & dteClose.DateTime.ToString("MM/dd/yyyy") & "?", "Ship Date" _
                           , MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If


        Try
            BeginTrans()
            ASCDATA1.ExecuteSQL("update sotparm1 set SO_PARM_WH_SHIP_DATE = '" & dteClose.DateTime.ToString("dd-MMM-yyyy") & "'")
            ASCDATA1.ExecuteSQL("update sotinvh1 set ORDR_WEB_IND = '2' where ORDR_WEB_IND = '1'")
            CommitTrans("Ship Date Changed")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Dim CARRIER_CODE As String = Absx1.txtFor("CARRIER_CODE").Text
        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTCARR1 where CARRIER_CODE = :PARM1", "V", New Object() {CARRIER_CODE})

        If rowSOTCARR1 Is Nothing Then
            MessageBox.Show("Invalid or missing Carrier.", "Close for the day", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If rowSOTCARR1.Item("PROVIDER_TYPE") <> "F" Then
            MessageBox.Show("The selected carrier is not a Federal Express account.", "Close for the day", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If MessageBox.Show("Do you want to Close the Fedex Shipments for the day?", "Close Fedex Shipments", _
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        CloseGroundShipments()

        Get_PARM("SOTPARM1")

        If IsDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) Then
            If CDate(DateTime.Now.ToString("MM/dd/yyyy")) >= CDate(ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") & String.Empty) Then
                If MessageBox.Show("Do you want to advance the shipment date?", "Shipment Date", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Dim daysToAdd As Int16 = 1
                Select Case DateTime.Now.ToString("ddd").ToUpper
                    Case "FRI"
                        daysToAdd = 3
                    Case "SAT"
                        daysToAdd = 2
                End Select

                dteClose.DateTime = DateAdd(DateInterval.Day, daysToAdd, CDate(DateTime.Now.ToString("MM/dd/yyyy")))
                Try
                    BeginTrans()
                    ASCDATA1.ExecuteSQL("update sotparm1 set SO_PARM_WH_SHIP_DATE = '" & dteClose.DateTime.ToString("dd-MMM-yyyy") & "'")
                    ASCDATA1.ExecuteSQL("update sotinvh1 set ORDR_WEB_IND = '2' where ORDR_WEB_IND = '1'")
                    CommitTrans("Ship Date Changed")
                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

            End If
        End If


    End Sub

    Private Sub tabOptions_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabOptions.SelectedTabChanged
        UltraExplorerBar1.Groups("Screen Control").Visible = tabOptions.SelectedTab.Key = "Shipments"
        UltraExplorerBar1.Groups("Close Ground Shipments").Visible = tabOptions.SelectedTab.Key = "Fedex Close"
    End Sub

    Private Sub CloseGroundShipments()

        Try
            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
            GetCredentials()

            If ShippingLabelDirectory.Length = 0 Then
                MessageBox.Show("You need to setup the Carrier Archive Directory directory in the Carrier Master", "Close Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim reportFile As String = ShippingLabelDirectory & "manifest_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
            With clsShip.FedexClose
                .Date = dteClose.Text
                .ReportFile = reportFile
                '.Time = Val(numHH.Value & String.Empty).ToString("#0") & ":" & Val(numMM.Value & String.Empty).ToString("00") & ":00"
            End With

            txtManifest.Clear()

            clsShip.FedexCloseGroundShipments()

            If My.Computer.FileSystem.FileExists(reportFile) Then
                DisplayManifest(reportFile)
            End If

            If clsShip.LastError.Length > 0 Then
                MessageBox.Show(clsShip.LastError, "Close Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Close Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DisplayManifest(ByVal manifestFile As String)

        Try

            If manifestFile.Length = 0 Then

                clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                GetCredentials()

                If ShippingLabelDirectory.Length = 0 Then
                    MessageBox.Show("You need to setup the Carrier Archive Directory directory in the Carrier Master", "Open Manifest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim openFileDialog1 As New OpenFileDialog

                openFileDialog1.InitialDirectory = ShippingLabelDirectory
                openFileDialog1.Title = "Open Manifest File"
                openFileDialog1.Filter = "Manifest files (*.txt)|*.txt"
                openFileDialog1.FilterIndex = 1
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    manifestFile = openFileDialog1.FileName
                Else
                    Exit Sub
                End If
            End If

            txtManifest.Clear()

            Dim sr As New IO.StreamReader(manifestFile)
            txtManifest.Text = sr.ReadToEnd
            sr.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub myDocument_PrintPage(sender As System.Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'e.Graphics.DrawString(txtManifest.Text, txtManifest.Font, Drawing.Brushes.Black, 25, 25)
        'e.Graphics.PageUnit = Drawing.GraphicsUnit.Inch


        Dim charactersOnPage As Int32 = 0
        Dim linesPerPage As Int16 = 0

        'Dim StringFormat As New System.Drawing.StringFormat

        '' Sets the value of charactersOnPage to the number of characters 
        '' of stringToPrint that will fit within the bounds of the page.
        'e.Graphics.MeasureString(manifestData, txtManifest.Font,
        '    e.MarginBounds.Size, Drawing.StringFormat.GenericTypographic,
        '      charactersOnPage, linesPerPage)

        Dim lines As String() = manifestData.Split(vbCr)
        Dim dataToPrint As String = String.Empty

        For Each chra As Char In manifestData
            If chra = "3" Then
                Stop
            End If
        Next


        ' Draws the string within the bounds of the page
        e.Graphics.DrawString(manifestData, txtManifest.Font, Drawing.Brushes.Black,
             25, 25) 'e.MarginBounds, Drawing.StringFormat.GenericTypographic)

        ' Remove the portion of the string that has been printed.
        manifestData = manifestData.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed.
        e.HasMorePages = (manifestData.Length > 0)
    End Sub

    Private Sub PrintDocument1_QueryPageSettings(sender As Object, e As System.Drawing.Printing.QueryPageSettingsEventArgs) Handles PrintDocument1.QueryPageSettings

        Dim margins As New System.Drawing.Printing.Margins
        margins.Bottom = 0.0
        margins.Left = 0.0
        margins.Right = 0.0
        margins.Top = 0.0

        e.PageSettings.Margins = margins
        e.PageSettings.Landscape = True

    End Sub

    Private Sub GetCredentials()

        Dim CARRIER_CODE As String = Absx1.txtFor("CARRIER_CODE").Text
        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTCARR1 where CARRIER_CODE = :PARM1", "V", New Object() {CARRIER_CODE})

        If rowSOTCARR1 Is Nothing Then
            MessageBox.Show("Invalid or missing Carrier.", "Close for the day", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If rowSOTCARR1.Item("PROVIDER_TYPE") <> "F" Then
            MessageBox.Show("The selected carrier is not a Federal Express account.", "Close for the day", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rowSOTCARR3 As DataRow = ASCDATA1.GetDataRow("Select * from SOTCARR3 where CARRIER_CODE = :PARM1", "V", New Object() {CARRIER_CODE})
        If rowSOTCARR3 Is Nothing Then
            MessageBox.Show("Invalid or missing Carrier Credentials.", "Close for the day", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        ShippingLabelDirectory = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

        Try
            If ASCMAIN1.Running_in_VS Then
                'ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
                ShippingLabelDirectory = ShippingLabelDirectory.Replace(ASCMAIN1.Folders("SharedRoot"), "N:\")
            End If
            If ShippingLabelDirectory.Length > 0 Then
                If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                End If
            End If
        Catch ex As Exception
            ShippingLabelDirectory = String.Empty
        End Try

        If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
            ShippingLabelDirectory = ShippingLabelDirectory & "\"
        End If

        ' Credentials
        clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
        clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
        clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
        clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
        clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
        clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
        clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
        clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

End Class