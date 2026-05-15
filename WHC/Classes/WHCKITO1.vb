Public Class WHCKITO1

    ' Create Kittimg Request

    ' CREATE TABLE BMTMNLP1 (
    ' LP_CODE VARCHAR2(6),
    ' BM_PROD_ITEM VARCHAR2(22),
    ' BM_ISSUE_NO VARCHAR2(2),
    ' INIT_OPER  VARCHAR2(20),
    ' INIT_XMIT DATE,
    ' PRIMARY KEY (LP_CODE, BM_PROD_ITEM, BM_ISSUE_NO));

    'INSERT INTO BMTMNLP1
    'SELECT 'ADS', BM_PROD_ITEM, BM_ISSUE_NO, 'conv', SYSDATE
    'FROM BMTMAIN2

    Inherits WHC000O1

    Private lstFilesToSend As New List(Of String)
    Private BM_ISSUE_NO As String = "00"
    Private PlannedQuantity As Int32 = 0
    Private PO_ORDER_NO As String = String.Empty

    Private BM_PROD_ITEM As String = String.Empty

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHCKITO1"

        Create_Work_Table()

        With dst
            ASCMAIN1.sql = "SELECT BMTMAIN1.*, ICTITEM1.ITEM_DESC
                                FROM BMTMAIN1, ICTITEM1 
                                WHERE BMTMAIN1.BM_PROD_ITEM = :PARM1
                                AND BMTMAIN1.BM_PROD_ITEM = ICTITEM1.ITEM_CODE (+)"
            Create_TDA(.Tables.Add, "BMTMAIN1", ASCMAIN1.sql, 0, True, "V", 1)
            Create_TDA(.Tables.Add, "BMTMAIN2", "*", 2)
            Create_TDA(.Tables.Add, "BMTMAIN3", "*", 2)
            Create_TDA(.Tables.Add, "BMTMNLP1", "*")
            Create_TDA(.Tables.Add, "POTORDR1", "*")
        End With

        Main_Process()
    End Sub

    Public Sub Main_Process()

        EnforceConstraints(False)

        Dim BM_PROD_ITEM As String = G.APP_KEY.Split(",")(0)
        Dim BM_ISSUE_STATUS As String = G.APP_KEY.Split(",")(1)
        PlannedQuantity = Val(G.APP_KEY.Split(",")(2) & String.Empty)
        PO_ORDER_NO = G.APP_KEY.Split(",")(3)

        Fill_Records("BMTMAIN1", {BM_PROD_ITEM})
        Dim rowBMTMAIN1 As DataRow = dst.Tables("BMTMAIN1").Rows.Find(BM_PROD_ITEM)
        Dim BM_ISSUE_COUNTER As Integer = Val(rowBMTMAIN1.Item("BM_ISSUE_COUNTER") & "")
        If BM_ISSUE_STATUS = "C" Then BM_ISSUE_NO = Format(BM_ISSUE_COUNTER, "00")
        If BM_ISSUE_STATUS = "S" Then BM_ISSUE_NO = rowBMTMAIN1.Item("BM_ISSUE_STD") & ""

        Fill_Records("BMTMAIN2", {BM_PROD_ITEM, BM_ISSUE_NO})
        Fill_Records("BMTMAIN3", {BM_PROD_ITEM, BM_ISSUE_NO})

        Fill_Records("BMTMNLP1", {"ADS", BM_PROD_ITEM, BM_ISSUE_NO})
        Fill_Records("POTORDR1", PO_ORDER_NO)

        EnforceConstraints(True)

        Update_Record()

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Select Case LP_CODE

            Case "ADS"
                Dim xsdFolder As String = ASCMAIN1.Folders("SharedRoot") & "XSDs\"

                If ASCMAIN1.Running_in_VS Then
                    Stop
                    xsdFolder = "C:\SFTP\cfg\PROD\XSDs\"
                End If

                Dim dstWWIMPZMFG As New DataSet
                dstWWIMPZMFG.ReadXmlSchema(xsdFolder & "WWIMPZMFG" & ".XSD")
                'dstWWIMPZMFG.Tables("Line").Columns("PlannedQuantity").DataType = GetType(Int32)
                dstWWIMPZMFG.EnforceConstraints = False

                Dim dstWWIMPZBOMP As New DataSet
                dstWWIMPZBOMP.ReadXmlSchema(xsdFolder & "WWIMPZBOMP" & ".XSD")
                'dstWWIMPZMFG.Tables("Line").Columns("LinkQuantity").DataType = GetType(Int32)
                dstWWIMPZBOMP.EnforceConstraints = False

                Dim drPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)

                For Each drBMTMAIN1 As DataRow In dst.Tables("BMTMAIN1").Select("", "")
                    BM_PROD_ITEM = drBMTMAIN1.Item("BM_PROD_ITEM") & String.Empty
                    Dim drBMTMAIN2 As DataRow = dst.Tables("BMTMAIN2").Rows.Find({BM_PROD_ITEM, BM_ISSUE_NO})

                    Dim drLine As DataRow = dstWWIMPZMFG.Tables("Line").NewRow
                    drLine.Item("ProductionSite") = "S1"
                    drLine.Item("BomCode") = "2"
                    drLine.Item("Product") = BM_PROD_ITEM
                    drLine.Item("StartDate") = CDate(drPOTORDR1.Item("PO_DATE_REQUIRED") & String.Empty).ToString("yyyy-MM-dd")
                    drLine.Item("EndDate") = CDate(drPOTORDR1.Item("PO_DATE_CANCEL") & String.Empty).ToString("yyyy-MM-dd")
                    drLine.Item("PlannedQuantity") = PlannedQuantity
                    drLine.Item("WorkOrderNumber") = PO_ORDER_NO & "_" & System.DateTime.Now.ToString("yyyyMMddhhmmss") ' Unique value
                    drLine.Item("RoutingNumber") = "STDROUTING"
                    drLine.Item("RoutingCode") = "1"
                    drLine.Item("UpdateCode") = "C"
                    drLine.Item("OrderStatus") = "1"
                    dstWWIMPZMFG.Tables("Line").Rows.Add(drLine)

                    Dim BM_ISSUE_TEXT As String = drBMTMAIN2.Item("BM_ISSUE_TEXT") & ""
                    BM_ISSUE_TEXT = BM_ISSUE_TEXT.Replace(vbCr, "").Replace(vbLf, "")
                    While BM_ISSUE_TEXT.Contains(Space(2))
                        BM_ISSUE_TEXT = BM_ISSUE_TEXT.Replace(Space(2), Space(1))
                    End While

                    Dim drHeader As DataRow = dstWWIMPZBOMP.Tables("Header").NewRow
                    drHeader.Item("Product") = BM_PROD_ITEM
                    drHeader.Item("Description") = drBMTMAIN1.Item("ITEM_DESC")

                    ' 2025/10/31 - as per Stephanie (As per ADS) do not send the Headertext. It will appear in the email.
                    drHeader.Item("HeaderText1") = "" ' If(BM_ISSUE_TEXT.Length > 0, BM_ISSUE_TEXT.Substring(0, Math.Min(100, BM_ISSUE_TEXT.Length)), "")
                    drHeader.Item("HeaderText2") = "" ' If(BM_ISSUE_TEXT.Length > 100, BM_ISSUE_TEXT.Substring(100, Math.Min(100, BM_ISSUE_TEXT.Length - 100)), "")
                    drHeader.Item("HeaderText3") = "" ' If(BM_ISSUE_TEXT.Length > 200, BM_ISSUE_TEXT.Substring(200, Math.Min(100, BM_ISSUE_TEXT.Length - 200)), "")
                    drHeader.Item("ManagementUnit") = "One"
                    drHeader.Item("BomCode") = "2"
                    drHeader.Item("UseStatus") = "2"
                    drHeader.Item("header_id") = R
                    dstWWIMPZBOMP.Tables("Header").Rows.Add(drHeader)

                    Dim Sequence As Int16 = 1
                    For Each drBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select($"BM_PROD_ITEM = '{BM_PROD_ITEM}' and BM_ISSUE_NO = {BM_ISSUE_NO}", "BM_SEQ")
                        drLine = dstWWIMPZBOMP.Tables("Line").NewRow
                        drLine.Item("Sequence") = Sequence
                        Sequence += 1
                        drLine.Item("ComponentProduct") = drBMTMAIN3.Item("BM_COMP_ITEM")
                        drLine.Item("ComponentType") = "1"
                        drLine.Item("LineText1") = ""
                        drLine.Item("LineText2") = ""
                        drLine.Item("LineText3") = ""
                        drLine.Item("LinkQuantity") = drBMTMAIN3.Item("BM_QTY_PER_ASSY")
                        drLine.Item("LinkQuantityCode") = "1"
                        drLine.Item("header_id") = R
                        dstWWIMPZBOMP.Tables("Line").Rows.Add(drLine)
                    Next

                    R += 1
                Next

                If R > 0 Then
                    ' This is the Item information. Only send this once.
                    If dst.Tables("BMTMNLP1").Rows.Count = 0 Then
                        FILENAME_TO_SEND = $"WWIMPZMFG_{XMIT_NO}.XML"
                        lstFilesToSend.Add(FILENAME_TO_SEND)
                        dstWWIMPZMFG.WriteXml(FILENAME_TO_SEND)
                    End If

                    FILENAME_TO_SEND = $"WWIMPZBOMP_{XMIT_NO}.XML"
                    lstFilesToSend.Add(FILENAME_TO_SEND)
                    dstWWIMPZBOMP.WriteXml(FILENAME_TO_SEND)
                End If
        End Select

    End Sub

    Overrides Sub Update_Archive()

        ' Added 09/26/2025 to prevent sending test environment data
        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                MessageBox.Show("You are not in production. File transfer avoided.", "Update Archive", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        MyBase.Update_Archive()

        Dim subDir As String = DateTime.Now.ToString("yyyyMM")
        ' If the Archive direcory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir) Then
            My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir)
        End If

        Select Case LP_CODE

            Case "ADS"

                If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
                    ' do not send
                Else
                    For Each FILENAME_TO_SEND As String In lstFilesToSend
                        sftp_put("ADS", True, FILENAME_TO_SEND, FILENAME_TO_SEND)
                    Next
                End If

                For Each FILENAME_TO_SEND As String In lstFilesToSend
                    My.Computer.FileSystem.MoveFile(FILENAME_TO_SEND, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & FILENAME_TO_SEND)
                Next

                Try
                    If dst.Tables("BMTMNLP1").Rows.Count = 0 Then
                        Dim sql = "INSERT INTO BMTMNLP1
                        (LP_CODE, BM_PROD_ITEM, BM_ISSUE_NO, INIT_OPER, INIT_XMIT)
                        values
                        (:PARM1, :PARM2, :PARM3, :PARM4, SYSDATE)"

                        ASCDATA1.ExecuteSQL(sql, "VVVV", {LP_CODE, BM_PROD_ITEM, BM_ISSUE_NO, ASCMAIN1.USER_ID})
                    End If
                Catch ex As Exception

                End Try

        End Select

    End Sub

    Sub Create_Work_Table()

    End Sub

End Class