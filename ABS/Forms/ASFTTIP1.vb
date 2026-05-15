Imports Infragistics.Win.FormattedLinkLabel

Public Class ASFTTIP1
    Public rowASTTTIP1 As DataRow
    Dim FORM_NAME As String
    Dim TOOLTIP_TITLE As String
    Dim TOOLTIP_TEXT As String
    Public FP As ASFBASE1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        FORM_NAME = rowASTTTIP1.Item("FORM_NAME")
        TABLE_NAME = rowASTTTIP1.Item("TABLE_NAME")
        COLUMN_NAME = rowASTTTIP1.Item("COLUMN_NAME")

        TOOLTIP_TITLE = rowASTTTIP1.Item("TOOLTIP_TITLE") & ""
        TOOLTIP_TEXT = rowASTTTIP1.Item("TOOLTIP_TEXT") & ""

        With dst
            Create_TDA(.Tables.Add, "ASTSLID1", "*")
        End With

        Dim rowASTSLID1 As DataRow = Fill_Record("ASTSLID1", FORM_NAME)
        If rowASTSLID1 IsNot Nothing Then
            optGVL.Value = rowASTSLID1.Item("FORM_SLIDE_OPTION")
            txtTAGS.Text = rowASTSLID1.Item("FORM_TAGS")
            TextControl1.Load(rowASTSLID1.Item("FORM_DESCRIPTION"), TXTextControl.StringStreamType.HTMLFormat)
        End If

        txtTitle.MaxLength = rowASTTTIP1.Table.Columns("TOOLTIP_TITLE").MaxLength

        grdASTTTIP1.DataSource = FP.dst.Tables("ASTTTIP1")
        grdASTTTIP1.Text = "Tool Tips defined for " & FORM_NAME

        grdASTTTIP1.ActiveRow = grdASTTTIP1.Rows.GetRowWithListIndex _
            (FP.dst.Tables("ASTTTIP1").Rows.IndexOf(rowASTTTIP1))
        'next line gets an error
        'txtText.DataBindings.Add("BodyHTML", FP.dst.Tables("ASTTTIP1"), "TOOLTIP_TEXT")
        txtTitle.DataBindings.Add("Value", FP.dst.Tables("ASTTTIP1"), "TOOLTIP_TITLE")

        tab.SelectedTab = tab.Tabs("Tip")
        Setup_tab()

    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        For Each row As DataRow In FP.dst.Tables("ASTTTIP1").Select("", "", DataViewRowState.ModifiedCurrent)
            rowASTTTIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowASTTTIP1.Item("LAST_DATE") = DATETIME_STAMP
        Next

        Dim TTI As New UltraWinToolTip.UltraToolTipInfo
        TTI.ToolTipTitle = rowASTTTIP1.Item("TOOLTIP_TITLE") & ""
        TTI.ToolTipTextFormatted = rowASTTTIP1.Item("TOOLTIP_TEXT") & ""
        TTI.ToolTipTextStyle = ToolTipTextStyle.Formatted
        If TABLE_NAME <> "*" Then
            FP.tip.SetUltraToolTip(FP.CurrentControl, TTI)
        Else
            FP.tip.SetUltraToolTip(FP, TTI)
        End If

        Dim html As String = Get_HTML()
        If html <> "" Then
            Dim FILENAME As String = ASCMAIN1.Folders("Help") & "html\" & FORM_NAME & ".HTM"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                My.Computer.FileSystem.DeleteFile(FILENAME)
            End If
            Using SW As New System.IO.StreamWriter(FILENAME)
                SW.Write(html)
            End Using
        End If

        FP.Update_Record_TDA("ASTTTIP1", "FORM_NAME = '" & FORM_NAME & "'")
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        FP.dst.Tables("ASTTTIP1").RejectChanges()
        Me.Close()
    End Sub

    Private Sub grdASTTTIP1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTTTIP1.AfterRowActivate
        Setup_ASTTTIP1()
    End Sub

    Private Sub grdASTTTIP1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTTTIP1.BeforeRowUpdate
        If grdASTTTIP1.ActiveRow.IsAddRow Then
            grdASTTTIP1.ActiveRow.Cells("FORM_NAME").Value = FORM_NAME
            If grdASTTTIP1.ActiveRow.Cells("TABLE_NAME").Value & "" = "*" And _
               grdASTTTIP1.ActiveRow.Cells("COLUMN_NAME").Value & "" = "HEADER" Then
                grdASTTTIP1.ActiveRow.Cells("TOOLTIP_TITLE").Value = FP.MENU_ITEM_DESC
            End If
        End If
    End Sub

    Private Sub tab_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab.SelectedTabChanged
        Setup_tab()
    End Sub

    Sub Setup_ASTTTIP1()
        Dim TABLE_NAME As String = grdASTTTIP1.ActiveRow.Cells("TABLE_NAME").Text
        Dim COLUMN_NAME As String = grdASTTTIP1.ActiveRow.Cells("COLUMN_NAME").Text
        tab.Tabs("Tip").Text = TABLE_NAME & ":" & COLUMN_NAME
    End Sub

    Sub Setup_tab()
        If tab.SelectedTab Is Nothing Then Exit Sub

        lblTitle.Visible = (tab.SelectedTab.Key = "Tip") Or (tab.SelectedTab.Key = "Marketing Slide")
        txtTitle.Visible = (tab.SelectedTab.Key = "Tip") Or (tab.SelectedTab.Key = "Marketing Slide")

        If tab.SelectedTab.Key = "Page" Then
            Dim html As String = Get_HTML()
            html = Replace(html, "../css/ABS.css", "file:///" & ASCMAIN1.Folders("Help") & "css/ABS.css")
            html = Replace(html, "../images/ABS.png", "file:///" & ASCMAIN1.Folders("Help") & "images/ABS.png")

            WebBrowser1.DocumentText = html
            'WebBrowser1.Navigate("www.absolution.com")
        ElseIf tab.SelectedTab.Key = "Tip" Then

        Else

        End If
    End Sub

    Function Get_HTML() As String

        Dim html As String = ""
        Dim quo As String = Chr(34)

        Dim rowHEADER As DataRow = FP.dst.Tables("ASTTTIP1").Rows.Find(New String() {FORM_NAME, "*", "HEADER"})
        Dim TOOLTIP_TITLE As String = "{TITLE}"
        Dim TOOLTIP_HEADER As String = "{HEADER}"
        If rowHEADER IsNot Nothing Then
            TOOLTIP_HEADER = Strip_P(rowHEADER.Item("TOOLTIP_TEXT") & "")
            TOOLTIP_TITLE = rowHEADER.Item("TOOLTIP_TITLE") & ""
        End If

        Dim rowBODY As DataRow = FP.dst.Tables("ASTTTIP1").Rows.Find(New String() {FORM_NAME, "*", "BODY"})
        Dim TOOLTIP_BODY As String = "{BODY}"
        If rowBODY IsNot Nothing Then
            TOOLTIP_BODY = Strip_P(rowBODY.Item("TOOLTIP_TEXT") & "")
        End If

        TOOLTIP_BODY = Replace(TOOLTIP_BODY, "file:///C:/VS/ODG/Help/html/", "")
        html = "<!DOCTYPE HTML PUBLIC " & quo & "-//IETF//DTD HTML//EN" & quo & ">" & vbCrLf _
        & "<html>" & vbCrLf _
        & "<head>" & vbCrLf _
        & "<meta http-equiv=" & quo & "Content-Type" & quo & vbCrLf _
        & "content=" & quo & "text/html; charset=iso-8859-1" & quo & ">" & vbCrLf _
        & "<meta name=" & quo & "GENERATOR" & quo & " content=" & quo & "Microsoft FrontPage 5.0" & quo & ">" & vbCrLf _
        & "<title>ABSolution " & Replace(TOOLTIP_TITLE, "&", "&amp;") & "</title>" & vbCrLf _
        & "<style>@import url(../css/ABS.css);</style>" & vbCrLf _
        & "<link disabled rel=" & quo & "stylesheet" & quo & " href=" & quo & "../css/ABS.css" & quo & ">" & vbCrLf _
        & "</head>" & vbCrLf _
        & "<body>" & vbCrLf _
        & "<h4>" _
        & Replace(TOOLTIP_TITLE, "&", "&amp;") & "<a href=" & quo & "http://www.absolution.com" & quo & ">" & vbCrLf _
        & "<IMG SRC=" & quo & "../images/ABS.png" & quo & vbCrLf _
        & " style='border:none; float:right;position:relative;margin-top:-45px;margin-right:-3px; height: 40px;'/></a></h4>" & vbCrLf _
        & "<h3>" & vbCrLf & TOOLTIP_HEADER & vbCrLf _
        & "<BR>Please click <A href=" & quo & "file:///C:/VS/SEA/Help/STD/001.htm" & quo & ">here</A> for general help with the operation of any File Maintenance screen." _
        & "</h3><br>" & vbCrLf _
        & "" _
        & vbCrLf & TOOLTIP_BODY & "<BR><BR>"

        Dim TABLE_NAME As String = ""
        For Each row2 As DataRow In FP.dst.Tables("ASTTTIP1").Select _
            ("TABLE_NAME <> '*' and COLUMN_NAME <> '*'", "TABLE_NAME,COLUMN_NAME")
            If row2.Item("TABLE_NAME") <> TABLE_NAME Then
                If TABLE_NAME <> "" Then html &= "</TABLE>"
                TABLE_NAME = row2.Item("TABLE_NAME")
                html &= "<TABLE>"
                html &= "<TR class=" & quo & "table_hdr" & quo & ">" & "<TD  width='25%' >" & "Field" & "</TD>" & "<TD>" & "Description (" & TABLE_NAME & ")" & "</TD>" & "</TR>"
            End If
            html &= "<TR>" & "<TD  width='25%'>" & row2.Item("TOOLTIP_TITLE") & "</TD>" & "<TD>" & row2.Item("TOOLTIP_TEXT") & "</TD>" & "</TR>"
        Next
        If TABLE_NAME <> "" Then html &= "</TABLE>"

        html &= "" _
        & "<br><h3>" & vbCrLf _
        & "<a href=" & quo & "mailto:support@absolution.com?subject=ABSolution%20Documentation%20Feedback" & quo & ">" _
        & "Send comments on this topic.</a>" & vbCrLf _
        & "<br><a>&copy; Applied Business Systems, Inc. All rights reserved.</a></h3>"

        html &= "" _
        & "</body>" & vbCrLf _
        & "</html>" & vbCrLf

        Return html

    End Function

    Private Sub grdASTTTIP1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTTTIP1.DoubleClickRow
        tab.SelectedTab = tab.Tabs("Tip")
    End Sub

    Function Strip_P(ByVal T As String) As String
        If T.StartsWith("<P>") Then
            T = Mid(T, 4)
            T = Mid(T, 1, Len(T) - 4)
        End If
        Return T
    End Function
     

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim rowASTSLID1 As DataRow = Fill_Record("ASTSLID1", FORM_NAME)
        If rowASTSLID1 Is Nothing Then
            rowASTSLID1 = dst.Tables("ASTSLID1").NewRow
            rowASTSLID1.Item("FORM_NAME") = FORM_NAME
            dst.Tables("ASTSLID1").Rows.Add(rowASTSLID1)
        End If
        rowASTSLID1.Item("FORM_TITLE") = txtTitle.Text
        rowASTSLID1.Item("FORM_SLIDE_OPTION") = optGVL.Value
        'Dim st As TXTextControl.StreamType = TXTextControl.StreamType.HTMLFormat
        Dim txt As String = ""
        TextControl1.Save(txt, TXTextControl.StringStreamType.HTMLFormat)
        rowASTSLID1.Item("FORM_DESCRIPTION") = txt
        rowASTSLID1.Item("FORM_TAGS") = txtTAGS.Text
        Update_Record_TDA("ASTSLID1")

        MsgBox("Slide Saved")
    End Sub

    Private Sub btnJSON_Click(sender As Object, e As EventArgs) Handles btnJSON.Click

        Dim txt As String = ""
        TextControl1.Save(txt, TXTextControl.StringStreamType.HTMLFormat)

        Dim quo As String = Chr(34)

        Dim x As Integer = InStr(txt, "</body>")
        txt = Mid(txt, 1, x - 1)
        x = InStr(txt, "<body")
        txt = Mid(txt, x + 1)
        x = InStr(txt, ">")
        txt = Mid(txt, x + 1)
        txt = Replace(txt, quo, "'")

        Dim tagstxt As String = txtTAGS.Text
        Do While tagstxt.EndsWith(vbCrLf)
            tagstxt = Mid(tagstxt, 1, tagstxt.Length - 2)
        Loop
        Do While tagstxt.Contains(vbCrLf & vbCrLf)
            tagstxt = Replace(tagstxt, vbCrLf & vbCrLf, vbCrLf)
        Loop

        Dim tagsarray() As String = Split(tagstxt, vbCrLf)
        Dim tagslist As String = Join(tagsarray, quo & "," & quo)
        Dim tags As String = ""
        If tagslist <> "" Then
            tags = quo & tagslist & quo
        End If
        Dim txt2 As String = TextControl1.Text

        Dim image As String = ""
        Dim image_filename As String = FORM_NAME & ".png"
        If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("SharedRoot") & "\Slides\" & image_filename) Then
            image = quo & image_filename & quo
        End If

        Dim video As String = ""
        Dim video_filename As String = FORM_NAME & ".mp4"
        If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("SharedRoot") & "\Slides\" & video_filename) Then
            video = quo & video_filename & quo
        End If

        Dim gallery As String = ""
        For Each FILENAME As String In My.Computer.FileSystem.GetFiles _
            (ASCMAIN1.Folders("SharedRoot") & "\Slides\", _
             FileIO.SearchOption.SearchTopLevelOnly, _
             FORM_NAME & "*.png")

            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)

            If fi.Name = FORM_NAME & ".png" Then
                ' skip this one
            Else

                gallery &= "," & quo & fi.Name & quo
            End If
        Next
        If gallery <> "" Then gallery = Mid(gallery, 2)

        Dim JSON As String = "{" & vbCrLf _
                             & "key:" & quo & FORM_NAME & quo & "," & vbCrLf _
                             & "title:" & quo & txtTitle.Text & quo & "," & vbCrLf _
                             & "image:" & quo & FORM_NAME & ".png" & quo & "," & vbCrLf _
                             & "gallery:[" & gallery & "]" & "," & vbCrLf _
                             & "video:" & video & "," & vbCrLf _
                             & "command:" & quo & optGVL.Value & quo & "," & vbCrLf _
                             & "description:" & quo & txt & quo & "," & vbCrLf _
                             & "tags:[" & tags & "]" & vbCrLf _
                             & "}"
        'fragrance = [{
        'key:"ARFCINQ1",
        'title:"Customer Inquiry",
        'image:"ARFCINQ1.PNG",
        'gallery:[],
        'video:"",
        'command:"G/V/L",
        'description:"Lorem Ipsum",
        'tags:["EDI","AR","Sales Admin"]
        '}]

        Using frm As New ASFTEXT1
            frm.t = JSON
            frm.Text = "json object rendered for " & FORM_NAME
            frm.ShowDialog()
        End Using

    End Sub
End Class