Public Class WHCRF000
    Inherits ASCBASE0

    Public Event RespondToScan(THREAD_NO As Integer, ByVal RESPONSE As String)

    Public AppState As String
    Public AppStates As New Dictionary(Of String, String)

    Public RESPONSE As String
    Public PREAMBLE_remembered As String
    Public PREAMBLE_remembered_for As String
    Public LAST_CLR As String
    Public RESPONSE_anticipated_next As String
    Public AppState_initial As String

    Sub New(g As GunEnvironment)
        MyBase.New(g)

    End Sub

    Public Overridable Function Hello() As String
        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overridable Sub GetResponseToScan(SCANTEXT As String)
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

    End Sub

    Sub CreateResponse(new_AppState As String, CLR As String, PREAMBLE As String, Optional remember_PREAMBLE As Boolean = False)
        If new_AppState <> "" Then AppState = new_AppState

        ' CLR B=BLUE, R=RED, G=GREEN

        'Dim RESPONSE As String = ""
        'If PRE <> "" Then RESPONSE = CLR & Format(PRE.Length, "000") & PRE
        'RESPONSE &= "X" & AppStates(AppState)

        RESPONSE = PREAMBLE
        LAST_CLR = CLR

        If AppState = PREAMBLE_remembered_for Then
            If PREAMBLE_remembered <> "" Then RESPONSE &= vbCrLf & PREAMBLE_remembered
        Else
            PREAMBLE_remembered = ""
            PREAMBLE_remembered_for = ""
            If new_AppState <> "" And remember_PREAMBLE Then
                PREAMBLE_remembered = PREAMBLE
                PREAMBLE_remembered_for = new_AppState
            End If
        End If

        If New String() {"EXIT", "NEXT_INST"}.Contains(AppState) Then
        Else
            If RESPONSE <> "" Then RESPONSE &= vbCrLf
            RESPONSE &= GetStatus()
        End If

        If RESPONSE <> "" Then
            RESPONSE &= vbCrLf & vbCrLf
        End If
        RESPONSE &= AppStates(AppState)

        If AppState = AppState_initial Then
            RESPONSE_anticipated_next = ""
        Else
            RESPONSE_anticipated_next = Get_Anticipated_Next_Response()
        End If


        RaiseEvent RespondToScan(G.THREAD_NO, RESPONSE)
    End Sub

    Overridable Function GetStatus() As String
        Dim STATUS As String = ""
        Return STATUS
    End Function

    Public Sub Dispose()
        If ASCMAIN1 IsNot Nothing Then
            Me.ASCMAIN1.MultiTask_Release()
        End If
        Me.clsASCBASE1.Dispose()
        Me.ASCDATA1.Dispose()
    End Sub

    Overridable Function Get_Anticipated_Next_Response() As String
        Return ""
    End Function
End Class