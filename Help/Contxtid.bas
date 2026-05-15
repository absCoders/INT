Attribute VB_Name = "ContextIDs"
Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <Adi>
'=====================================================================
Global Const Hlp_ABSolution = 10    'Main Help Window
Global Const Hlp_LogOn = 240    'Main Help Window
Global Const Hlp_AR = 450    'Main Help Window
Global Const Hlp_SOFOREL1 = 460    'Main Help Window
Global Const Hlp_AP = 470    'Main Help Window
Global Const Hlp_BM = 480    'Main Help Window
Global Const Hlp_ED = 490    'Main Help Window
Global Const Hlp_GL = 500    'Main Help Window
Global Const Hlp_EDI_x = 510    'Main Help Window
Global Const Hlp_EDI_x1 = 520    'Main Help Window
Global Const Hlp_ARE = 530    'Main Help Window
Global Const Hlp_ARM = 540    'Main Help Window
Global Const Hlp_Copyright = 550    'Main Help Window
Global Const Hlp_ARFCUST1 = 790    'Main Help Window
Global Const Hlp_General_Notes = 820    'Main Help Window
Global Const Hlp_ARFCUST2 = 830    'Main Help Window
Global Const Hlp_SOFTCLS1 = 860    'Main Help Window
Global Const Hlp_SOFBILL1 = 870    'Main Help Window
Global Const Hlp_ARFTERM1 = 880    'Main Help Window
Global Const Hlp_ARFPARM1 = 890    'Main Help Window
Global Const Hlp_SO = 910    'Main Help Window
Global Const Hlp_SOFORDR1 = 920    'Main Help Window
Global Const Hlp_SOFPCLS1 = 930    'Main Help Window
Global Const Hlp_IC = 940    'Main Help Window
Global Const Hlp_ICE = 950    'Main Help Window
Global Const Hlp_ICFITEM1 = 960    'Main Help Window
Global Const Hlp_ICM = 970    'Main Help Window
Global Const Hlp_ICFPROD1 = 980    'Main Help Window
Global Const Hlp_SOFCPRE1 = 990    'Main Help Window
Global Const Hlp_ICFBRAN1 = 1000    'Main Help Window
Global Const Hlp_ICFCLAS1 = 1010    'Main Help Window
Global Const Hlp_ICFTYPE1 = 1020    'Main Help Window
Global Const Hlp_ICFADJR1 = 1030    'Main Help Window
Global Const Hlp_ICFWHSE1 = 1040    'Main Help Window
Global Const Hlp_ICFCATG1 = 1050    'Main Help Window
Global Const Hlp_ICFCURR1 = 1060    'Main Help Window
Global Const Hlp_ICFFRTC1 = 1070    'Main Help Window
Global Const Hlp_SOFSREP1 = 1080    'Main Help Window
Global Const Hlp_ARFCOLL1 = 1090    'Main Help Window
Global Const Hlp_ARFCLAS1 = 1100    'Main Help Window
Global Const Hlp_ARFSTATE = 1110    'Main Help Window
Global Const Hlp_SOFSVIA1 = 1120    'Main Help Window
Global Const Hlp_ARFSTAX1 = 1130    'Main Help Window
Global Const Hlp_SOFSREG1 = 1140    'Main Help Window
Global Const Hlp_ARFREAS1 = 1150    'Main Help Window
Global Const Hlp_ARFPOST1 = 1160    'Main Help Window
Global Const Hlp_ARFDMSG1 = 1170    'Main Help Window
Global Const Hlp_SOE = 1180    'Main Help Window
Global Const Hlp_SOM = 1190    'Main Help Window
Global Const Hlp_SOFREAS1 = 1200    'Main Help Window
Global Const Hlp_SOFFORM1 = 1210    'Main Help Window
Global Const Hlp_SOFPROM1 = 1220    'Main Help Window
Global Const Hlp_ICFPARM1 = 1230    'Main Help Window
Global Const Hlp_SOFPARM1 = 1240    'Main Help Window
Global Const Hlp_ARFCSOH1 = 1250    'Main Help Window
Global Const Hlp_Systems_Functions = 1270    'Main Help Window
Global Const Hlp_File_Maintenance = 1280    'Main Help Window
Global Const Hlp_SOTORDR2 = 1290    'Main Help Window
Global Const Hlp_Sales_Rep = 1300    'Main Help Window
Global Const Hlp_Reports = 1310    'Main Help Window
Global Const Hlp_SOFCPRR1 = 1320    'Main Help Window
Global Const Hlp_General_Notes1 = 1330    'Main Help Window
Global Const Hlp_Operational_Calendar = 1340    'Main Help Window
Global Const Hlp_Item_Status = 1350    'Main Help Window
Global Const Hlp_SOTORDR1 = 1360    'Main Help Window
Global Const Hlp_SOFOPEN1 = 1370    'Main Help Window
Global Const Hlp_SOFTREO1 = 1380    'Main Help Window
Global Const Hlp_SOFZERO1 = 1390    'Main Help Window
Global Const Hlp_SOFTPRE1 = 1400    'Main Help Window
Global Const Hlp_Territory_Allotment = 1410    'Main Help Window
Global Const Hlp_SOFSDIV1 = 1420    'Main Help Window
Global Const Hlp_Customer_x = 1430    'Main Help Window
Global Const Hlp_Bill_of = 1440    'Main Help Window
Global Const Hlp_General_Notes2 = 1450    'Main Help Window
Global Const Hlp_SOFSLSB1 = 1460    'Main Help Window
Global Const Hlp_SOFSLST1 = 1470    'Main Help Window
Global Const Hlp_SOFMKTC1 = 1480    'Main Help Window
Global Const Hlp_IC_x = 1490    'Main Help Window
Global Const Hlp_PostxPeriod_End = 1500    'Main Help Window
Global Const Hlp_Period_End = 1510    'Main Help Window
Global Const Hlp_PrexPeriod_End = 1520    'Main Help Window
Global Const Hlp_Daily_Transaction = 1530    'Main Help Window
Global Const Hlp_SOFMTDS1 = 1540    'Main Help Window
Global Const Hlp_SOFCGSJ1 = 1550    'Main Help Window
Global Const Hlp_ICFIVAL1 = 1560    'Main Help Window
Global Const Hlp_ICFISTA1 = 1570    'Main Help Window
Global Const Hlp_ARFMEMJ1 = 1580    'Main Help Window
Global Const Hlp_ICFTRNR1 = 1590    'Main Help Window
Global Const Hlp_MR = 1600    'Main Help Window
Global Const Hlp_PO = 1610    'Main Help Window
Global Const Hlp_POE = 1620    'Main Help Window
Global Const Hlp_POM = 1630    'Main Help Window
Global Const Hlp_POR = 1640    'Main Help Window
Global Const Hlp_MRE = 1650    'Main Help Window
Global Const Hlp_MRM = 1660    'Main Help Window
Global Const Hlp_MRR = 1670    'Main Help Window
Global Const Hlp_MRFLOST1 = 1680    'Main Help Window
Global Const Hlp_MRFMSVF1 = 1690    'Main Help Window
Global Const Hlp_MRFSAFE1 = 1700    'Main Help Window
Global Const Hlp_POFOPEN1 = 1710    'Main Help Window
Global Const Hlp_Item_Cost = 1720    'Main Help Window
Global Const Hlp_BMFLIST1 = 1730    'Main Help Window
Global Const Hlp_BMFWUSR1 = 1740    'Main Help Window
Global Const Hlp_Item_Master = 1750    'Main Help Window
Global Const Hlp_MRFACTM1 = 1760    'Main Help Window
Global Const Hlp_MRFCRDM1 = 1770    'Main Help Window
Global Const Hlp_Cost_x = 1780    'Main Help Window
Global Const Hlp_Reports1 = 1790    'Main Help Window
Global Const Hlp_Inventory_Reserves = 1800    'Main Help Window
Global Const GLOS_Collateral_Material = 1810
Global Const GLOS_Allocated = 1820
Global Const GLOS_Reserved = 1830
Global Const GLOS_Credit_Limit_Expiration_Date = 1840
Global Const GLOS_Credit_Available = 1850
Global Const GLOS_Sales_Order_Number = 1860
Global Const GLOS_Qty_Available = 1870
Global Const GLOS_Ship_From_Warehouse = 1880
Global Const GLOS_SoldxTo_Customer = 1890
Global Const GLOS_BillxTo_Customer = 1900
Global Const GLOS_Unconditionally_Delete = 1910
Global Const GLOS_RexOrder_Memo = 1920
Global Const GLOS_Suppress_if_Short = 1930
Global Const GLOS_Credit_Limit_is_Zerox_or_has_Expired = 1940
Global Const GLOS_Pick_Ticket = 1950
Global Const GLOS_Order_Priority = 1960
Global Const GLOS_Past_Due = 1970
Global Const GLOS_Sales_Order_Hold = 1980
Global Const GLOS_Credit_Hold = 1990
Global Const GLOS_ShipxFrom_Warehouse = 2000
Global Const GLOS_Customer_Priority = 2010
Global Const GLOSSARY = 64000
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

    Type HELPWININFO
      wStructSize As Long
      X As Long
      Y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow As Long ' hWnd to tell WINHELP the helpfile owner

Dim MainWindowInfo As HELPWININFO
Public Sub SetAppHelp(ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path), 1) = "\" Then
        'App.HelpFile = App.Path + "ICI.CHM"
    Else
        'App.HelpFile = App.Path + "\ICI.CHM"
    End If
    MainWindowInfo.wStructSize = 26
    MainWindowInfo.X = 256
    MainWindowInfo.Y = 256
    MainWindowInfo.dX = 512
    MainWindowInfo.dY = 512
    MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub
Public Sub QuitHelp()
    Dim Result As Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Public Sub ShowHelpTopic(ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(ContextID))

End Sub
Public Sub ShowHelpTopic2(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd02", HELP_CONTEXT, CLng(ContextID))

End Sub
Public Sub ShowHelpTopic3(ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd03", HELP_CONTEXT, CLng(ContextID))

End Sub
Public Sub ShowGlossary()
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CLng(64000))

End Sub
Public Sub ShowPopupHelp(ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXTPOPUP, CLng(ContextID))

End Sub
Public Sub DoHelpMacro(ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByStr(m_hWndMainWindow, App.HelpFile, HELP_COMMAND, ByVal (MacroString))

End Sub
Public Sub ShowHelpContents()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub
Public Sub ShowContentsTab()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_FINDER, CLng(0))

End Sub
Public Sub ShowHelpOnHelp()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_HELPONHELP, CLng(0))

End Sub

Public Sub SearchHelp()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result As Variant

    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal "")

End Sub

Public Sub SearchHelpKeyWord(Argument As String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result As Variant

    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Public Sub HelpWindowSize(X As Integer, Y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result As Variant
    MainWindowInfo.X = X
    MainWindowInfo.Y = Y
    MainWindowInfo.dX = wx
    MainWindowInfo.dY = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub
