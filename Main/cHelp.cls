VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_DISPLAY_TEXT_POPUP = &HE   ' text in a pop-up window
Private Const HH_HELP_CONTEXT = &HF         ' mapped context ID in dwData
Private Const HH_TP_HELP_CONTEXTMENU = &H10 ' text pop-up help
Private Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help
Private Const HH_CLOSE_ALL = &H12

'HTML Help API Declaration
Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, ByVal pszFile As String, _
        ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Declare Function HTMLHelpT Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, ByVal pszFile As String, _
        ByVal uCommand As Long, ByVal dwData As String) As Long

Private Type mPrivate
    hWnd As Long
    strHelpPath As String
    bUsed As Boolean
End Type
Private m As mPrivate

Public Sub Init(ByVal hWnd As Long, ByVal strHelpPath As String)
On Error GoTo ErrSection:

    ' TLB: If we just use "0" for the window handle, then the user can use "Alt-Tab"
    ' to switch between Help/TradeNav (and it will no longer stay on top of TradeNav)
    ' -- passing "0" makes the help system act more like all other help programs.
    'm.hWnd = hWnd
    
    m.hWnd = GetDesktopWindow
    m.strHelpPath = strHelpPath

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cHelp.Init", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Public Sub ShowF1Help(frm As Form, Optional ByVal strTopic$ = "")
On Error GoTo ErrSection:

    Dim strPathName As String
    Dim strChar As String
    
DoNewHelp
Exit Sub
    
    If Not HelpExists Then Exit Sub
    
    ' Set topic if not sent
    If Len(strTopic) = 0 And Not frm Is Nothing Then
        Select Case frm.Name
            Case "frmSpecial"           'example of special case (ex. multiple caption names)
                strTopic = "Special"
            Case Else                   'standard form summary
                strTopic = frm.Name
        End Select
    End If

    ' Set to default
    If Len(strTopic) = 0 Then strTopic = "TradeNavigator2"
    
    ' Show help
'***If 0 Then   'code for pointing directly into F1 chm files (issues w/merging TOC,index,search)
    strPathName = AddSlash(m.strHelpPath) & "F1WinSummaries"
    strChar = UCase(Mid(strTopic, 4, 1))    'topic example-frmChart
    Select Case strChar
    Case "A", "B"
        strPathName = strPathName & "AB.chm"
    Case "C"
        strPathName = strPathName & "C.chm"
    Case "D"
        strPathName = strPathName & "D.chm"
    Case "E"
        strPathName = strPathName & "E.chm"
    Case "F", "G", "H"
        strPathName = strPathName & "FH.chm"
    Case "I", "J", "K", "L", "M"
        strPathName = strPathName & "IM.chm"
    Case "N", "O"
        strPathName = strPathName & "NO.chm"
    Case "P"
        strPathName = strPathName & "P.chm"
    Case "Q", "R"
        strPathName = strPathName & "QR.chm"
    Case "S"
        strPathName = strPathName & "S.chm"
    Case "T"
        strPathName = strPathName & "T.chm"
    Case "U", "V", "W", "X", "Y", "Z"
        strPathName = strPathName & "UZ.chm"
    
    Case Else
        'strPathName = AddSlash(m.strHelpPath) & "F1WindowSummaries.chm"
        strPathName = strPathName & "AB.chm"
    End Select
    
    ShowHelpByFileName strPathName, strTopic & ".htm"
'***End If

    '***strPathName = AddSlash(m.strHelpPath) & "TradeNavHelp.chm"  'point to main help
    '***ShowHelpByFileName strPathName, strTopic & "x.htm"  'point to tiny F1 topic file

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cHelp.ShowF1Help", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub ShowHelpByFileName(ByVal strHelpFile As String, ByVal strTopic As String)
On Error Resume Next

DoNewHelp
Exit Sub

    m.bUsed = True ' Set the initialization flag
    
    ' TLB 8/14/2007: For some unknown reason with Vista, not all the
    ' TradeNav help will display with the normal HTMLHelp call -- so
    ' we can do a workaround: just call the HH program directly
    If IsAtLeastVista And FileExist(WindowsPath & "HH.exe") Then
        RunProcess WindowsPath & "HH.exe", strHelpFile & "::" & strTopic
    Else
        HTMLHelpT m.hWnd, strHelpFile, HH_DISPLAY_TOPIC, strTopic
    End If

End Sub

Public Sub ShowHelpDefault()
On Error Resume Next
    
DoNewHelp
Exit Sub
    
    If Not HelpExists Then Exit Sub
    
    m.bUsed = True ' Set the initialization flag
    
    ' TLB 8/14/2007: For some unknown reason with Vista, not all the
    ' TradeNav help will display with the normal HTMLHelp call -- so
    ' we can do a workaround: just call the HH program directly
    If IsAtLeastVista And FileExist(WindowsPath & "HH.exe") Then
        RunProcess WindowsPath & "HH.exe", AddSlash(m.strHelpPath) & "TradeNavHelp.chm"
    Else
        HTMLHelp m.hWnd, AddSlash(m.strHelpPath) & "TradeNavHelp.chm", HH_DISPLAY_TOPIC, 0
    End If

End Sub

Private Sub ShowHelpByID(ByVal strHelpFile As String, ByVal lMapNumber As Long)
On Error Resume Next

DoNewHelp
Exit Sub
    
    m.bUsed = True ' Set the initialization flag
    HTMLHelp m.hWnd, strHelpFile, HH_HELP_CONTEXT, lMapNumber

End Sub

Private Sub Class_Initialize()

    m.bUsed = False

End Sub

Private Sub Class_Terminate()
On Error Resume Next
        
    ' MSDN documentation specifies that hWndCaller must be NULL when doing CLOSE_ALL
    ' (don't know if it really matters or not, but better do what they say!)
    If m.bUsed Then HTMLHelp 0, "", HH_CLOSE_ALL, 0

End Sub

Private Function HelpExists() As Boolean
    
    #If TRADENAV_EXE Then
        ' don't show our help for Extreme version
        If ExtremeCharts >= 1 Then
            Beep
            Exit Function
        End If
    #End If
    
    If FileExist(App.Path & "\Help\*.chm") Then
        HelpExists = True
    Else
        #If TRADENAV_EXE Then
            If InfBox("The Trade Navigator help files do not exist.||Would you like to download them now?", _
                    "?", "+Download|-Not now", "Trade Navigator Help") = "D" Then
                If Not ProcessIsBusy Then
                    frmMain.tmrMain.Tag = "DownloadHelp"
                End If
            End If
        #Else
            InfBox "You need to download the special file 'HELP'| to use the Trade Navigator help system.", _
                "i", , "Help System"
        #End If
    End If

End Function

' TLB 5/13/2015: since our old context-sensitive help system is so outdated (and still contains LW stuff),
' we will now instead just point people to our online help manuals ...
Private Sub DoNewHelp()

    On Error Resume Next
    Dim s$
    s = "http://www.tradenavigator.com/manuals.php"
    #If TRADENAV_EXE Then
        s = FixURL(s)
    #End If
    RunProcess InternetBrowser, Chr(34) & s & Chr(34)

End Sub
