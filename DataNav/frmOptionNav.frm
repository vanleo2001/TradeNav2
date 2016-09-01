VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmOptionNav 
   Caption         =   "Option Navigator"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "frmOptionNav.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11595
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   1020
      Top             =   6540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   4
      Tools           =   "frmOptionNav.frx":0A02
      ToolBars        =   "frmOptionNav.frx":3CC0
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      ExtentX         =   13467
      ExtentY         =   9340
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmOptionNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    strStartingAddress As String
End Type
Private m As mPrivate

Public Sub ShowMe()

    ShowForm Me, , frmMain

End Sub

Private Sub Form_Load()

    Dim strText As String
    Dim aStrings As New cGdArray

    ' form placement
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If

    ' get starting address
    aStrings.FromFile App.Path & "\Provided\OptionNav.adr"
    m.strStartingAddress = aStrings(0)
    If m.strStartingAddress = "" Then m.strStartingAddress = "http://options.genesisft.com/default.aspx"
        
    ' Get the data service ID, encrypt it, convert it to hex and add it to the address
    m.strStartingAddress = m.strStartingAddress & "?s=" & EncryptToHex(RI_GetDataServiceID)
    
    ' Add the password
    m.strStartingAddress = m.strStartingAddress & "&p=" & EncryptToHex(RI_GetUserPassword)
    
    ' Add a date string
    m.strStartingAddress = m.strStartingAddress & "&d=" & EncryptToHex(CStr(ConvertTimeZone(Now, "", "NY")))
    
    ' Load the initial page while navigating to the options site
'    webBrowser.Navigate2 "about:blank"
'    Dim dtStart As Date
'    dtStart = Now
'    Do While DateDiff("s", dtStart, Now) < 5
'        If webBrowser.Document.ReadyState = "complete" Then Exit Do
'    Loop
'    webBrowser.Document.body.innerHtml = "<br><br><br><br><p align=center><font size=5 color='#009900'><bold>Loading...</bold></font></p>"

    ' Load the options validation page
    webBrowser.Navigate m.strStartingAddress

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If LimitFormSize(Me, 2000, 2000) Then Exit Sub
    
    With webBrowser
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, Me.ScaleHeight - .Left * 2
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile

End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    On Error Resume Next
    With Tool
        Select Case UCase(.ID)
        Case "ID_BACK"
            webBrowser.GoBack
        Case "ID_FORWARD"
            webBrowser.GoForward
        Case "ID_REFRESH"
            webBrowser.Refresh
        Case "ID_HOME"
            webBrowser.Navigate m.strStartingAddress
        End Select
    End With

End Sub

Private Sub webBrowser_DownloadBegin()
    On Error Resume Next
    Me.Caption = "Working..."
    Me.MousePointer = vbHourglass
End Sub

Private Sub webBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = "Option Navigator"
    Me.MousePointer = vbDefault
End Sub

Public Sub HtmlEvents()
    On Error Resume Next
    'MsgBox "HtmlEvent"
    Dim nIndex As Integer
    nIndex = webBrowser.Document.All("lstOpts").SelectedIndex
    If nIndex >= 0 Then
        MsgBox webBrowser.Document.All("lstOpts").Item(nIndex).Text
        'Unload Me
    End If
End Sub

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    ' Below is an example of how HTML events in the web browser control can be
    ' handeled in this form.  This is currently not being used. Requires adding
    ' the cWebBrowserFwd class to the project.
    
    ' Set up the HTML event forwarding for the OptSel.aspx page
'    If InStr(1, URL, "OptSel.aspx") > 0 Then
'        Dim clsWebBrowserFwd As cWebBrowserFwd
'        Set clsWebBrowserFwd = New cWebBrowserFwd
'        clsWebBrowserFwd.Set_Destination Me, "HtmlEvents"
'        webBrowser.Document.All("btnOK").onclick = clsWebBrowserFwd
'
'        Set clsWebBrowserFwd = New cWebBrowserFwd
'        clsWebBrowserFwd.Set_Destination Me, "btnCancel_onclick"
'        webBrowser.Document.All("btnCancel").onclick = clsWebBrowserFwd
'    End If
End Sub
