VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmLoginTradier 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   4048
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
Attribute VB_Name = "frmLoginTradier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLoginTradier.frm
'' Description: Class to manange communications with Tradier servers
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/14/2015   DAJ         Created
'' 09/15/2015   DAJ         Set WebBrowser.Silent to true to suppress Java Script errors
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user OK the dialog?
    Broker As cBrokerTradier            ' Broker object
    strState As String                  ' State that should be coming back on web calls
    strCode As String                   ' Code back on the URL
End Type
Private m As mPrivate

Public Property Get Code() As String
    Code = m.strCode
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form
'' Inputs:      Broker object, URL, State
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Broker As cBrokerTradier, ByVal strUrl As String, ByVal strState As String) As Boolean
On Error GoTo ErrSection:

    Set m.Broker = Broker
    Caption = "Login to " & Broker.Broker.BrokerName
    m.strState = strState
    
    WebBrowser1.Silent = True
    WebBrowser1.Navigate2 strUrl
    
    ShowForm Me, eForm_Modal, frmMain
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLoginTradier.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Do some initialization when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = Picture16("kBlank")
    
    g.Styler.StyleForm Me
    
    Width = 15315
    Height = 14445
    CenterTheForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginTradier.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WebBrowser1_NavigateComplete2
'' Description: Notification that the navigation to a site is complete
'' Inputs:      Display Object, URL
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
On Error GoTo ErrSection:

    If CheckURL = True Then
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLoginTradier.WebBrowser1_NavigateComplete2"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckURL
'' Description: Check the URL for certain strings
'' Inputs:      None
'' Returns:     True if need to close form, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckURL() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrUrl As cGdArray             ' URL broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    m.strCode = ""
    
    Set astrUrl = New cGdArray
    astrUrl.SplitFields WebBrowser1.LocationURL, "?&"
    frmTest2.AddList astrUrl.JoinFields(vbTab)
    
    For lIndex = 0 To astrUrl.Size - 1
        If InStr(astrUrl(lIndex), "code=") <> 0 Then
            m.Broker.Broker.DumpDebug "Login: '" & astrUrl(lIndex) & "'"
            
            m.bOK = True
            m.strCode = Mid(astrUrl(lIndex), 6)
            bReturn = True
            Exit For
        ElseIf InStr(astrUrl(lIndex), "error=") <> 0 Then
            m.Broker.Broker.DumpDebug "Login: '" & astrUrl(lIndex) & "'"
            
            m.bOK = False
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    CheckURL = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmLoginTradier.CheckURL"
    
End Function

