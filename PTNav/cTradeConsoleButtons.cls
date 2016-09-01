VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeConsoleButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeConsoleButtons.cls
'' Description: Information about Trade Console toolbar buttons
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eGDTcButtons
    eGDTcButtons_Dashboard = 0
    eGDTcButtons_OpenOrders
    eGDTcButtons_Positions
    eGDTcButtons_Accounts
    eGDTcButtons_AutoTrading
    eGDTcButtons_TradeSenseOrders
    eGDTcButtons_ActivityLog
    eGDTcButtons_TodaysFills
    eGDTcButtons_BuySell
    eGDTcButtons_Settings
    eGDTcButtons_Tracking
    eGDTcButtons_Reports
    eGDTcButtons_Journals
    eGDTcButtons_NumButtons
End Enum

Private Type TcBtnSettings
    strName As String                   ' Name of the button
    bShow As Boolean                    ' Show the button?
    bShowInLists As Boolean             ' Show the button in lists?
End Type

Private Type mPrivate
    astrButtons As cGdArray             ' Collection of buttons
End Type
Private m As mPrivate

Private Function Button(ByVal nButton As eGDTcButtons) As Long
    Button = nButton
End Function

Public Property Get Size()
    Size = m.astrButtons.Size
End Property

Public Property Get Name(ByVal nButton As eGDTcButtons) As String
On Error GoTo ErrSection:

    Dim TcBtn As TcBtnSettings          ' Settings for the given button
    Dim strReturn As String             ' Return value for the function

    strReturn = ""
    If (nButton >= 0) And (nButton < Button(eGDTcButtons_NumButtons)) Then
        TcBtn = TcBtnFromString(m.astrButtons(nButton))
        strReturn = TcBtn.strName
    End If
    
    Name = strReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Name.Get"
    
End Property

Public Property Get Show(ByVal nButton As eGDTcButtons) As Boolean
On Error GoTo ErrSection:

    Dim TcBtn As TcBtnSettings          ' Settings for the given button
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (nButton >= 0) And (nButton < Button(eGDTcButtons_NumButtons)) Then
        TcBtn = TcBtnFromString(m.astrButtons(nButton))
        
        If (nButton = eGDTcButtons_AutoTrading) Or (nButton = eGDTcButtons_TradeSenseOrders) Then
            bReturn = TcBtn.bShow And (ExtremeCharts <> 1)
        Else
            bReturn = TcBtn.bShow
        End If
    End If
    
    Show = bReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Show.Get"
    
End Property
Public Property Let Show(ByVal nButton As eGDTcButtons, ByVal bShow As Boolean)
On Error GoTo ErrSection:

    Dim TcBtn As TcBtnSettings          ' Settings for the given button
    
    If (nButton >= 0) And (nButton < Button(eGDTcButtons_NumButtons)) Then
        TcBtn = TcBtnFromString(m.astrButtons(nButton))
        TcBtn.bShow = bShow
        m.astrButtons(nButton) = TcBtnToString(TcBtn)
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Show.Let"
    
End Property

Public Property Get ShowInLists(ByVal nButton As eGDTcButtons) As Boolean
On Error GoTo ErrSection:

    Dim TcBtn As TcBtnSettings          ' Settings for the given button
    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If (nButton >= 0) And (nButton < Button(eGDTcButtons_NumButtons)) Then
        TcBtn = TcBtnFromString(m.astrButtons(nButton))
        
        If (nButton = eGDTcButtons_AutoTrading) Or (nButton = eGDTcButtons_TradeSenseOrders) Then
            bReturn = TcBtn.bShowInLists And (ExtremeCharts <> 1)
        Else
            bReturn = TcBtn.bShowInLists
        End If
    End If
    
    ShowInLists = bReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeConsoleButtons.ShowInLists.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the information from the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim strSettings As String           ' Settings from the INI file
    Dim astrSettings As cGdArray        ' Settings broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim strIndex As String              ' Index into the array
    Dim strValue As String              ' Value for the array element
    
    strSettings = GetIniFileProperty("ToolbarButtons", "", "TTSummary", g.strIniFile)
    If Len(strSettings) > 0 Then
        Set astrSettings = New cGdArray
        astrSettings.SplitFields strSettings, ","
        
        For lIndex = 0 To astrSettings.Size - 1
            strIndex = Parse(astrSettings(lIndex), "=", 1)
            strValue = Parse(astrSettings(lIndex), "=", 2)
            
            m.astrButtons(CLng(strIndex)) = strValue
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information to the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrSettings As cGdArray        ' Array of settings information
    
    Set astrSettings = New cGdArray
    
    For lIndex = 0 To m.astrButtons.Size - 1
        astrSettings.Add Str(lIndex) & "=" & m.astrButtons(lIndex)
    Next lIndex
    
    SetIniFileProperty "ToolbarButtons", astrSettings.JoinFields(","), "TTSummary", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TcBtnFromString
'' Description: Fill a Trade Console button settings structure from a string
'' Inputs:      Settings
'' Returns:     Trade Console button settings object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TcBtnFromString(ByVal strSettings As String) As TcBtnSettings
On Error GoTo ErrSection:

    Dim TcBtn As TcBtnSettings             ' Return value for the function
    
    TcBtn.strName = Parse(strSettings, "|", 1)
    TcBtn.bShow = (Parse(strSettings, "|", 2) = "1")
    TcBtn.bShowInLists = (Parse(strSettings, "|", 3) = "1")
    
    TcBtnFromString = TcBtn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleButtons.TcBtnFromString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TcBtnToString
'' Description: Create a string from a Trade Console button settings structure
'' Inputs:      Trade Console button settings object
'' Returns:     Settings String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TcBtnToString(TcBtn As TcBtnSettings) As String
On Error GoTo ErrSection:

    Dim astrSettings As cGdArray        ' Settings in an array
    
    Set astrSettings = New cGdArray
    astrSettings(0) = TcBtn.strName
    If TcBtn.bShow Then astrSettings(1) = "1" Else astrSettings(1) = "0"
    If TcBtn.bShowInLists Then astrSettings(2) = "1" Else astrSettings(2) = "0"
        
    TcBtnToString = astrSettings.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeConsoleButtons.TcBtnToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.astrButtons = New cGdArray
    m.astrButtons.Create eGDARRAY_Strings, Button(eGDTcButtons_NumButtons)
    
    m.astrButtons(eGDTcButtons_Dashboard) = "Dashboard|1|0"
    m.astrButtons(eGDTcButtons_OpenOrders) = "Open Orders|1|0"
    m.astrButtons(eGDTcButtons_Positions) = "Positions|1|0"
    m.astrButtons(eGDTcButtons_Accounts) = "Accounts|1|0"
    m.astrButtons(eGDTcButtons_AutoTrading) = "Auto Trading|1|1"
    m.astrButtons(eGDTcButtons_TradeSenseOrders) = "TradeSense Orders|1|1"
    m.astrButtons(eGDTcButtons_ActivityLog) = "Activity Log|1|1"
    m.astrButtons(eGDTcButtons_TodaysFills) = "Todays Fills|1|1"
    m.astrButtons(eGDTcButtons_BuySell) = "Buy/Sell|1|1"
    m.astrButtons(eGDTcButtons_Settings) = "Settings|1|0"
    m.astrButtons(eGDTcButtons_Tracking) = "Tracking|1|1"
    m.astrButtons(eGDTcButtons_Reports) = "Reports|1|1"
    m.astrButtons(eGDTcButtons_Journals) = "Journals|1|1"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrButtons = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeConsoleButtons.Class_Terminate"
    
End Sub
