VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditorOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==========================================================================
' Class: cEditorOptions
'
' Description: Use to save and retrieve global properties use to control the
' Rule editor display of information
'
'==========================================================================

Private Type mPrivate
    iRowsToDisplay As Integer
    bFillWordsOn As Boolean
    lEditorOn As Long
    iEFontSize As Integer
    strEFontName As String
    
    'Colors
    lFunctionsColor As Long
    bFunctionsBoldFace As Boolean
    bFunctionsItalics As Boolean
    lParmColor As Long
    bParmBoldFace As Boolean
    bParmItalics As Boolean
    lOperatorsColor As Long
    bOperatorsBoldFace As Boolean
    bOperatorsItalics As Boolean
    lErrorColor As Long
    bErrorBoldFace As Boolean
    bErrorItalics As Boolean
End Type
Private m As mPrivate

Property Get RowsToDisplay() As Integer
    RowsToDisplay = m.iRowsToDisplay
End Property
Property Get FillWordsOn() As Boolean
    FillWordsOn = m.bFillWordsOn
End Property
Property Get EditorOn() As Boolean
    EditorOn = m.lEditorOn
End Property
Property Get eFontName() As String
    eFontName = m.strEFontName
End Property
Property Get eFontSize() As Integer
    eFontSize = m.iEFontSize
End Property
Property Get FunctionsColor() As Long
    FunctionsColor = m.lFunctionsColor
End Property
Property Get FunctionsBoldFace() As Boolean
    FunctionsBoldFace = m.bFunctionsBoldFace
End Property
Property Get FunctionsItalics() As Boolean
    FunctionsItalics = m.bFunctionsItalics
End Property
Property Get ParmColor() As Long
    ParmColor = m.lParmColor
End Property
Property Get ParmBoldFace() As Boolean
    ParmBoldFace = m.bParmBoldFace
End Property
Property Get ParmItalics() As Boolean
    ParmItalics = m.bParmItalics
End Property
Property Get OperatorsColor() As Long
    OperatorsColor = m.lOperatorsColor
End Property
Property Get OperatorsBoldFace() As Boolean
    OperatorsBoldFace = m.bOperatorsBoldFace
End Property
Property Get OperatorsItalics() As Boolean
    OperatorsItalics = m.bOperatorsItalics
End Property
Property Get ErrorColor() As Long
    ErrorColor = m.lErrorColor
End Property
Property Get ErrorBoldFace() As Boolean
    ErrorBoldFace = m.bErrorBoldFace
End Property
Property Get ErrorItalics() As Boolean
    ErrorItalics = m.bErrorItalics
End Property

Property Let RowsToDisplay(pData As Integer)
    m.iRowsToDisplay = pData
End Property
Property Let FillWordsOn(pData As Boolean)
    m.bFillWordsOn = pData
End Property
Property Let EditorOn(pData As Boolean)
    m.lEditorOn = pData
End Property
Property Let eFontName(pData As String)
    m.strEFontName = pData
End Property
Property Let eFontSize(pData As Integer)
    m.iEFontSize = pData
End Property
Property Let FunctionsColor(pData As Long)
    m.lFunctionsColor = pData
End Property
Property Let FunctionsBoldFace(pData As Boolean)
    m.bFunctionsBoldFace = pData
End Property
Property Let FunctionsItalics(pData As Boolean)
    m.bFunctionsItalics = pData
End Property
Property Let ParmColor(pData As Long)
    m.lParmColor = pData
End Property
Property Let ParmBoldFace(pData As Boolean)
    m.bParmBoldFace = pData
End Property
Property Let ParmItalics(pData As Boolean)
    m.bParmItalics = pData
End Property
Property Let OperatorsColor(pData As Long)
    m.lOperatorsColor = pData
End Property
Property Let OperatorsBoldFace(pData As Boolean)
    m.bOperatorsBoldFace = pData
End Property
Property Let OperatorsItalics(pData As Boolean)
    m.bOperatorsItalics = pData
End Property
Property Let ErrorColor(pData As Long)
    m.lErrorColor = pData
End Property
Property Let ErrorBoldFace(pData As Boolean)
    m.bErrorBoldFace = pData
End Property
Property Let ErrorItalics(pData As Boolean)
    m.bErrorItalics = pData
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save changes to Editor options
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:
    
    ' Default Rows to display to 1 to 8...
    If m.iRowsToDisplay < 1 Then m.iRowsToDisplay = 1
    If m.iRowsToDisplay > 8 Then m.iRowsToDisplay = 8
    
    SetIniFileProperty "RowsToDisplay", m.iRowsToDisplay, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "FillWordsOn", m.bFillWordsOn, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "EditorOn", m.lEditorOn, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "eFontSize", m.iEFontSize, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "eFontName", m.strEFontName, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "FunctionsColor", m.lFunctionsColor, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "FunctionsBoldFace", m.bFunctionsBoldFace, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "FunctionsItalics", m.bFunctionsItalics, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ParmColor", m.lParmColor, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ParmBoldFace", m.bParmBoldFace, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ParmItalics", m.bParmItalics, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "OperatorsColor", m.lOperatorsColor, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "OperatorsBoldFace", m.bOperatorsBoldFace, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "OperatorsItalics", m.bOperatorsItalics, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ErrorColor", m.lErrorColor, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ErrorBoldFace", m.bErrorBoldFace, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    SetIniFileProperty "ErrorItalics", m.bErrorItalics, "Editor", AddSlash(g.strAppPath) & "navdefs.ini"
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEditorOptions.Save", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load Editor options
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    m.iRowsToDisplay = GetIniFileProperty("RowsToDisplay", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bFillWordsOn = GetIniFileProperty("FillWordsOn", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.lEditorOn = GetIniFileProperty("EditorOn", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.iEFontSize = GetIniFileProperty("eFontSize", "9", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.strEFontName = GetIniFileProperty("eFontName", "Arial", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.lFunctionsColor = GetIniFileProperty("FunctionsColor", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bFunctionsBoldFace = GetIniFileProperty("FunctionsBoldFace", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bFunctionsItalics = GetIniFileProperty("FunctionsItalics", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.lParmColor = GetIniFileProperty("ParmColor", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bParmBoldFace = GetIniFileProperty("ParmBoldFace", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bParmItalics = GetIniFileProperty("ParmItalics", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.lOperatorsColor = GetIniFileProperty("OperatorsColor", "16711680", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bOperatorsBoldFace = GetIniFileProperty("OperatorsBoldFace", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bOperatorsItalics = GetIniFileProperty("OperatorsItalics", "0", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.lErrorColor = GetIniFileProperty("ErrorColor", "255", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bErrorBoldFace = GetIniFileProperty("ErrorBoldFace", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    m.bErrorItalics = GetIniFileProperty("ErrorItalics", "-1", "Editor", _
        AddSlash(g.strAppPath) & "navdefs.ini")
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEditorOptions.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

