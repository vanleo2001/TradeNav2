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

Private Type mPrivate
    iRowsToDisplay As Integer
    bFillWordsOn As Boolean
    lEditorOn As Long
    iFontSize As Integer
    strFontName As String
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
Property Get FontName() As String
    FontName = m.strFontName
End Property
Property Get FontSize() As Integer
    FontSize = m.iFontSize
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
Property Let FontName(pData As String)
    m.strFontName = pData
End Property
Property Let FontSize(pData As Integer)
    m.iFontSize = pData
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


'This method is used to save changes to Editor options
Public Sub Save()
On Error GoTo ErrSection:
    
    'Default Rows to display to 1 to 8...
    If m.iRowsToDisplay < 1 Then m.iRowsToDisplay = 1
    If m.iRowsToDisplay > 8 Then m.iRowsToDisplay = 8
    
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="RowsToDisplay", setting:=m.iRowsToDisplay
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FillWordsOn", setting:=m.bFillWordsOn
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="EditorOn", setting:=m.lEditorOn
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FontSize", setting:=m.iFontSize
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FontName", setting:=m.strFontName
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FunctionsColor", setting:=m.lFunctionsColor
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FunctionsBoldFace", setting:=m.bFunctionsBoldFace
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="FunctionsItalics", setting:=m.bFunctionsItalics
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ParmColor", setting:=m.lParmColor
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ParmBoldFace", setting:=m.bParmBoldFace
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ParmItalics", setting:=m.bParmItalics
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="OperatorsColor", setting:=m.lOperatorsColor
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="OperatorsBoldFace", setting:=m.bOperatorsBoldFace
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="OperatorsItalics", setting:=m.bOperatorsItalics
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ErrorColor", setting:=m.lErrorColor
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ErrorBoldFace", setting:=m.bErrorBoldFace
    SaveSetting appname:="MoneyManager", Section:="EditorSettings", _
                        Key:="ErrorItalics", setting:=m.bErrorItalics
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditorOptions.Save", eGDRaiseError_Raise, g.strAppPath

End Sub

'This method is used to Load the current Rule Editor Settings
Public Sub Load()
On Error GoTo ErrSection:

    m.iRowsToDisplay = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="RowsToDisplay", Default:=0)
    m.bFillWordsOn = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FillWordsOn", Default:=-1)
    m.lEditorOn = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="EditorOn", Default:=-1)
    m.iFontSize = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FontSize", Default:=9)
    m.strFontName = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FontName", Default:="Arial")
    m.lFunctionsColor = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FunctionsColor", Default:=0)
    m.bFunctionsBoldFace = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FunctionsBoldFace", Default:=0)
    m.bFunctionsItalics = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="FunctionsItalics", Default:=-1)
    m.lParmColor = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ParmColor", Default:=0)
    m.bParmBoldFace = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ParmBoldFace", Default:=-1)
    m.bParmItalics = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ParmItalics", Default:=-1)
    m.lOperatorsColor = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="OperatorsColor", Default:=16711680)
    m.bOperatorsBoldFace = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="OperatorsBoldFace", Default:=0)
    m.bOperatorsItalics = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="OperatorsItalics", Default:=0)
    m.lErrorColor = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ErrorColor", Default:=255)
    m.bErrorBoldFace = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ErrorBoldFace", Default:=-1)
    m.bErrorItalics = GetSetting(appname:="MoneyManager", _
            Section:="EditorSettings", Key:="ErrorItalics", Default:=-1)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEditorOptions.Load", eGDRaiseError_Raise, g.strAppPath

End Sub
