Attribute VB_Name = "mCommon"
Option Compare Text
Option Explicit

Type gGlobal
    strAppPath As String
End Type
Global g As gGlobal

Global Const C_ALL = "<ALL>"
Global Const C_BARSAGO = "BarsAgo"

'Return Types
Global Const gRetConstantNbr = 1
Global Const gRetConstantText = 2
Global Const gRetConstantBoolean = 6
Global Const gRetSeriesBoolean = 3
Global Const gRetSeriesNbr = 4
Global Const gRetSeriesText = 8
Global Const gRetBars = 5
Global Const gRetTrades = 7
Global Const gRetPortfolio = 9
Global Const gRetSystems = 10
Global Const gRetMarkets = 11
Global Const gRetRuleEntry = 14

'Implementation Types
Global Const gCompiled = 1
Global Const gTradeSense = 2
Global Const gInternalData = 3
Global Const gCompiledAction = 4

'System Navigator Token ID's.  Used for compatibility with System Navigator
'Make consistent later with new tokens defined in mComEval
Global Const gFUNC_NUMERIC = 1
Global Const gFUNC_BOOLEAN = 2
Global Const gFUNC_BOOLEAN_CONSTANT = 3
Global Const gFUNC_NUMERIC_CONSTANT = 4
Global Const gPARM_NUMERIC = 5
Global Const gPARM_BOOLEAN = 6
Global Const gPARM_BARS = 7
Global Const gADD = 8
Global Const gOR = 9
Global Const gMULTI = 10
Global Const gCOMPARE = 11
Global Const gAND = 12
Global Const gNUMERIC = 13
Global Const gLEFTPAREN = 14
Global Const gRIGHTPAREN = 15
Global Const gFLEFTPAREN = 16
Global Const gFRIGHTPAREN = 17
Global Const gError = 18
Global Const gOFFSET = 19
Global Const gTextT = 20
Global Const gOF = 21
Global Const gCOMMA = 22
Global Const gENTER = 23
Global Const gIF = 24
Global Const gPARM_TRADES = 25
Global Const gNOT = 26
Global Const gPARM_NUMERIC_ARRAY = 27
Global Const gPARM_BOOLEAN_ARRAY = 28
Global Const gFUNC_TEXT_CONSTANT = 30
Global Const gPARAGRAPH = 80
'the sysnav engine has set aside 45 for the comment brackets
'but we should never be setting the comments in coded text,
'so we should be alright with this id
Global Const gCOMMENT = 81  'this is the {} comment entity

'Security levels
Global Const EDITANDVIEW = 0
Global Const VIEWONLY = 1
Global Const NOVIEW = 2
Global Const NOLIST = 3

Global Const gUserErr = vbObjectError + 1000

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Global Const VARIABLE_PREFIX = "&"

Global gEditingArea         As cEditingArea

'Text caret screen coordinates
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Global PT                      As POINTAPI
Global ScreenX                 As Long
Global ScreenY                 As Long

'Forces a form to always to on top of all other forms
Public Sub FormOnTop(Handle As Long, OnTop As Boolean, pX As Long, pY As Long)
    Dim wFlags As Long, PosFlag As Long
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or _
        SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Select Case OnTop
        Case True
            PosFlag = HWND_TOPMOST
        Case False
            PosFlag = HWND_NOTOPMOST
    End Select
    SetWindowPos Handle, PosFlag, pX, pY, 0, 0, wFlags
End Sub
