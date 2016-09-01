VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeRule.cls
'' Description: Wrapper class for a trade rule object
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Unique ID for the trade rule
    strName As String                   ' Name of the trade rule
    strAbbreviation As String           ' Abbreviation of the trade rule
    strDescription As String            ' Description of the trade rule
    strRequiredModule As String         ' Required module to see this trade rule
    
    bProvided As Boolean                ' Is this a provided or custom trade rule?
    nRuleType As eGDTradeRuleTypes      ' Type of the trade rule (Entry or Exit)
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Public Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Abbreviation() As String
    Abbreviation = m.strAbbreviation
End Property
Public Property Let Abbreviation(ByVal strAbbreviation As String)
    m.strAbbreviation = strAbbreviation
End Property

Public Property Get Description() As String
    Description = m.strDescription
End Property
Public Property Let Description(ByVal strDescription As String)
    m.strDescription = strDescription
End Property

Public Property Get RequiredModule() As String
    RequiredModule = m.strRequiredModule
End Property
Public Property Let RequiredModule(ByVal strRequiredModule As String)
    m.strRequiredModule = strRequiredModule
End Property

Public Property Get Provided() As Boolean
    Provided = m.bProvided
End Property
Public Property Let Provided(ByVal bProvided As Boolean)
    m.bProvided = bProvided
End Property

Public Property Get RuleType() As eGDTradeRuleTypes
    RuleType = m.nRuleType
End Property
Public Property Let RuleType(ByVal nRuleType As eGDTradeRuleTypes)
    m.nRuleType = nRuleType
End Property

Public Property Get FullName() As String
    If Len(m.strName) > 0 Then
        FullName = m.strAbbreviation & " - " & m.strName
    Else
        FullName = m.strAbbreviation
    End If
End Property
Public Property Let FullName(ByVal strFullName As String)
    m.strAbbreviation = Parse(strFullName, "-", 1)
    m.strName = Parse(strFullName, "-", 2)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Load up the properties from the string that came from the file
'' Inputs:      String, Provided?, Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strTradeRule As String, ByVal bProvided As Boolean, ByVal nRuleType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim astrRule As cGdArray            ' Trade rule fields broken out into an array
    
    Set astrRule = New cGdArray
    astrRule.SplitFields strTradeRule, vbTab
    
    m.lID = CLng(Val(astrRule(0)))
    m.strAbbreviation = Parse(astrRule(1), "-", 1)
    m.strName = Parse(astrRule(1), "-", 2)
    m.strRequiredModule = astrRule(2)
    m.strDescription = astrRule(3)
    
    m.bProvided = bProvided
    m.nRuleType = nRuleType

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRule.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build string to send to file from the properties
'' Inputs:      None
'' Returns:     File String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information to return
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn(0) = Str(m.lID)
    astrReturn(1) = m.strAbbreviation & " - " & m.strName
    astrReturn(2) = m.strRequiredModule
    astrReturn(3) = m.strDescription
    
    ToString = astrReturn.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRule.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ComboString
'' Description: Build a flex grid combo string from the properties
'' Inputs:      None
'' Returns:     Combo String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ComboString() As String
On Error GoTo ErrSection:

    ComboString = "#" & Str(m.lID) & ";" & FullName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRule.ComboString"
    
End Function
