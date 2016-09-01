VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtPosition.cls
'' Description: Wrapper object for a position
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/06/2015   DAJ         Added clearing date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAccountID As Long                  ' Account ID for the position
    strAccountNumber As String          ' Account number for the position
    strSymbol As String                 ' Symbol or the position
    lSymbolID As Long                   ' Symbol ID for the position
    lCurrentPosition As Long            ' Current position
    lCarriedPosition As Long            ' Carried position
    dAverageEntry As Double             ' Average entry price for the position
    dClearingDate As Double             ' Clearing date for the position
    
    CarriedFills As cPtFills            ' Fills from previous sessions that make up carried position
    CurrentFills As cPtFills            ' Fills from current session that make up position
End Type
Private m As mPrivate

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
    m.strAccountNumber = g.Broker.AccountNumberForID(lAccountID)
End Property

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
    m.lAccountID = g.Broker.AccountIDForNumber(strAccountNumber)
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
End Property

Public Property Get CurrentPosition() As Long
    CurrentPosition = m.lCurrentPosition
End Property
Public Property Let CurrentPosition(ByVal lCurrentPosition As Long)
    m.lCurrentPosition = lCurrentPosition
End Property

Public Property Get CarriedPosition() As Long
    CarriedPosition = m.lCarriedPosition
End Property
Public Property Let CarriedPosition(ByVal lCarriedPosition As Long)
    m.lCarriedPosition = lCarriedPosition
End Property

Public Property Get AverageEntry() As Double
    AverageEntry = m.dAverageEntry
End Property
Public Property Let AverageEntry(ByVal dAverageEntry As Double)
    m.dAverageEntry = dAverageEntry
End Property

Public Property Get ClearingDate() As Double
    ClearingDate = m.dClearingDate
End Property
Public Property Let ClearingDate(ByVal dClearingDate As Double)
    m.dClearingDate = dClearingDate
End Property

Public Property Get CarriedFills() As cPtFills
    Set CarriedFills = m.CarriedFills
End Property
Public Property Let CarriedFills(ByVal Fills As cPtFills)
    Set m.CarriedFills = Fills
End Property

Public Property Get CurrentFills() As cPtFills
    Set CurrentFills = m.CurrentFills
End Property
Public Property Let CurrentFills(ByVal Fills As cPtFills)
    Set m.CurrentFills = Fills
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill the properties from the given information string
'' Inputs:      Position String
'' Returns:     None
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strPositionInfo As String)
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Information broken out from the string
    
    Set astrInfo = New cGdArray
    astrInfo.SplitFields strPositionInfo, vbTab
    
    AccountNumber = astrInfo(0)
    SymbolOrSymbolID = astrInfo(1)
    CurrentPosition = CLng(Val(astrInfo(2)))
    AverageEntry = Val(astrInfo(3))
    CarriedPosition = CLng(Val(astrInfo(4)))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtPosition.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a position information string from the properties
'' Inputs:      None
'' Returns:     Position String
''
'' Fields:      Account, Symbol, Position, Average Entry, Overnight Position
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString(Optional ByVal strDelimiter As String = vbTab) As String
On Error GoTo ErrSection:

    Dim astrInfo As cGdArray            ' Array of position information
    
    Set astrInfo = New cGdArray
    astrInfo.Create eGDARRAY_Strings
    
    astrInfo(0) = AccountNumber
    astrInfo(1) = Symbol
    astrInfo(2) = Str(CurrentPosition)
    astrInfo(3) = Str(AverageEntry)
    astrInfo(4) = Str(CarriedPosition)
    
    ToString = astrInfo.JoinFields(strDelimiter)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtPosition.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is created
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.lAccountID = 0&
    m.strSymbol = ""
    m.lSymbolID = 0&
    m.lCurrentPosition = 0&
    m.lCarriedPosition = kNullData
    m.dAverageEntry = 0#
    m.dClearingDate = 0#
    
    Set m.CarriedFills = New cPtFills
    Set m.CurrentFills = New cPtFills

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtPosition.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.CarriedFills = Nothing
    Set m.CurrentFills = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtPosition.Class_Terminate"
    
End Sub

