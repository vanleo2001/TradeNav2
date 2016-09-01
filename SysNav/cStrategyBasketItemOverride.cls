VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBasketItemOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBasketItemOverride.cls
'' Description: Wrapper class for a strategy basket item override
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author  Description
'' 02/21/2012   DAJ     Added contract multiplier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lSymbolID As Long                   ' Symbol ID for the override
    strSymbol As String                 ' Symbol for the override
    strOverrides As String              ' Overrides
    dContractMultiplier As Double       ' Contract multiplier for the item
End Type
Private m As mPrivate

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0 Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
End Property

Public Property Get Overrides() As String
    Overrides = m.strOverrides
End Property
Public Property Let Overrides(ByVal strOverrides As String)
    m.strOverrides = strOverrides
End Property

Public Property Get ContractMultiplier() As Double
    ContractMultiplier = m.dContractMultiplier
End Property
Public Property Let ContractMultiplier(ByVal dContractMultiplier As Double)
    m.dContractMultiplier = dContractMultiplier
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Extract the fields from the given string
'' Inputs:      Info string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strInfo As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields split out from the string
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strInfo, ";"
    
    m.lSymbolID = CLng(Val(astrFields(0)))
    m.strSymbol = astrFields(1)
    m.strOverrides = astrFields(2)
    
    If astrFields.Size > 3 Then
        m.dContractMultiplier = Val(astrFields(3))
    Else
        m.dContractMultiplier = 1#
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cStrategyBasketOverride.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string from the class properties
'' Inputs:      None
'' Returns:     Info string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields split out from the string
    
    Set astrFields = New cGdArray
    astrFields(0) = Str(m.lSymbolID)
    astrFields(1) = m.strSymbol
    astrFields(2) = m.strOverrides
    astrFields(3) = Str(m.dContractMultiplier)

    ToString = astrFields.JoinFields(";")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketOverride.ToString"
    
End Function
