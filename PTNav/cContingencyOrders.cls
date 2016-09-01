VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContingencyOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cContingencyOrders.cls
'' Description: Wrapper object for contingency orders for an order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/31/2014   DAJ         Have user specify TIF on contingecy orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kVersion As Long = 1&

Private Type mPrivate
    bUseProfitTarget As Boolean         ' Is the profit target being used?
    bProfitInDollars As Boolean         ' Is the profit target being specified in dollars?
    dProfitDollarAmount As Double       ' Dollar amount for the profit target
    bProfitInPoints As Boolean          ' Is the profit target being specified in points?
    dProfitPointsAmount As Double       ' Points amount for the profit target
    lProfitOrderId As Long              ' Access Order ID for the profit target order
    nProfitTif As eTT_TimeInForce       ' Time-In-Force for the profit target order

    bUseStopLoss As Boolean             ' Is the stop loss being used?
    bStopInDollars As Boolean           ' Is the stop loss being specified in dollars?
    dStopDollarAmount As Double         ' Dollar amount for the stop loss
    bStopInPoints As Boolean            ' Is the stop loss being specified in points?
    dStopPointsAmount As Double         ' Points amount for the stop loss
    lStopOrderId As Long                ' Access order ID for the stop loss order
    nStopTif As eTT_TimeInForce         ' Time-In-Force for the stop loss order
End Type
Private m As mPrivate

Public Property Get UseProfitTarget() As Boolean
    UseProfitTarget = m.bUseProfitTarget
End Property
Public Property Let UseProfitTarget(ByVal bUseProfitTarget As Boolean)
    m.bUseProfitTarget = bUseProfitTarget
End Property

Public Property Get ProfitInDollars() As Boolean
    ProfitInDollars = m.bProfitInDollars
End Property
Public Property Let ProfitInDollars(ByVal bProfitInDollars As Boolean)
    m.bProfitInDollars = bProfitInDollars
End Property

Public Property Get ProfitDollarAmount() As Double
    ProfitDollarAmount = m.dProfitDollarAmount
End Property
Public Property Let ProfitDollarAmount(ByVal dProfitDollarAmount As Double)
    If dProfitDollarAmount > 2000 Then
        dProfitDollarAmount = dProfitDollarAmount
    End If
    m.dProfitDollarAmount = dProfitDollarAmount
End Property

Public Property Get ProfitPointsAmount() As Double
    ProfitPointsAmount = m.dProfitPointsAmount
End Property
Public Property Let ProfitPointsAmount(ByVal dProfitPointsAmount As Double)
    m.dProfitPointsAmount = dProfitPointsAmount
End Property

Public Property Get ProfitOrderId() As Long
    ProfitOrderId = m.lProfitOrderId
End Property
Public Property Let ProfitOrderId(ByVal lProfitOrderId As Long)
    m.lProfitOrderId = lProfitOrderId
End Property

Public Property Get ProfitTif() As eTT_TimeInForce
    ProfitTif = m.nProfitTif
End Property
Public Property Let ProfitTif(ByVal nProfitTif As eTT_TimeInForce)
    m.nProfitTif = nProfitTif
End Property

Public Property Get UseStopLoss() As Boolean
    UseStopLoss = m.bUseStopLoss
End Property
Public Property Let UseStopLoss(ByVal bUseStopLoss As Boolean)
    m.bUseStopLoss = bUseStopLoss
End Property

Public Property Get StopInDollars() As Boolean
    StopInDollars = m.bStopInDollars
End Property
Public Property Let StopInDollars(ByVal bStopInDollars As Boolean)
    m.bStopInDollars = bStopInDollars
End Property

Public Property Get StopDollarAmount() As Double
    StopDollarAmount = m.dStopDollarAmount
End Property
Public Property Let StopDollarAmount(ByVal dStopDollarAmount As Double)
    If dStopDollarAmount > 2000 Then
        dStopDollarAmount = dStopDollarAmount
    End If
    m.dStopDollarAmount = dStopDollarAmount
End Property

Public Property Get StopPointsAmount() As Double
    StopPointsAmount = m.dStopPointsAmount
End Property
Public Property Let StopPointsAmount(ByVal dStopPointsAmount As Double)
    m.dStopPointsAmount = dStopPointsAmount
End Property

Public Property Get StopOrderId() As Long
    StopOrderId = m.lStopOrderId
End Property
Public Property Let StopOrderId(ByVal lStopOrderId As Long)
    m.lStopOrderId = lStopOrderId
End Property

Public Property Get StopTif() As eTT_TimeInForce
    StopTif = m.nStopTif
End Property
Public Property Let StopTif(ByVal nStopTif As eTT_TimeInForce)
    m.nStopTif = nStopTif
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear the class members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo ErrSection:

    m.bUseProfitTarget = False
    m.bProfitInDollars = True
    m.dProfitDollarAmount = 0#
    m.bProfitInPoints = False
    m.dProfitPointsAmount = 0#
    m.lProfitOrderId = 0&
    m.nProfitTif = eTT_TimeInForce_Day
    
    m.bUseStopLoss = False
    m.bStopInDollars = True
    m.dStopDollarAmount = 0#
    m.bStopInPoints = False
    m.dStopPointsAmount = 0#
    m.lStopOrderId = 0&
    m.nStopTif = eTT_TimeInForce_Day

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cContingencyOrders.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill the class members from the given string
'' Inputs:      Contingency Options
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strContingencyOptions As String)
On Error GoTo ErrSection:

    Dim lVersion As Long                ' Version of the string
    Dim astrOptions As cGdArray         ' Options split out into an array
    Dim astrSubOptions As cGdArray      ' Sub options split out into an array
    
    If (Len(strContingencyOptions) > 0) And (InStr(strContingencyOptions, ";") <> 0) Then
        Set astrOptions = New cGdArray
        astrOptions.SplitFields strContingencyOptions, ";"
        Set astrSubOptions = New cGdArray
        
        lVersion = CLng(Val(astrOptions(0)))
        
        astrSubOptions.SplitFields astrOptions(1), ","
        m.bUseProfitTarget = CBool(Val(astrSubOptions(0)))
        m.bProfitInDollars = CBool(Val(astrSubOptions(1)))
        ProfitDollarAmount = Val(astrSubOptions(2))
        m.bProfitInPoints = CBool(Val(astrSubOptions(3)))
        m.dProfitPointsAmount = Val(astrSubOptions(4))
        m.lProfitOrderId = CLng(Val(astrSubOptions(5)))
        If astrSubOptions.Size > 6 Then
            m.nProfitTif = CLng(Val(astrSubOptions(6)))
        End If
        
        astrSubOptions.SplitFields astrOptions(2), ","
        m.bUseStopLoss = CBool(Val(astrSubOptions(0)))
        m.bStopInDollars = CBool(Val(astrSubOptions(1)))
        StopDollarAmount = Val(astrSubOptions(2))
        m.bStopInPoints = CBool(Val(astrSubOptions(3)))
        m.dStopPointsAmount = Val(astrSubOptions(4))
        m.lStopOrderId = CLng(Val(astrSubOptions(5)))
        If astrSubOptions.Size > 6 Then
            m.nStopTif = CLng(Val(astrSubOptions(6)))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cContingencyOrders.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string from the class members
'' Inputs:      None
'' Returns:     Contingency Options
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrOptions As cGdArray         ' Options split out into an array
    Dim astrSubOptions As cGdArray      ' Sub options split out into an array
    
    strReturn = ""
    If m.bUseProfitTarget Or m.bUseStopLoss Then
        Set astrOptions = New cGdArray
        astrOptions.Create eGDARRAY_Strings, 3
        Set astrSubOptions = New cGdArray
        astrSubOptions.Create eGDARRAY_Strings, 6
        
        astrOptions(0) = Str(kVersion)
        
        astrSubOptions(0) = Str(m.bUseProfitTarget)
        astrSubOptions(1) = Str(m.bProfitInDollars)
        astrSubOptions(2) = Str(m.dProfitDollarAmount)
        astrSubOptions(3) = Str(m.bProfitInPoints)
        astrSubOptions(4) = Str(m.dProfitPointsAmount)
        astrSubOptions(5) = Str(m.lProfitOrderId)
        astrSubOptions(6) = Str(m.nProfitTif)
        astrOptions(1) = astrSubOptions.JoinFields(",")
        
        astrSubOptions(0) = Str(m.bUseStopLoss)
        astrSubOptions(1) = Str(m.bStopInDollars)
        astrSubOptions(2) = Str(m.dStopDollarAmount)
        astrSubOptions(3) = Str(m.bStopInPoints)
        astrSubOptions(4) = Str(m.dStopPointsAmount)
        astrSubOptions(5) = Str(m.lStopOrderId)
        astrSubOptions(6) = Str(m.nStopTif)
        astrOptions(2) = astrSubOptions.JoinFields(",")
        
        strReturn = astrOptions.JoinFields(";")
    End If
    
    ToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cContingencyOrders.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the contingency options
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cContingencyOrders
On Error GoTo ErrSection:

    Dim aCopy As New cContingencyOrders
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cContingencyOrders.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the contingency options
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cContingencyOrders.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cContingencyOrders.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cContingencyOrders.Class_Terminate"
    
End Sub
