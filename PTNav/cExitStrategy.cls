VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExitStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cExitStrategy.cls
'' Description: Class to handle storage and retrieval of an exit strategy
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Bvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/07/2009   DAJ         Allow for broker held linked orders for simple auto exits
'' 05/17/2010   DAJ         Added support for TradeSense orders
'' 11/07/2011   DAJ         Added required module flag
'' 04/17/2013   DAJ         Flatten if Stop on wrong side, Cancel manual orders on opposite side
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strFileName As String               ' Filename for the exit strategy
    strStrategyName As String           ' Unique Name for the exit strategy
    strDescription As String            ' Description for the exit strategy
    bSpecifyLots As Boolean             ' Specify Lots or Equal Lots?
    bExitEntirePosition As Boolean      ' Exit the entire position on the profit side?
    bUseTarget1 As Boolean              ' Are we using the first profit target?
    dTarget1Ticks As Double             ' Number of ticks away from entry to place first profit target
    lTarget1Quantity As Long            ' Quantity for the first profit target
    bUseTarget2 As Boolean              ' Are we using the second profit target?
    dTarget2Ticks As Double             ' Number of ticks away from entry to place second profit target
    lTarget2Quantity As Long            ' Quantity for the second profit target
    bUseTarget3 As Boolean              ' Are we using the third profit target?
    dTarget3Ticks As Double             ' Number of ticks away from entry to place third profit target
    lTarget3Quantity As Long            ' Quantity for the third profit target
    nStopLossType As eGDStopLossType    ' Type of the stop loss
    dStopLossTicks As Double            ' Number of ticks away from entry to place the stop loss
    bUseWithLimit As Boolean            ' Is the stop loss a Stop with Limit order?
    dWithLimitTicks As Double           ' Number of ticks away from entry to place the with limit amount
    dAfterTicks As Double               ' Number of ticks away from entry before we will move the stop
    bUseMoveTo As Boolean               ' Move the stop loss order after a number of ticks?
    dMoveToTicks As Double              ' Number of ticks away from entry to move stop loss order to
    bUseTrail As Boolean                ' Do we want to utilize a trailing stop?
    dTrailTicks As Double               ' Number of ticks away from order to trail
    lTickMode As Long                   ' Full ticks, Half ticks, or Quarter ticks
    bCancelOpposite As Boolean          ' Cancel all orders on the opposite side?
    bMarketIfWrongSide As Boolean       ' Submit market if stop on wrong side of the market?
    
    bProvided As Boolean                ' Is this a provided exit order strategy?
    
    ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
    nProfitType As eGDProfitTargetType  ' Type of the profit target
    tsoProfitLong As cTradeSenseOrder   ' Trade Sense order for the Long Profit Target
    tsoProfitShort As cTradeSenseOrder  ' Trade Sense order for the Short Profit Target
    tsoStopLong As cTradeSenseOrder     ' Trade Sense order for the Long Stop Loss
    tsoStopShort As cTradeSenseOrder    ' Trade Sense order for the Short Stop Loss
    
    strRequired As String               ' Required module to see or use this auto exit
End Type
Private m As mPrivate

Public Property Get FileName() As String
    FileName = m.strFileName
End Property
Public Property Let FileName(ByVal strFileName As String)
    m.strFileName = strFileName
End Property

Public Property Get StrategyName() As String
    StrategyName = m.strStrategyName
End Property
Public Property Let StrategyName(ByVal strStrategyName As String)
    m.strStrategyName = strStrategyName
End Property

Public Property Get Description() As String
    Description = m.strDescription
End Property
Public Property Let Description(ByVal strDescription As String)
    m.strDescription = strDescription
End Property

Public Property Get SpecifyLots() As Boolean
    SpecifyLots = m.bSpecifyLots
End Property
Public Property Let SpecifyLots(ByVal bSpecifyLots As Boolean)
    m.bSpecifyLots = bSpecifyLots
End Property

Public Property Get ExitEntirePosition() As Boolean
    ExitEntirePosition = m.bExitEntirePosition
End Property
Public Property Let ExitEntirePosition(ByVal bExitEntirePosition As Boolean)
    m.bExitEntirePosition = bExitEntirePosition
End Property

Public Property Get UseTarget1() As Boolean
    UseTarget1 = m.bUseTarget1 = True
End Property
Public Property Let UseTarget1(ByVal bUseTarget1 As Boolean)
    m.bUseTarget1 = bUseTarget1
End Property

Public Property Get Target1Ticks() As Double
    Target1Ticks = m.dTarget1Ticks
End Property
Public Property Let Target1Ticks(ByVal dTarget1Ticks As Double)
    m.dTarget1Ticks = dTarget1Ticks
End Property

Public Property Get Target1Quantity() As Long
    Target1Quantity = m.lTarget1Quantity
End Property
Public Property Let Target1Quantity(ByVal lTarget1Quantity As Long)
    m.lTarget1Quantity = lTarget1Quantity
End Property

Public Property Get UseTarget2() As Boolean
    UseTarget2 = m.bUseTarget2
End Property
Public Property Let UseTarget2(ByVal bUseTarget2 As Boolean)
    m.bUseTarget2 = bUseTarget2
End Property

Public Property Get Target2Ticks() As Double
    Target2Ticks = m.dTarget2Ticks
End Property
Public Property Let Target2Ticks(ByVal dTarget2Ticks As Double)
    m.dTarget2Ticks = dTarget2Ticks
End Property

Public Property Get Target2Quantity() As Long
    Target2Quantity = m.lTarget2Quantity
End Property
Public Property Let Target2Quantity(ByVal lTarget2Quantity As Long)
    m.lTarget2Quantity = lTarget2Quantity
End Property

Public Property Get UseTarget3() As Boolean
    UseTarget3 = m.bUseTarget3
End Property
Public Property Let UseTarget3(ByVal bUseTarget3 As Boolean)
    m.bUseTarget3 = bUseTarget3
End Property

Public Property Get Target3Ticks() As Double
    Target3Ticks = m.dTarget3Ticks
End Property
Public Property Let Target3Ticks(ByVal dTarget3Ticks As Double)
    m.dTarget3Ticks = dTarget3Ticks
End Property

Public Property Get Target3Quantity() As Long
    Target3Quantity = m.lTarget3Quantity
End Property
Public Property Let Target3Quantity(ByVal lTarget3Quantity As Long)
    m.lTarget3Quantity = lTarget3Quantity
End Property

Public Property Get StopLossType() As eGDStopLossType
    StopLossType = m.nStopLossType
End Property
Public Property Let StopLossType(ByVal nStopLossType As eGDStopLossType)
    m.nStopLossType = nStopLossType
End Property

Public Property Get StopLossTicks() As Double
    StopLossTicks = m.dStopLossTicks
End Property
Public Property Let StopLossTicks(ByVal dStopLossTicks As Double)
    m.dStopLossTicks = dStopLossTicks
End Property

Public Property Get UseWithLimit() As Boolean
    UseWithLimit = m.bUseWithLimit
End Property
Public Property Let UseWithLimit(ByVal bUseWithLimit As Boolean)
    m.bUseWithLimit = bUseWithLimit
End Property

Public Property Get WithLimitTicks() As Double
    WithLimitTicks = m.dWithLimitTicks
End Property
Public Property Let WithLimitTicks(ByVal dWithLimitTicks As Double)
    m.dWithLimitTicks = dWithLimitTicks
End Property

Public Property Get AfterTicks() As Double
    AfterTicks = m.dAfterTicks
End Property
Public Property Let AfterTicks(ByVal dAfterTicks As Double)
    m.dAfterTicks = dAfterTicks
End Property

Public Property Get UseMoveTo() As Boolean
    UseMoveTo = m.bUseMoveTo
End Property
Public Property Let UseMoveTo(ByVal bUseMoveTo As Boolean)
    m.bUseMoveTo = bUseMoveTo
End Property

Public Property Get MoveToTicks() As Double
    MoveToTicks = m.dMoveToTicks
End Property
Public Property Let MoveToTicks(ByVal dMoveToTicks As Double)
    m.dMoveToTicks = dMoveToTicks
End Property

Public Property Get UseTrail() As Boolean
    UseTrail = m.bUseTrail
End Property
Public Property Let UseTrail(ByVal bUseTrail As Boolean)
    m.bUseTrail = bUseTrail
End Property

Public Property Get TrailTicks() As Double
    TrailTicks = m.dTrailTicks
End Property
Public Property Let TrailTicks(ByVal dTrailTicks As Double)
    m.dTrailTicks = dTrailTicks
End Property

Public Property Get TickMode() As Long
    TickMode = m.lTickMode
End Property
Public Property Let TickMode(ByVal lTickMode As Long)
    m.lTickMode = lTickMode
End Property

Public Property Get CancelOpposite() As Boolean
    CancelOpposite = m.bCancelOpposite
End Property
Public Property Let CancelOpposite(ByVal bCancelOpposite As Boolean)
    m.bCancelOpposite = bCancelOpposite
End Property

Public Property Get MarketIfWrongSide() As Boolean
    MarketIfWrongSide = m.bMarketIfWrongSide
End Property
Public Property Let MarketIfWrongSide(ByVal bMarketIfWrongSide As Boolean)
    m.bMarketIfWrongSide = bMarketIfWrongSide
End Property

Public Property Get Provided() As Boolean
    Provided = m.bProvided
End Property
Public Property Let Provided(ByVal bProvided As Boolean)
    m.bProvided = bProvided
End Property

' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
Public Property Get ProfitTargetType() As eGDProfitTargetType
    ProfitTargetType = m.nProfitType
End Property
Public Property Let ProfitTargetType(ByVal nProfitTargetType As eGDProfitTargetType)
    m.nProfitType = nProfitTargetType
End Property

Public Property Get TsProfitLong() As cTradeSenseOrder
    Set TsProfitLong = m.tsoProfitLong
End Property
Public Property Let TsProfitLong(ByVal tsoProfitLong As cTradeSenseOrder)
    Set m.tsoProfitLong = tsoProfitLong
End Property

Public Property Get TsProfitShort() As cTradeSenseOrder
    Set TsProfitShort = m.tsoProfitShort
End Property
Public Property Let TsProfitShort(ByVal tsoProfitShort As cTradeSenseOrder)
    Set m.tsoProfitShort = tsoProfitShort
End Property

Public Property Get TsStopLong() As cTradeSenseOrder
    Set TsStopLong = m.tsoStopLong
End Property
Public Property Let TsStopLong(ByVal tsoStopLong As cTradeSenseOrder)
    Set m.tsoStopLong = tsoStopLong
End Property

Public Property Get TsStopShort() As cTradeSenseOrder
    Set TsStopShort = m.tsoStopShort
End Property
Public Property Let TsStopShort(ByVal tsoStopShort As cTradeSenseOrder)
    Set m.tsoStopShort = tsoStopShort
End Property

Public Property Get TsProfitOrder(ByVal lPosition As Long) As cTradeSenseOrder
    If lPosition = 0 Then
        Set TsProfitOrder = Nothing
    ElseIf lPosition > 0 Then
        Set TsProfitOrder = m.tsoProfitLong
    Else
        Set TsProfitOrder = m.tsoProfitShort
    End If
End Property

Public Property Get TsStopOrder(ByVal lPosition As Long) As cTradeSenseOrder
    If lPosition = 0 Then
        Set TsStopOrder = Nothing
    ElseIf lPosition > 0 Then
        Set TsStopOrder = m.tsoStopLong
    Else
        Set TsStopOrder = m.tsoStopShort
    End If
End Property

Public Property Get Required() As String
    Required = m.strRequired
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the exit strategy information from the database
'' Inputs:      Strategy Filename
'' Returns:     True if loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal strStrategyFile As String) As Boolean
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array to be dumped to a file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key for the record
    Dim strValue As String              ' Value for the record
    Dim lPos As Long                    ' Position of '=' in the string
    
    If astrFile.FromFile(AddSlash(App.Path) & strStrategyFile) Then
        m.strFileName = strStrategyFile
        
        ' Default this value to True if not in file...
        m.bExitEntirePosition = True
        
        For lIndex = 0 To astrFile.Size - 1
            ' Don't use Parse here because the user could have an "=" in the name of
            ' their strategy (10/15/2008 DAJ)...
            lPos = InStr(astrFile(lIndex), "=")
            If lPos > 0 Then
                strKey = Left(astrFile(lIndex), lPos - 1)
                strValue = Mid(astrFile(lIndex), lPos + 1)
            Else
                strKey = astrFile(lIndex)
                strValue = ""
            End If
            
            Select Case UCase(strKey)
                Case "STRATEGYNAME"
                    m.strStrategyName = strValue
                Case "DESCRIPTION"
                    m.strDescription = strValue
                Case "SPECIFYLOTS"
                    m.bSpecifyLots = CBool(Val(strValue))
                Case "EXITENTIREPOSITION"
                    m.bExitEntirePosition = CBool(Val(strValue))
                Case "USETARGET1"
                    m.bUseTarget1 = CBool(Val(strValue))
                Case "TARGET1TICKS"
                    m.dTarget1Ticks = Val(strValue)
                Case "TARGET1QUANTITY"
                    m.lTarget1Quantity = CLng(Val(strValue))
                Case "USETARGET2"
                    m.bUseTarget2 = CBool(Val(strValue))
                Case "TARGET2TICKS"
                    m.dTarget2Ticks = Val(strValue)
                Case "TARGET2QUANTITY"
                    m.lTarget2Quantity = CLng(Val(strValue))
                Case "USETARGET3"
                    m.bUseTarget3 = CBool(Val(strValue))
                Case "TARGET3TICKS"
                    m.dTarget3Ticks = Val(strValue)
                Case "TARGET3QUANTITY"
                    m.lTarget3Quantity = CLng(Val(strValue))
                Case "STOPLOSSTYPE"
                    m.nStopLossType = CLng(Val(strValue))
                Case "STOPLOSSTICKS"
                    m.dStopLossTicks = Val(strValue)
                Case "USEWITHLIMIT"
                    m.bUseWithLimit = CBool(Val(strValue))
                Case "WITHLIMITTICKS"
                    m.dWithLimitTicks = Val(strValue)
                Case "AFTERTICKS"
                    m.dAfterTicks = Val(strValue)
                Case "USEMOVETO"
                    m.bUseMoveTo = CBool(Val(strValue))
                Case "MOVETOTICKS"
                    m.dMoveToTicks = Val(strValue)
                Case "USETRAIL"
                    m.bUseTrail = CBool(Val(strValue))
                Case "TRAILTICKS"
                    m.dTrailTicks = Val(strValue)
                Case "TICKMODE"
                    m.lTickMode = CLng(Val(strValue))
                    
                ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
                Case "PROFITTARGETTYPE"
                    m.nProfitType = CLng(Val(strValue))
                Case "TSPROFITLONG"
                    m.tsoProfitLong.FromString strValue
                Case "TSPROFITSHORT"
                    m.tsoProfitShort.FromString strValue
                Case "TSSTOPLONG"
                    m.tsoStopLong.FromString strValue
                Case "TSSTOPSHORT"
                    m.tsoStopShort.FromString strValue
                    
                ' 11/07/2011 DAJ: New field for the required module flag...
                Case "REQUIRED"
                    m.strRequired = strValue
            
                ' 04/17/2013 DAJ: New fields for how to handle situations...
                Case "CANCELOPPOSITE"
                    m.bCancelOpposite = CBool(Val(strValue))
                Case "MARKETIFWRONGSIDE"
                    m.bMarketIfWrongSide = CBool(Val(strValue))
            End Select
        Next lIndex
        
        If UCase(Parse(strStrategyFile, "\", 1)) = "PROVIDED" Then
            m.bProvided = True
        Else
            m.bProvided = False
        End If
        
        Load = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExitStrategy.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the exit strategy information to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array to be dumped to a file

    astrFile.Add "StrategyName=" & m.strStrategyName
    astrFile.Add "Description=" & m.strDescription
    astrFile.Add "SpecifyLots=" & Str(m.bSpecifyLots)
    astrFile.Add "ExitEntirePosition=" & Str(m.bExitEntirePosition)
    astrFile.Add "UseTarget1=" & Str(m.bUseTarget1)
    astrFile.Add "Target1Ticks=" & Str(m.dTarget1Ticks)
    astrFile.Add "Target1Quantity=" & Str(m.lTarget1Quantity)
    astrFile.Add "UseTarget2=" & Str(m.bUseTarget2)
    astrFile.Add "Target2Ticks=" & Str(m.dTarget2Ticks)
    astrFile.Add "Target2Quantity=" & Str(m.lTarget2Quantity)
    astrFile.Add "UseTarget3=" & Str(m.bUseTarget3)
    astrFile.Add "Target3Ticks=" & Str(m.dTarget3Ticks)
    astrFile.Add "Target3Quantity=" & Str(m.lTarget3Quantity)
    astrFile.Add "StopLossType=" & Str(m.nStopLossType)
    astrFile.Add "StopLossTicks=" & Str(m.dStopLossTicks)
    astrFile.Add "UseWithLimit=" & Str(m.bUseWithLimit)
    astrFile.Add "WithLimitTicks=" & Str(m.dWithLimitTicks)
    astrFile.Add "AfterTicks=" & Str(m.dAfterTicks)
    astrFile.Add "UseMoveTo=" & Str(m.bUseMoveTo)
    astrFile.Add "MoveToTicks=" & Str(m.dMoveToTicks)
    astrFile.Add "UseTrail=" & Str(m.bUseTrail)
    astrFile.Add "TrailTicks=" & Str(m.dTrailTicks)
    astrFile.Add "TickMode=" & Str(m.lTickMode)
    
    ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
    astrFile.Add "ProfitTargetType=" & Str(m.nProfitType)
    astrFile.Add "TsProfitLong=" & m.tsoProfitLong.ToString
    astrFile.Add "TsProfitShort=" & m.tsoProfitShort.ToString
    astrFile.Add "TsStopLong=" & m.tsoStopLong.ToString
    astrFile.Add "TsStopShort=" & m.tsoStopShort.ToString
                    
    ' 11/07/2011 DAJ: New field for the required module flag...
    astrFile.Add "Required=" & m.strRequired
    
    ' 04/17/2013 DAJ: New fields for how to handle situations...
    astrFile.Add "CancelOpposite=" & Str(m.bCancelOpposite)
    astrFile.Add "MarketIfWrongSide=" & Str(m.bMarketIfWrongSide)
    
    If m.bProvided Then
        m.strFileName = "Provided\" & m.strStrategyName & ".XOS"
    Else
        m.strFileName = "Custom\" & m.strStrategyName & ".XOS"
    End If
    astrFile.ToFile AddSlash(App.Path) & m.strFileName
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitStrategy.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsSimpleBracket
'' Description: Determine if the auto exit is a simple bracket or not
'' Inputs:      None
'' Returns:     True if Simple Bracket, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSimpleBracket() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If (m.bUseTarget1 = True) And (m.bUseTarget2 = False) And (m.bUseTarget3 = False) Then
        If m.nStopLossType <> eGDStopLossType_None Then
            bReturn = True
        End If
    End If
    
    IsSimpleBracket = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExitStrategy.IsSimpleBracket"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UdpateTsOrderPos
'' Description: Update the position info for the Trade Sense orders
'' Inputs:      Account position
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateTsOrderPos(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    If Not AcctPos Is Nothing Then
        If m.nProfitType = eGDProfitTargetType_TradeSense Then
            If AcctPos.CurrentPositionSnapshot > 0 Then
                m.tsoProfitLong.UpdatePosition AcctPos
            ElseIf AcctPos.CurrentPositionSnapshot < 0 Then
                m.tsoProfitShort.UpdatePosition AcctPos
            End If
        End If
        
        If m.nStopLossType = eGDStopLossType_TradeSense Then
            If AcctPos.CurrentPositionSnapshot > 0 Then
                m.tsoProfitLong.UpdatePosition AcctPos
            Else
                m.tsoProfitShort.UpdatePosition AcctPos
            End If
        End If
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitStrategy.UpdateTsOrderPos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the exit strategy
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cExitStrategy
On Error GoTo ErrSection:

    Dim aCopy As New cExitStrategy
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExitStrategy.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the exit strategy
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    
    ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
    Set m.tsoProfitLong = mCopyFrom.tsoProfitLong.MakeCopy
    Set m.tsoProfitShort = mCopyFrom.tsoProfitShort.MakeCopy
    Set m.tsoStopLong = mCopyFrom.tsoStopLong.MakeCopy
    Set m.tsoStopShort = mCopyFrom.tsoStopShort.MakeCopy

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitStrategy.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    m.bExitEntirePosition = True
    
    ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
    m.nProfitType = eGDProfitTargetType_Standard
    Set m.tsoProfitLong = New cTradeSenseOrder
    Set m.tsoProfitShort = New cTradeSenseOrder
    Set m.tsoStopLong = New cTradeSenseOrder
    Set m.tsoStopShort = New cTradeSenseOrder
    
    m.bCancelOpposite = False
    m.bMarketIfWrongSide = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitStrategy.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    ' 04/14/2010 DAJ: New fields for the new Trade Sense auto exit orders...
    Set m.tsoProfitLong = Nothing
    Set m.tsoProfitShort = Nothing
    Set m.tsoStopLong = Nothing
    Set m.tsoStopShort = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExitStrategy.Class_Terminate"
    
End Sub
