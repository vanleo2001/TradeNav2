VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCattleStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCattleStats.cls
'' Description: Class for calculating statistics for Turnkey
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/23/2014   DAJ         Fix for overflow error in FinalADG function
'' 02/28/2014   DAJ         Pass LotColumn collection into cTurnkeyStats
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/21/2014   DAJ         Change generic objects back to gdTree in public function
'' 04/08/2014   DAJ         Added Cattle Cost per CWT
'' 05/22/2014   DAJ         Renamed cTurnkeyStats to cCattleStats
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Lot As cBrokerMessage               ' Lot object
    LotDetails As cGdTree               ' Collection of lot details
    
    LotColumnMap As cGdTree             ' Map of lot column KeyValueField to ID
    strHeadInColumn As String           ' Lot column ID for the Head In
    strNumShippedColumn As String       ' Lot column ID for the Number Shipped
    strNumDeadColumn As String          ' Lot column ID for the Number Dead
    
    lEarliestDateIn As Long             ' Earliest date of incoming cattle
    lLatestDateOut As Long              ' Latest date of outgoing cattle
    lHeadIn As Long                     ' Total number of head in
    lHeadOut As Long                    ' Total number of head out
    lNumDead As Long                    ' Total number of dead
    lCalendarDaysOnFeed As Long         ' Calendar days on feed
    lTotalHeadDays As Long              ' Total head days
    lTotalPayWeight As Long             ' Total pay weight
    lTotalSalesWeight As Long           ' Total sales weight
    dProjCurWeight As Double            ' Projected current weight
    
    dTotalFeedCost As Double            ' Total feed cost
    dTotalDebit As Double               ' Total of debits
    dTotalDebitCog As Double            ' Total of COG debits
    dTotalCredit As Double              ' Total of credits
    dTotalCreditCog As Double           ' Total of COG credits
    dTotalCattleCost As Double          ' Total cost of the cattle
    dTotalSalesAmount As Double         ' Total sales amount
End Type
Private m As mPrivate

Public Property Get CalendarDaysOnFeed() As Long
    CalendarDaysOnFeed = m.lCalendarDaysOnFeed
End Property

Public Property Get TotalHeadDays() As Long
    TotalHeadDays = m.lTotalHeadDays
End Property

Public Property Get TotalGain() As Long
    If m.lHeadOut = m.lHeadIn Then
        TotalGain = m.lTotalSalesWeight - m.lTotalPayWeight
    Else
        TotalGain = 0&
    End If
End Property

Public Property Get FinalADG() As Double
    If (m.lHeadOut = m.lHeadIn) And (m.lTotalHeadDays > 0&) Then
        FinalADG = TotalGain / m.lTotalHeadDays
    Else
        FinalADG = 0#
    End If
End Property

Public Property Get AveragePayWeight() As Double
    If m.lHeadIn = 0& Then
        AveragePayWeight = 0#
    Else
        AveragePayWeight = Round(CDbl(m.lTotalPayWeight) / CDbl(m.lHeadIn), 1)
    End If
End Property

Public Property Get AverageSalesWeight() As Double
    If ((m.lHeadOut - m.lNumDead) > 0) And (m.lTotalSalesWeight > 0) Then
        AverageSalesWeight = CDbl(m.lTotalSalesWeight) / CDbl(m.lHeadOut - m.lNumDead)
    Else
        AverageSalesWeight = 0#
    End If
End Property

Public Property Get CattleCostPerCwt() As Double
    If m.lTotalPayWeight = 0 Then
        CattleCostPerCwt = 0#
    Else
        CattleCostPerCwt = (m.dTotalCattleCost / (CDbl(m.lTotalPayWeight) / 100))
    End If
End Property

Public Property Get EstimatedTotalGain() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    Dim dCurrentWeight As Double        ' Current weight for existing cattle
    Dim dDeadWeight As Double           ' Pay weight for the dead cattle

    dReturn = 0#

    If m.dProjCurWeight <> 0# Then
        ' (PCW * (HI - HO) + TSW + (APW * ND)) - TPW
        
        dCurrentWeight = m.dProjCurWeight * CDbl(m.lHeadIn - m.lHeadOut)
        dDeadWeight = AveragePayWeight * CDbl(m.lNumDead)
        
        dReturn = (dCurrentWeight + m.lTotalSalesWeight + dDeadWeight) - CDbl(m.lTotalSalesWeight)
    End If
    
    EstimatedTotalGain = dReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.EstimatedTotalGain.Get"
    
End Property

Public Property Get FeedCostPerCwt() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    Dim lTotalGain As Long              ' Total gain
    Dim dEstimatedGain As Double        ' Estimated gain

    dReturn = 0#
    lTotalGain = TotalGain
    dEstimatedGain = EstimatedTotalGain
    
    If (m.lHeadOut = m.lHeadIn) And (lTotalGain <> 0&) Then
        dReturn = m.dTotalFeedCost / (lTotalGain / 100)
    ElseIf (dEstimatedGain <> 0#) Then
        dReturn = m.dTotalFeedCost / (dEstimatedGain / 100)
    End If
    
    FeedCostPerCwt = dReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.FeedCostPerCwt.Get"
    
End Property

Public Property Get TotalAllCost() As Double
    TotalAllCost = m.dTotalFeedCost + m.dTotalDebit + m.dTotalDebitCog + m.dTotalCattleCost - m.dTotalCredit - m.dTotalCreditCog
End Property

Public Property Get CurrentBreakEven() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    Dim dCurrentWeight As Double        ' Current weight for existing cattle

    dReturn = 0#

    If m.dProjCurWeight <> 0# Then
        ' TAC / ((PCW * (HI - NS - ND)) / 100)
        
        dCurrentWeight = m.dProjCurWeight * CDbl(m.lHeadIn - m.lHeadOut)
        
        If dCurrentWeight <> 0# Then
            dReturn = TotalAllCost / (dCurrentWeight / 100)
        End If
    End If
    
    CurrentBreakEven = dReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.CurrentBreakEven.Get"
    
End Property

Public Property Get FinalBreakEven() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    
    dReturn = 0#
    If (m.lHeadIn = m.lHeadOut) And (m.lTotalSalesWeight <> 0&) Then
        dReturn = m.dTotalSalesAmount / (CDbl(m.lTotalSalesWeight) / 100)
    End If

    FinalBreakEven = dReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.FinalBreakEven.Get"
    
End Property

Public Property Get CostOfGain() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    
    dReturn = 0#
    If (m.lHeadIn = m.lHeadOut) And (TotalGain <> 0) Then
        dReturn = (m.dTotalSalesAmount - m.dTotalCattleCost) / (TotalGain / 100)
    End If
    
    CostOfGain = dReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.CostOfGain"
    
End Property

Public Property Get FinalCostOfGain() As Double
On Error GoTo ErrSection:

    Dim dReturn As Double               ' Return value for the property
    
    dReturn = 0#
    If (m.lHeadIn = m.lHeadOut) And (TotalGain <> 0) Then
        dReturn = (m.dTotalFeedCost + m.dTotalDebitCog - m.dTotalCreditCog) / (TotalGain / 100)
    End If
    
    FinalCostOfGain = dReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCattleStats.FinalCostOfGain"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Calculate
'' Description: Calculate the stats based on the given information
'' Inputs:      Lot, Lot Details
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Calculate(ByVal LotDetails As cGdTree, ByVal LotColumns As cGdTree)
On Error GoTo ErrSection:

    LoadLotColumnMap LotColumns
    Set m.LotDetails = LotDetails
        
    CalcTotals
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleStats.Calculate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotals
'' Description: Calculate totals to use in other calculations
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcTotals()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotDetail As cBrokerMessage     ' Lot detail
    Dim dDate As Double                 ' Date of the detail
    Dim lHeadDaysIn As Long             ' Total head days in
    Dim lHeadDaysOut As Long            ' Total head days out
    Dim strPayWeightCol As String       ' Pay weight column
    Dim strSalesWeightCol As String     ' Sales weight column
    Dim strTotalFeedCostCol As String   ' Total feed cost column
    Dim strProjCurWeightCol As String   ' Projected current weight column
    Dim strDebitCol As String           ' Debit column
    Dim strDebitCogCol As String        ' Debit COG column
    Dim strCreditCol As String          ' Credit column
    Dim strCreditCogCol As String       ' Credit COG column
    Dim strCattleCostCol As String      ' Cost of the cattle column
    Dim strSalesAmountCol As String     ' Sales amount column
    
    m.lEarliestDateIn = Abs(kNullData)
    m.lLatestDateOut = kNullData
    m.lHeadIn = 0&
    m.lHeadOut = 0&
    m.lNumDead = 0&
    lHeadDaysIn = 0&
    lHeadDaysOut = 0&
    m.lTotalPayWeight = 0&
    m.lTotalSalesWeight = 0&
    m.dTotalFeedCost = 0#
    m.dProjCurWeight = 0#
    m.dTotalDebit = 0#
    m.dTotalDebitCog = 0#
    m.dTotalCredit = 0#
    m.dTotalCreditCog = 0#
    m.dTotalCattleCost = 0#
    m.dTotalSalesAmount = 0#
    
    strPayWeightCol = m.LotColumnMap("TotalPayWeight")
    strSalesWeightCol = m.LotColumnMap("TotalSalesWeight")
    strTotalFeedCostCol = m.LotColumnMap("TotalFeedCost")
    strProjCurWeightCol = m.LotColumnMap("ProjectedCurrentWeight")
    strDebitCol = m.LotColumnMap("Debit")
    strDebitCogCol = m.LotColumnMap("DebitCOG")
    strCreditCol = m.LotColumnMap("Credit")
    strCreditCogCol = m.LotColumnMap("CreditCOG")
    strCattleCostCol = m.LotColumnMap("TotalCostOfCattle")
    strSalesAmountCol = m.LotColumnMap("TotalSalesAmount")
    
    For lIndex = 1 To m.LotDetails.Count
        Set LotDetail = m.LotDetails(lIndex)
        
        Select Case LotDetail("LotColumnID")
            Case m.strHeadInColumn
                dDate = Val(LotDetail("Date"))
                If (dDate <> 0) Then
                    If (dDate < m.lEarliestDateIn) Then
                        m.lEarliestDateIn = dDate
                    End If
                    
                    lHeadDaysIn = lHeadDaysIn + ((Date - dDate) * CLng(Val(LotDetail("Value"))))
                End If
                m.lHeadIn = m.lHeadIn + CLng(Val(LotDetail("Value")))
        
            Case m.strNumShippedColumn, m.strNumDeadColumn
                dDate = Val(LotDetail("Date"))
                If (dDate <> 0) Then
                    If (dDate > m.lLatestDateOut) Then
                        m.lLatestDateOut = dDate
                    End If
                    
                    lHeadDaysOut = lHeadDaysOut + ((Date - dDate) * CLng(Val(LotDetail("Value"))))
                End If
                m.lHeadOut = m.lHeadOut + CLng(Val(LotDetail("Value")))
                
                If LotDetail("LotColumnID") = m.strNumDeadColumn Then
                    m.lNumDead = m.lNumDead + CLng(Val(LotDetail("Value")))
                End If
        
            Case strPayWeightCol
                m.lTotalPayWeight = m.lTotalPayWeight + CLng(Val(LotDetail("Value")))
            
            Case strSalesWeightCol
                m.lTotalSalesWeight = m.lTotalSalesWeight + CLng(Val(LotDetail("Value")))
                
            Case strTotalFeedCostCol
                m.dTotalFeedCost = m.dTotalFeedCost + Val(LotDetail("Value"))
                
            Case strProjCurWeightCol
                m.dProjCurWeight = Val(LotDetail("Value"))
                
            Case strDebitCol
                m.dTotalDebit = m.dTotalDebit + Val(LotDetail("Value"))
            
            Case strDebitCogCol
                m.dTotalDebitCog = m.dTotalDebitCog + Val(LotDetail("Value"))
                
            Case strCreditCol
                m.dTotalCredit = m.dTotalCredit + Val(LotDetail("Value"))
            
            Case strCreditCogCol
                m.dTotalCreditCog = m.dTotalCreditCog + Val(LotDetail("Value"))
            
            Case strCattleCostCol
                m.dTotalCattleCost = m.dTotalCattleCost + Val(LotDetail("Value"))
                
            Case strSalesAmountCol
                m.dTotalSalesAmount = m.dTotalSalesAmount + Val(LotDetail("Value"))
                
        End Select
    Next lIndex
    
    If m.lHeadIn = 0 Then
        m.lCalendarDaysOnFeed = 0&
        m.lTotalHeadDays = 0&
    ElseIf m.lHeadIn = m.lHeadOut Then
        m.lCalendarDaysOnFeed = (m.lLatestDateOut - m.lEarliestDateIn) + 1&
        m.lTotalHeadDays = lHeadDaysIn - lHeadDaysOut
    Else
        m.lCalendarDaysOnFeed = (Date - m.lEarliestDateIn) + 1&
        m.lTotalHeadDays = lHeadDaysIn - lHeadDaysOut
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleStats.CalcTotals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLotColumnMap
'' Description: Load up the lot column map
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLotColumnMap(ByVal LotColumns As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object

    Set m.LotColumnMap = New cGdTree
        
    For lIndex = 1 To LotColumns.Count
        Set LotColumn = LotColumns(lIndex)
        m.LotColumnMap.Add Str(LotColumn.ID), LotColumn.KeyValueField
        
        Select Case LotColumn.KeyValueField
            Case "HeadIn"
                m.strHeadInColumn = Str(LotColumn.ID)
            Case "NumberShip"
                m.strNumShippedColumn = Str(LotColumn.ID)
            Case "NumberDeads"
                m.strNumDeadColumn = Str(LotColumn.ID)
        End Select
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleStats.LoadLotColumnMap"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotColumnID
'' Description: Determine the lot column ID for the given field
'' Inputs:      Key Value Field
'' Returns:     Lot Column ID ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LotColumnID(ByVal strKeyValueField As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Not m.LotColumnMap Is Nothing Then
        If m.LotColumnMap.Exists(strKeyValueField) Then
            strReturn = m.LotColumnMap(strKeyValueField)
        End If
    End If
    
    LotColumnID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCattleStats.LotColumnID"
    
End Function

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

    Set m.Lot = New cBrokerMessage
    Set m.LotDetails = New cGdTree
    
    Set m.LotColumnMap = New cGdTree
    m.strHeadInColumn = ""
    m.strNumShippedColumn = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleStats.Class_Initialize"
    
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

    Set m.Lot = Nothing
    Set m.LotDetails = Nothing
    Set m.LotColumnMap = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCattleStats.Class_Terminate"
    
End Sub
