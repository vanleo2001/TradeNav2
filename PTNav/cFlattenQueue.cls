VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlattenQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFlattenQueue.cls
'' Description: Queue of items for "flatten" or "cancel all"
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/04/2011   DAJ         After auto exit cancel done, tell it to refresh position
'' 12/11/2012   DAJ         Use the flatten queue for position reversals
'' 06/24/2013   DAJ         Timer Logging
'' 02/18/2015   DAJ         Added the Remove function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents m_tmrFlatten As Timer
Attribute m_tmrFlatten.VB_VarHelpID = -1

Private Type mPrivate
    Flatten As cGdTree                  ' Flatten queue
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Flatten Timer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(tmrFlatten As Timer)
On Error GoTo ErrSection:

    Set m_tmrFlatten = tmrFlatten

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlattenQueue.Init"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddToFlattenQueue
'' Description: Add an item to the flatten queue
'' Inputs:      Account, Symbol, Auto Trade ID, Operation, Skip Order ID,
''              Auto Exit Only?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddToFlattenQueue(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, ByVal nOperation As eGDFlattenQueueOperations, Optional ByVal lSkipOrderID As Long = -1&, Optional ByVal bAutoExitOnly As Boolean = False)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the queue
    Dim FlattenItem As cFlatten         ' Flatten item
    
    strKey = strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab & Str(nOperation)
    If m.Flatten.Exists(strKey) = False Then
        Set FlattenItem = New cFlatten
        With FlattenItem
            .AccountNumber = strAccount
            .Symbol = strSymbol
            .AutoTradeItemID = lAtID
            .Operation = nOperation
            .SkipOrderID = lSkipOrderID
            .AutoExitOnly = bAutoExitOnly
        End With
        
        g.Broker.BrokerDebug FlattenItem.Broker, "Added to Flatten Queue: '" & strKey & "'"
        m.Flatten.Add FlattenItem, strKey
        
        If m_tmrFlatten.Enabled = False Then
            m_tmrFlatten.Enabled = True
        End If
    Else
        Set FlattenItem = m.Flatten(strKey)
        g.Broker.BrokerDebug FlattenItem.Broker, "Attempted to add '" & strKey & "' to the flatten queue, but it already exists"
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlattenQueue.AddToFlattenQueue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove an item from the flatten queue
'' Inputs:      Account, Symbol, Auto Trade ID, Reason for removing
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Remove(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, ByVal strReason As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim FlattenItem As cFlatten         ' Flatten item
    Dim strTradeSymbol As String        ' Trade symbol
    
    bReturn = False
    For lIndex = m.Flatten.Count To 1 Step -1
        Set FlattenItem = m.Flatten(lIndex)
        If FlattenItem.AccountNumber = strAccount Then
            strTradeSymbol = GetSymbol(FlattenItem.TradeSymbolOrSymbolID)
            
            If strTradeSymbol = strSymbol Then
                If FlattenItem.AutoTradeItemID = lAtID Then
                    g.Broker.BrokerDebug FlattenItem.Broker, "Removing from Flatten Queue: '" & FlattenItem.Key & "' because '" & strReason & "'"
                    m.Flatten.Remove lIndex
                    bReturn = True
                    
                    If Len(strReason) > 0 Then
                        frmOnlineBroker.AddDialogMessage FlattenItem.Operation & " for " & strTradeSymbol & " in " & g.Broker.AccountNameForID(FlattenItem.AccountID) & " was cancelled because " & strReason, "!", , FlattenItem.Operation & " for " & strTradeSymbol
                    Else
                        frmOnlineBroker.AddDialogMessage FlattenItem.Operation & " for " & strTradeSymbol & " in " & g.Broker.AccountNameForID(FlattenItem.AccountID) & " was cancelled", "!", , FlattenItem.Operation & " for " & strTradeSymbol
                    End If
                End If
            End If
        End If
    Next lIndex
    
    Remove = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlattenQueue.Remove"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsGettingFlattened
'' Description: Is the given information in the queue?
'' Inputs:      Account, Symbol, Auto Trade ID, Operation
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsGettingFlattened(ByVal strAccount As String, ByVal strSymbol As String, ByVal lAtID As Long, ByVal nOperation As eGDFlattenQueueOperations) As Boolean
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the queue

    strKey = strAccount & vbTab & strSymbol & vbTab & Str(lAtID) & vbTab & Str(nOperation)
    IsGettingFlattened = m.Flatten.Exists(strKey)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFlattenQueue.IsGettingFlattened"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_tmrFlatten_Timer
'' Description: Attempt to flatten items that may be in the queue
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_tmrFlatten_Timer()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim FlattenItem As cFlatten         ' Flatten item

    TimerStart "cFlattenQueue.m_tmrFlatten"
    If m.Flatten.Count = 0 Then
        m_tmrFlatten.Enabled = False
    Else
        For lIndex = m.Flatten.Count To 1 Step -1
            Set FlattenItem = m.Flatten(lIndex)
            Select Case FlattenItem.Operation
                Case eGDFlattenQueueOperation_CancelAll
                    If FlattenItem.AutoExitOnly Then
                        If FlattenItem.CancelAllAutoExit Then
                            g.Broker.BrokerDebug FlattenItem.Broker, "Removed from Flatten Queue: '" & FlattenItem.Key & "' ( CancelAllAutoExit = True )"
                            m.Flatten.Remove lIndex
                            g.OrderStrategies.RefreshPosition FlattenItem.AccountNumberOrID, FlattenItem.SymbolOrSymbolID
                        End If
                    Else
                        If FlattenItem.CancelAll Then
                            g.Broker.BrokerDebug FlattenItem.Broker, "Removed from Flatten Queue: '" & FlattenItem.Key & "' ( CancelAll = True )"
                            m.Flatten.Remove lIndex
                        End If
                    End If
                    
                Case eGDFlattenQueueOperation_Flatten
                    If FlattenItem.Flatten Then
                        g.Broker.BrokerDebug FlattenItem.Broker, "Removed from Flatten Queue: '" & FlattenItem.Key & "' ( Flatten = True )"
                        m.Flatten.Remove lIndex
                    End If
                    
                Case eGDFlattenQueueOperation_Reverse
                    If FlattenItem.Reverse Then
                        g.Broker.BrokerDebug FlattenItem.Broker, "Removed from Flatten Queue: '" & FlattenItem.Key & "' ( Reverse = True )"
                        m.Flatten.Remove lIndex
                    End If
                    
            End Select
        Next lIndex
    End If
    TimerEnd "cFlattenQueue.m_tmrFlatten", m_tmrFlatten.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlattenQueue.m_tmrFlatten_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Flatten = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlattenQueue.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Flatten = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFlattenQueue.Class_Terminate"
    
End Sub
