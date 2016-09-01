VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEarliestDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cEarliestDate.cls
'' Description: Wrapper class for a collection of earliest dates per account
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
''
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    EarliestDate As cGdTree             ' Collection of earliest dates per account
    lEarliestDate As Long               ' Earliest date
End Type
Private m As mPrivate

Public Property Get Item(ByVal lAccountID As Long) As Long
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    
    lReturn = Abs(kNullData)
    If m.EarliestDate.Exists(Str(lAccountID)) = True Then
        lReturn = m.EarliestDate(Str(lAccountID))
    End If
    
    Item = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cEarliestDate.Item.Get"
    
End Property
Public Property Let Item(ByVal lAccountID As Long, ByVal lDate As Long)
On Error GoTo ErrSection:

    If m.EarliestDate.Exists(Str(lAccountID)) = True Then
        m.EarliestDate(Str(lAccountID)) = lDate
    Else
        m.EarliestDate.Add lDate, Str(lAccountID)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cEarliestDate.Item.Let"
    
End Property

Public Property Get EarliestDate() As Long
    EarliestDate = m.lEarliestDate
End Property
Private Property Let EarliestDate(ByVal lEarliestDate As Long)
    m.lEarliestDate = lEarliestDate
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the collection and the earliest date
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo ErrSection:

    m.EarliestDate.Clear
    EarliestDate = Abs(kNullData)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEarliestDate.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Update
'' Description: Update the earliest date as appropriate
'' Inputs:      Account ID, Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Update(ByVal lAccountID As Long, ByVal lDate As Long)
On Error GoTo ErrSection:

    If lDate < Item(lAccountID) Then
        Item(lAccountID) = lDate
    End If
    If lDate < EarliestDate Then
        EarliestDate = lDate
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEarliestDate.Update"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DumpDebug
'' Description: Dump the information to the appropriate log file
'' Inputs:      Broker, Title
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DumpDebug(ByVal nBroker As eTT_AccountType, ByVal strTitle As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lDate As Long                   ' Date out of the collection
    Dim lAccountID As Long              ' Account ID

    g.Broker.BrokerDebug nBroker, vbTab & vbTab & strTitle & " " & Str(EarliestDate) & " ( " & DateFormat(EarliestDate, MM_DD_YYYY) & " )"
    For lIndex = 1 To m.EarliestDate.Count
        lDate = m.EarliestDate(lIndex)
        lAccountID = CLng(Val(m.EarliestDate.Key(lIndex)))
        
        g.Broker.BrokerDebug nBroker, vbTab & vbTab & strTitle & " " & Str(lDate) & " ( " & DateFormat(lDate, MM_DD_YYYY) & " ) for account '" & g.Broker.AccountNumberForID(lAccountID) & "'"
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cEarliestDate.DumpDebug"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.EarliestDate = New cGdTree
    EarliestDate = Abs(kNullData)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEarliestDate.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Destroy the class members when the object is destroyed
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.EarliestDate = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cEarliestDate.Class_Terminate"
    
End Sub
