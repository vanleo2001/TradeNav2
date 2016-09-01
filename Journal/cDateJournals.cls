VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cDateJournals.cls
'' Description: Class that contains a collection of date journals
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the date journal object
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    DateJournals As cGdTree             ' Collection of date journals
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.DateJournals.Count
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cDateJournal
Attribute Item.VB_UserMemId = 0
    If m.DateJournals.Exists(vIndexOrKey) Then
        Set Item = m.DateJournals(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, ByVal DateJournal As cDateJournal)
    If m.DateJournals.Exists(vIndexOrKey) Then
        Set m.DateJournals(vIndexOrKey) = DateJournal
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given date journal to the collection
'' Inputs:      Date Journal
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(ByVal DateJournal As cDateJournal) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    If m.DateJournals.Exists(Str(DateJournal.DateJournalID)) Then
        Set m.DateJournals(Str(DateJournal.DateJournalID)) = DateJournal
        lReturn = m.DateJournals.Index(Str(DateJournal.DateJournalID))
    Else
        lReturn = m.DateJournals.Add(DateJournal, Str(DateJournal.DateJournalID))
    End If
    
    Add = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo ErrSection:

    m.DateJournals.Clear
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.Clear"
    
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

    Set m.DateJournals = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.DateJournals = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.Class_Terminate"
    
End Sub

