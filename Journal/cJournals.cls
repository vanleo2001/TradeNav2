VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournals.cls
'' Description: Class that holds a collection of order journals
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/23/2009   DAJ         Created
'' 09/20/2011   DAJ         Added the LoadForDate function
'' 09/28/2011   DAJ         Ensure order ID array gets updated when new journal added
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the order journal object
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Journals As cGdTree                 ' Collection of order journals
    astrByOrderID As cGdArray           ' Array sorted by order ID
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cJournal
Attribute Item.VB_UserMemId = 0
    If m.Journals.Exists(vIndexOrKey) Then
        Set Item = m.Journals.Item(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

Public Property Let Item(ByVal vIndexOrKey As Variant, ByVal Journal As cJournal)
    Set m.Journals.Item(vIndexOrKey) = Journal
End Property

Public Property Get Count() As Long
    Count = m.Journals.Count
End Property

Public Property Get Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Journals.Exists(vIndexOrKey)
End Property

Public Function Add(Journal As cJournal) As Long
On Error GoTo ErrSection:
    
    Dim lReturn As Long                 ' Return value for the function
    Dim lPos As Long                    ' Position in the array

    lReturn = m.Journals.Add(Journal, Str(Journal.JournalID))
    If m.astrByOrderID.BinarySearch(Str(Journal.OrderID) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        m.astrByOrderID(lPos) = Str(Journal.OrderID) & vbTab & Str(lReturn)
    Else
        m.astrByOrderID.Add Str(Journal.OrderID) & vbTab & Str(lReturn), lPos
    End If
    
    Add = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournals.Add"
    
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Journals.Remove(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear out the collections
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
On Error GoTo ErrSection:

    m.Journals.Clear
    m.astrByOrderID.Clear

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournals.Clear"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalForOrderID
'' Description: Return the journal for the given order ID
'' Inputs:      Order ID
'' Returns:     Journal (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function JournalForOrderID(ByVal lOrderID As Long) As cJournal
On Error GoTo ErrSection:

    Dim lPos As Long                    ' Position in the array
    Dim lItemNum As Long                ' Item number in the collection
    Dim Journal As cJournal             ' Journal item to return
    
    Set Journal = Nothing
    If m.astrByOrderID.BinarySearch(Str(lOrderID) & vbTab, lPos, eGdSort_MatchUsingSearchStringLength) Then
        lItemNum = CLng(Val(Parse(m.astrByOrderID(lPos), vbTab, 2)))
        Set Journal = Item(lItemNum)
    End If
    
    Set JournalForOrderID = Journal

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournals.JournalForOrderID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when class is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Journals = New cGdTree
    
    Set m.astrByOrderID = New cGdArray
    m.astrByOrderID.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournals.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Journals = Nothing
    Set m.astrByOrderID = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournals.Class_Terminate"

End Sub
