VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournal.cls
'' Description: Wrapper class for the order journal table in the database
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
'' 09/01/2009   DAJ         Added Null Checks on Load
'' 09/20/2011   DAJ         Added the journal date
'' 09/22/2011   DAJ         Added filename and caption for attached chart image
'' 09/23/2011   DAJ         Delete image file when order journal is deleted
'' 09/28/2011   DAJ         Notify date journals form when journal is deleted
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the order journal object
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lJournalID As Long                  ' Unique ID for the Journal entry
    lOrderID As Long                    ' ID of the associated order
    dNoteDate As Double                 ' Date/Time the journal entry was made
    strNote As String                   ' Notes on the journal entry
    strAction As String                 ' Action taken by the order
    strFeelings As String               ' How the trader was feeling at time of order
    strWhyTrade As String               ' Reason that the order was placed
    strThoughts As String               ' What trader was thinking when placed order
    lEmotionNumber As Long              ' 0-10 of how user felt when order was placed
    lAccountID As Long                  ' ID of the account for the order
    dJournalDate As Double              ' Date of the journal (without the time)
    lSymbolID As Long                   ' Symbol ID for the journal
    strSymbol As String                 ' Symbol for the journal
    JournalImages As cGdTree            ' Collection of journal images
End Type
Private m As mPrivate

Public Property Get JournalID() As Long
    JournalID = m.lJournalID
End Property
Public Property Let JournalID(ByVal lJournalID As Long)
    m.lJournalID = lJournalID
End Property

Public Property Get OrderID() As Long
    OrderID = m.lOrderID
End Property
Public Property Let OrderID(ByVal lOrderID As Long)
    m.lOrderID = lOrderID
End Property

Public Property Get NoteDate() As Double
    NoteDate = m.dNoteDate
End Property
Public Property Let NoteDate(ByVal dNoteDate As Double)
    m.dNoteDate = dNoteDate
End Property

Public Property Get Note() As String
    Note = m.strNote
End Property
Public Property Let Note(ByVal strNote As String)
    m.strNote = strNote
End Property

Public Property Get Action() As String
    Action = m.strAction
End Property
Public Property Let Action(ByVal strAction As String)
    m.strAction = strAction
End Property

Public Property Get Feelings() As String
    Feelings = m.strFeelings
End Property
Public Property Let Feelings(ByVal strFeelings As String)
    m.strFeelings = strFeelings
End Property

Public Property Get WhyTrade() As String
    WhyTrade = m.strWhyTrade
End Property
Public Property Let WhyTrade(ByVal strWhyTrade As String)
    m.strWhyTrade = strWhyTrade
End Property

Public Property Get Thoughts() As String
    Thoughts = m.strThoughts
End Property
Public Property Let Thoughts(ByVal strThoughts As String)
    m.strThoughts = strThoughts
End Property

Public Property Get EmotionNumber() As Long
    EmotionNumber = m.lEmotionNumber
End Property
Public Property Let EmotionNumber(ByVal lEmotionNumber As Long)
    m.lEmotionNumber = lEmotionNumber
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
End Property

Public Property Get JournalDate() As Double
    JournalDate = m.dJournalDate
End Property
Public Property Let JournalDate(ByVal dJournalDate As Double)
    m.dJournalDate = dJournalDate
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Private Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Private Property Let Symbol(ByVal strSymbol As String)
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
    m.lSymbolID = g.TnCore.GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = g.TnCore.GetSymbol(vSymbolOrSymbolID)
End Property

Public Property Get JournalImage(ByVal nImageType As eGDJournalImageTypes) As cJournalImage
On Error GoTo ErrSection:

    If m.JournalImages.Exists(Str(nImageType)) Then
        Set JournalImage = m.JournalImages(Str(nImageType))
    Else
        Set JournalImage = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cJournal.JournalImage.Get"
    
End Property
Public Property Let JournalImage(ByVal nImageType As eGDJournalImageTypes, ImageToSet As cJournalImage)
On Error GoTo ErrSection:

    If ImageToSet Is Nothing Then
        If m.JournalImages.Exists(Str(nImageType)) Then
            m.JournalImages.Remove Str(nImageType)
        End If
    ElseIf Len(ImageToSet.FileName) = 0 Then
        If m.JournalImages.Exists(Str(nImageType)) Then
            m.JournalImages.Remove Str(nImageType)
        End If
    Else
        If m.JournalImages.Exists(Str(nImageType)) Then
            Set m.JournalImages(Str(nImageType)) = ImageToSet
        Else
            m.JournalImages.Add ImageToSet, Str(nImageType)
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cJournal.JournalImage.Let"
    
End Property

Public Property Get JournalImages() As cGdTree
    Set JournalImages = m.JournalImages
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayString
'' Description: Build a display string out of the properties
'' Inputs:      None
'' Returns:     Display string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DisplayString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of information to return
    
    Set astrReturn = New cGdArray
    If m.strAction = "Reversal" Then
        astrReturn.Add "This order was a " & m.strAction
    ElseIf Len(m.strAction) > 0 Then
        astrReturn.Add "This order was an " & m.strAction
    End If
    
    If (m.lEmotionNumber >= 0) And (Len(m.strFeelings) > 0) Then
        astrReturn.Add "You were feeling: " & Str(m.lEmotionNumber) & " ; " & m.strFeelings
    ElseIf (m.lEmotionNumber >= 0) Then
        astrReturn.Add "You were feeling: " & Str(m.lEmotionNumber)
    ElseIf (Len(m.strFeelings) > 0) Then
        astrReturn.Add "You were feeling: " & m.strFeelings
    End If
    
    If Len(m.strWhyTrade) > 0 Then
        astrReturn.Add "You made this trade because: " & m.strWhyTrade
    End If
    
    If Len(m.strThoughts) > 0 Then
        astrReturn.Add "You were thinking: " & m.strThoughts
    End If
    
    If Len(m.strNote) > 0 Then
        astrReturn.Add "Notes: " & m.strNote
    End If
    
    DisplayString = astrReturn.JoinFields(vbLf)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournal.DisplayString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasValidJournalImages
'' Description: Does this journal entry have valid images?
'' Inputs:      None
'' Returns:     True if has valid images, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasValidJournalImages() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim JournalImage As cJournalImage   ' Journal image
    
    bReturn = False
    For lIndex = 1 To m.JournalImages.Count
        Set JournalImage = m.JournalImages(lIndex)
        If Len(JournalImage.FileName) > 0 Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    HasValidJournalImages = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournal.HasValidJournalImages"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.JournalImages = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournal.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.JournalImages = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournal.Class_Terminate"
    
End Sub
