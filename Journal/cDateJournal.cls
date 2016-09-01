VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cDateJournal.cls
'' Description: Class that contains information about a date journal
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/22/2011   DAJ         Added filename and caption for attached chart image
'' 09/23/2011   DAJ         Delete image file when date journal is deleted
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the date journal object
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lDateJournalID As Long              ' Unique identifier for the date journal entry
    dJournalDate As Double              ' Date of the journal entry
    lJournalCategoryID As Long          ' Category for the journal entry
    dJournalTime As Double              ' Time of the journal entry
    strText As String                   ' Text of the journal entry
    lSymbolID As Long                   ' Symbol ID for the journal
    strSymbol As String                 ' Symbol for the journal
    JournalImages As cGdTree            ' Collection of journal images
End Type
Private m As mPrivate

Public Property Get DateJournalID() As Long
    DateJournalID = m.lDateJournalID
End Property
Public Property Let DateJournalID(ByVal lDateJournalID As Long)
    m.lDateJournalID = lDateJournalID
End Property

Public Property Get JournalDate() As Double
    JournalDate = m.dJournalDate
End Property
Public Property Let JournalDate(ByVal dJournalDate As Double)
    m.dJournalDate = dJournalDate
End Property

Public Property Get JournalCategoryID() As Long
    JournalCategoryID = m.lJournalCategoryID
End Property
Public Property Let JournalCategoryID(ByVal lJournalCategoryID As Long)
    m.lJournalCategoryID = lJournalCategoryID
End Property

Public Property Get JournalTime() As Double
    JournalTime = m.dJournalTime
End Property
Public Property Let JournalTime(ByVal dJournalTime As Double)
    m.dJournalTime = dJournalTime
End Property

Public Property Get Text() As String
    Text = m.strText
End Property
Public Property Let Text(ByVal strText As String)
    m.strText = strText
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
    g.TnCore.RaiseError "cDateJournal.HasValidJournalImages"
    
End Function

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

    DateJournalID = 0&
    JournalDate = 0#
    JournalCategoryID = 0&
    JournalTime = 0#
    Text = ""
    SymbolID = 0&
    Symbol = ""
    
    Set m.JournalImages = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournal.Class_Initialize"
    
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

    Set m.JournalImages = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournal.Class_Terminate"
    
End Sub
