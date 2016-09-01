VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalCategories.cls
'' Description: Class that contains a collection of journal categories
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 09/30/2011   DAJ         Exposed the exists function
'' 07/30/2013   DAJ         Added IdForName lookup
'' 08/08/2013   DAJ         Added TypeForId lookup
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Categories As cGdTree               ' Collection of journal categories
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.Categories.Count
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cJournalCategory
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    If m.Categories.Exists(vIndexOrKey) Then
        Set Item = m.Categories(vIndexOrKey)
    Else
        Set Item = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.Item.Get"

End Property

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Categories.Exists(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the collection from the given file
'' Inputs:      Filename
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array of lines from the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim journalCategory As cJournalCategory ' Journal category object
    Dim lFileVersion As Long            ' Version of the file
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strFileName) Then
        lFileVersion = CLng(Val(astrFile(0)))
        
        For lIndex = 1 To astrFile.Size - 1
            Set journalCategory = New cJournalCategory
            journalCategory.FromString astrFile(lIndex)
            m.Categories.Add journalCategory, Str(journalCategory.ID)
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IdForName
'' Description: Determine the category ID for the given name
'' Inputs:      Name
'' Returns:     ID ( Zero if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IdForName(ByVal strCategoryName As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Category As cJournalCategory    ' Journal category
    
    lReturn = 0&
    For lIndex = 1 To m.Categories.Count
        Set Category = Item(lIndex)
        If Category.Text = strCategoryName Then
            lReturn = Category.ID
            Exit For
        End If
    Next lIndex
    
    IdForName = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.IdForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TypeForId
'' Description: Determine the category type for the given ID
'' Inputs:      Category ID
'' Returns:     Category Type ( Note if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TypeForId(ByVal lCategoryID As Long) As eGDJournalCategoryTypes
On Error GoTo ErrSection:

    Dim nReturn As eGDJournalCategoryTypes ' Return value for the function
    
    nReturn = eGDJournalCategoryType_Note
    If Exists(Str(lCategoryID)) Then
        nReturn = Item(Str(lCategoryID)).CategoryType
    End If
    
    TypeForId = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.TypeForId"
    
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

    Set m.Categories = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.Class_Initialize"
    
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

    Set m.Categories = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategories.Class_Terminate"
    
End Sub

