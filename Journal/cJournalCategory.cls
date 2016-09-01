VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalCategory.cls
'' Description: Class that contains information about a journal category
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/25/2012   DAJ         Added required module field
'' 08/08/2013   DAJ         Journal category type
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Category ID
    strText As String                   ' Text to display to user
    bActive As Boolean                  ' Is the category still active?
    strRequiredModule As String         ' Required module code for the category
    nType As eGDJournalCategoryTypes   ' Journal category type
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Private Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get Text() As String
    Text = m.strText
End Property
Private Property Let Text(ByVal strText As String)
    m.strText = strText
End Property

Public Property Get Active() As Boolean
    Active = m.bActive
End Property
Private Property Let Active(ByVal bActive As Boolean)
    m.bActive = bActive
End Property

Public Property Get RequiredModule() As String
    RequiredModule = m.strRequiredModule
End Property
Public Property Let RequiredModule(ByVal strRequiredModule As String)
    m.strRequiredModule = strRequiredModule
End Property

Public Property Get CategoryType() As eGDJournalCategoryTypes
    CategoryType = m.nType
End Property
Public Property Let CategoryType(ByVal nType As eGDJournalCategoryTypes)
    m.nType = nType
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Set the class members from the given string
'' Inputs:      String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strText As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields from the string
    Dim lActive As Long                 ' Is the category active?
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strText, vbTab
    
    lActive = CLng(Val(astrFields(2)))
    
    ID = CLng(Val(astrFields(0)))
    Text = astrFields(1)
    Active = (lActive > 0) And (lActive <= 2)
    RequiredModule = astrFields(3)
    
    If astrFields.Size > 4 Then
        CategoryType = CLng(Val(astrFields(4)))
    Else
        CategoryType = eGDJournalCategoryType_Note
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategory.FromString"
    
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

    ID = 0
    Text = ""
    Active = False
    RequiredModule = ""
    CategoryType = eGDJournalCategoryType_Note

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategory.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalCategory.Class_Terminate"
    
End Sub

