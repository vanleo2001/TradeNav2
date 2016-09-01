VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMostRecentlyUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cMostRecentlyUsed.cls
'' Description: Class for handling a list of most recently used items
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 10/13/2011   DAJ         Passing in INI file for MRU persistence
'' 01/25/2012   DAJ         Added INI file for defaults
'' 08/08/2013   DAJ         Optional different key to lookup default
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    astrList As cGdArray                ' List of most recently used items
    
    strKey As String                    ' Key into the INI file
    strDefaultKey As String             ' Key into the defaults INI file
    lMaxItems As Long                   ' Maximum number of items for the list
    lCount As Long                      ' Count of non-blank items
    strIniFile As String                ' INI file for persistence
    strDefaultsIni As String            ' INI file for defaults
End Type
Private m As mPrivate

Public Property Get RecentlyUsedList() As cGdArray
    Set RecentlyUsedList = m.astrList
End Property

Public Property Get Count() As Long
    Count = m.lCount
End Property

Public Property Get LastUsed() As String
    LastUsed = m.astrList(0)
End Property
Public Property Let LastUsed(ByVal strLastUsed As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndexForLast As Long           ' Index in the array of the last used string
    Dim lStartMove As Long              ' Index to start move
    
    If Count > 0 Then
        lIndexForLast = IndexForItem(strLastUsed)
        
        If lIndexForLast > -1& Then
            lStartMove = lIndexForLast
        ElseIf Count < m.lMaxItems Then
            lStartMove = Count
        Else
            lStartMove = m.lMaxItems - 1
        End If
        
        For lIndex = lStartMove To 1 Step -1
            m.astrList(lIndex) = m.astrList(lIndex - 1)
        Next lIndex
    End If
    
    m.astrList(0) = strLastUsed
    
    If Count < m.lMaxItems Then
        m.lCount = Count + 1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cMostRecentlyUsed.LastUsed"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the object
'' Inputs:      Key, Max number of items, Ini File, Defaults Ini, Defaults Key
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(ByVal strKey As String, ByVal lMaxItems As Long, ByVal strIniFile As String, Optional ByVal strDefaultsIni As String = "", Optional ByVal strDefaultsKey As String = "")
On Error GoTo ErrSection:

    m.strKey = strKey
    m.lMaxItems = lMaxItems
    m.strIniFile = strIniFile
    m.strDefaultsIni = strDefaultsIni
    m.strDefaultKey = strDefaultsKey
    
    Set m.astrList = New cGdArray
    m.astrList.Create eGDARRAY_Strings, lMaxItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load items for the given key out of the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    LoadFromIniFile m.strIniFile, False
    If (m.lCount = 0) And (Len(m.strDefaultsIni) > 0) Then
        LoadFromIniFile m.strDefaultsIni, True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save items for the given key to the INI file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    For lIndex = 1 To m.lMaxItems
        SetIniFileProperty m.strKey & Str(lIndex), m.astrList(lIndex - 1), "MRU", m.strIniFile
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexForItem
'' Description: Determine the index in the list for the given item
'' Inputs:      Item
'' Returns:     Index in List (-1 if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IndexForItem(ByVal strItem As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    For lIndex = 0 To Count - 1
        If m.astrList(lIndex) = strItem Then
            lReturn = lIndex
            Exit For
        End If
    Next lIndex
    
    IndexForItem = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cMostRecentlyUsed.IndexForItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFromIniFile
'' Description: Load the list from the given Ini File
'' Inputs:      Ini File, Defaults?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFromIniFile(ByVal strIniFile As String, ByVal bDefaults As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strItem As String               ' Item from the ini file
    Dim strKey As String                ' Key to use
    
    If (bDefaults = True) And (Len(m.strDefaultKey) > 0) Then
        strKey = m.strDefaultKey
    Else
        strKey = m.strKey
    End If
    
    m.lCount = 0
    For lIndex = 1 To m.lMaxItems
        strItem = GetIniFileProperty(strKey & Str(lIndex), "", "MRU", strIniFile)
        If Len(strItem) > 0 Then
            m.astrList(lIndex - 1) = strItem
            m.lCount = lIndex
        Else
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.LoadFromIniFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Do any initialization that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Do any clean up that needs to be done for the class
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrList = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cMostRecentlyUsed.Class_Terminate"
    
End Sub
