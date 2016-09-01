VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrderGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrderGroups.cls
'' Description: Class that handles a collection of Trade Sense order groups
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
'' 08/12/2010   DAJ         Added OrderInGroup function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Groups As cGdTree                   ' Collection of TradeSense order groups
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cTradeSenseOrderGroup
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object
    
    If Exists(vIndexOrKey) Then
        Set tsoGrp = m.Groups.Item(vIndexOrKey)
    Else
        Set tsoGrp = Nothing
    End If
    
    Set Item = tsoGrp

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.Item.Get"
    
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, tsoGrp As cTradeSenseOrderGroup)
On Error GoTo ErrSection:
    
    If Exists(vIndexOrKey) Then
        Set m.Groups(vIndexOrKey) = tsoGrp
    Else
        Add tsoGrp
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.Item.Let"
    
End Property

Public Function Count() As Long
    Count = m.Groups.Count
End Function

Public Function Exists(ByVal vIndexOrKey As Variant)
    Exists = m.Groups.Exists(vIndexOrKey)
End Function

Public Function Add(ByVal tsoGrp As cTradeSenseOrderGroup)
    Add = m.Groups.Add(tsoGrp, tsoGrp.ID)
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Groups.Remove(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the Trade Sense order groups from files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    m.Groups.Clear
    LoadFiles AddSlash(App.Path) & "Provided", False
    LoadFiles AddSlash(App.Path) & "Custom", True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderInGroup
'' Description: Is the given order contained in any group?
'' Inputs:      Order ID
'' Returns:     True if order in group, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderInGroup(ByVal strOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To Count
        If Item(lIndex).ContainsOrder(strOrderID) Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    OrderInGroup = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.OrderInGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFiles
'' Description: Load up files from the given path
'' Inputs:      Path, Custom?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFiles(ByVal strPath As String, ByVal bCustom As Boolean)
On Error GoTo ErrSection:

    Dim astrFiles As New cGdArray       ' Array of matching files
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsoGrp As cTradeSenseOrderGroup ' Trade Sense order group object

    astrFiles.GetMatchingFiles AddSlash(strPath) & "*.TSG", True, True
    For lIndex = 0 To astrFiles.Size - 1
        If FileLength(astrFiles(lIndex)) < 7 Then
            KillFile astrFiles(lIndex)
        Else
            Set tsoGrp = New cTradeSenseOrderGroup
            tsoGrp.FromFile astrFiles(lIndex), bCustom
            m.Groups.Add tsoGrp, tsoGrp.Name
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.LoadFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IdForName
'' Description: Determine the ID for the given name
'' Inputs:      Name
'' Returns:     ID (Blank if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IdForName(ByVal strName As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    strReturn = ""
    For lIndex = 1 To m.Groups.Count
        If Item(lIndex).Name = strName Then
            strReturn = Item(lIndex).ID
            Exit For
        End If
    Next lIndex
    
    IdForName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.IdForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the class is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Groups = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Groups = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroups.Class_Terminate"

End Sub
