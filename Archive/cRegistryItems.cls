VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistryItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRegistryItems.cls
'' Description: Class to handle a collection of registry items
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
'' 02/19/2009   DAJ         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    RegistryItems As cGdTree            ' Collection of registry items
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cRegistryItem
    If m.RegistryItems.Exists(vIndexOrKey) = True Then
        Set Item = m.RegistryItems(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromIniFile
'' Description: Get registry item information from the settings file
'' Inputs:      Settings File Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromIniFile(ByVal strIniFile As String)
On Error GoTo ErrSection:

    Dim lNumRegKeys As Long             ' Number of registry keys
    Dim lIndex As Long                  ' Index into a for loop
    Dim RegItem As cRegistryItem        ' Registry Item
    Dim strSection As String            ' Section in the settings.ini file
    Dim strDefault As String            ' Default value out of the ini file
    
    lNumRegKeys = GetIniFileProperty("RegKeys", 0&, "Main", strIniFile)
    
    For lIndex = 1 To lNumRegKeys
        strSection = "RegKey" & Str(lIndex)
        
        Set RegItem = New cRegistryItem
        
        RegItem.Root = GetIniFileProperty("Root", rkLocalMachine, strSection, strIniFile)
        RegItem.KeyName = GetIniFileProperty("KeyName", "", strSection, strIniFile)
        RegItem.ValueName = GetIniFileProperty("ValueName", "", strSection, strIniFile)
        RegItem.DataType = GetIniFileProperty("Type", "", strSection, strIniFile)
        
        Select Case UCase(RegItem.DataType)
            Case "STRING"
                strDefault = GetIniFileProperty("Default", "", strSection, strIniFile)
                RegItem.DefaultValue = strDefault
            Case "LONG"
                strDefault = GetIniFileProperty("Default", "", strSection, strIniFile)
                RegItem.DefaultValue = CLng(Val(strDefault))
            Case "DOUBLE"
                strDefault = GetIniFileProperty("Default", "", strSection, strIniFile)
                RegItem.DefaultValue = Val(strDefault)
        End Select
        
        RegItem.FromRegistry
        
        m.RegistryItems.Add RegItem, RegItem.Key
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItems.FromIniFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromRegFile
'' Description: Get registry item information from the registry file
'' Inputs:      Registry File Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromRegFile(ByVal strRegFile As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array of information from the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim RegItem As cRegistryItem        ' Registry Item
    
    Set astrFile = New cGdArray
    astrFile.FromFile strRegFile
    
    For lIndex = 0 To astrFile.Size - 1
        Set RegItem = New cRegistryItem
        
        RegItem.FromString astrFile(lIndex)
        RegItem.ToRegistry
        
        m.RegistryItems.Add RegItem, RegItem.Key
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItems.FromRegFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToRegFile
'' Description: Create registry file from the registry items in collection
'' Inputs:      Registry File Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToRegFile(ByVal strRegFile As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' Array of information from the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim RegItem As cRegistryItem        ' Registry Item
    
    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.RegistryItems.Count
        astrFile.Add Item(lIndex).ToString
    Next lIndex
    
    astrFile.ToFile strRegFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItems.ToRegFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToRegistry
'' Description: Save the entire collection to the registry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToRegistry()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.RegistryItems.Count
        Item(lIndex).ToRegistry
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRegistryItems.ToRegistry"
    
End Function

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

    Set m.RegistryItems = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItems.Class_Initialize"

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

    Set m.RegistryItems = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItems.Class_Terminate"

End Sub
