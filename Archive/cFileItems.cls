VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFileItems.cls
'' Description: Class to handle a collection of file items
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/19/2009   DAJ         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    FileItems As New cGdTree            ' Collection of file items
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.FileItems.Count
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cFileItem
Attribute Item.VB_UserMemId = 0
    If m.FileItems.Exists(vIndexOrKey) = True Then
        Set Item = m.FileItems.Item(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromIniFile
'' Description: Load up the file items from the Ini file
'' Inputs:      Ini File Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromIniFile(ByVal strIniFile As String)
On Error GoTo ErrSection:

    Dim lNumFileItems As Long           ' Number of file items
    Dim lIndex As Long                  ' Index into a for loop
    Dim strSection As String            ' Section in the ini file
    Dim FileItem As cFileItem           ' File item object
    
    lNumFileItems = GetIniFileProperty("NumGroups", 0&, "Main", strIniFile)
    
    For lIndex = 1 To lNumFileItems
        strSection = "Group" & Str(lIndex)
        
        Set FileItem = New cFileItem
        
        FileItem.GroupName = GetIniFileProperty("GroupName", "", strSection, strIniFile)
        FileItem.ZipName = GetIniFileProperty("ZipName", "", strSection, strIniFile)
        FileItem.UnzippedPath = GetIniFileProperty("UnzippedPath", "", strSection, strIniFile)
        FileItem.FileMasks = GetIniFileProperty("FileMasks", "", strSection, strIniFile)
        FileItem.DelFirst = GetIniFileProperty("DelFirst", 0&, strSection, strIniFile)
        FileItem.OverwriteMode = GetIniFileProperty("OverwriteMode", 0&, strSection, strIniFile)
        
        m.FileItems.Add FileItem, FileItem.ZipName
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItems.FromIniFile"
    
End Sub

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

    Set m.FileItems = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItems.Class_Initialize"

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

    Set m.FileItems = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItems.Class_Terminate"
    
End Sub
