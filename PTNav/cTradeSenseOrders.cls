VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrders.cls
'' Description: Class that handles a collection of Trade Sense orders
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/16/2010   DAJ         Allow for adding same order multiple times (#5800)
'' 06/17/2010   DAJ         Changed filenames to ID instead of name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    tsOrders As cGdTree                 ' Collection of Trade Sense orders
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As cTradeSenseOrder
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object

    If Exists(vIndexOrKey) Then
        Set tsOrder = m.tsOrders(vIndexOrKey)
    Else
        Set tsOrder = Nothing
    End If
    
    Set Item = tsOrder

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrders.Item.Get"

End Property

Public Function Count() As Long
    Count = m.tsOrders.Count
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.tsOrders.Exists(vIndexOrKey)
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.tsOrders.Remove(vIndexOrKey)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given order to the collection
'' Inputs:      Order, Key
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(tsOrder As cTradeSenseOrder, Optional ByVal strKey As String = "") As Long
On Error GoTo ErrSection:

    If Len(strKey) = 0 Then
        m.tsOrders.Add tsOrder, tsOrder.ID
    Else
        m.tsOrders.Add tsOrder, strKey
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrders.Add"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the collection of orders
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:

    Dim astrFiles As New cGdArray       ' Array of matching files
    
    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Provided\*.TSO", True, True
    LoadFiles astrFiles

    astrFiles.GetMatchingFiles AddSlash(App.Path) & "Custom\*.TSO", True, True
    LoadFiles astrFiles

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrders.Load"
    
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
    For lIndex = 1 To m.tsOrders.Count
        If Item(lIndex).Name = strName Then
            strReturn = Item(lIndex).ID
            Exit For
        End If
    Next lIndex
    
    IdForName = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrders.IdForName"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFiles
'' Description: Load up the orders in the given array of files
'' Inputs:      Files
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFiles(ByVal astrFiles As cGdArray)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade sense order object

    For lIndex = 0 To astrFiles.Size - 1
        Set tsOrder = New cTradeSenseOrder
        tsOrder.FromFile astrFiles(lIndex)
        
        Add tsOrder
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrders.LoadFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.tsOrders = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrders.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.tsOrders = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrders.Class_Terminate"
    
End Sub
