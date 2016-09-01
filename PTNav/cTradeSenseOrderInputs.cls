VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrderInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrderInputs.cls
'' Description: Class for a collection of TradeSense order inputs
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/11/2010   DAJ         Fixed DefaultValueForName -- set to value, not name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Inputs As cGdTree                   ' Collection of inputs
    bForGroups As Boolean               ' Is this collection for a group?
End Type
Private m As mPrivate

Public Property Get ForGroups() As Boolean
    ForGroups = m.bForGroups
End Property
Public Property Let ForGroups(ByVal bForGroups As Boolean)
    m.bForGroups = bForGroups
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As cTradeSenseOrderInput
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrSection:

    If Exists(vIndexOrKey) Then
        Set Item = m.Inputs(vIndexOrKey)
    Else
        Set Item = Nothing
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.Item.Get"
    
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, ByVal tsInput As cTradeSenseOrderInput)
On Error GoTo ErrSection:

    If Exists(vIndexOrKey) Then
        Set m.Inputs(vIndexOrKey) = tsInput
    Else
        Add tsInput
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.Item.Let"
    
End Property

Public Function Add(ByVal tsInput As cTradeSenseOrderInput) As Long
    Add = m.Inputs.Add(tsInput, tsInput.Key(m.bForGroups))
End Function

Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    Remove = m.Inputs.Remove(vIndexOrKey)
End Function

Public Function Exists(ByVal vIndexOrKey As Variant) As Boolean
    Exists = m.Inputs.Exists(vIndexOrKey)
End Function

Public Function Count() As Long
    Count = m.Inputs.Count
End Function

Public Sub Clear()
    m.Inputs.Clear
End Sub

Public Property Get DefaultValueForName(ByVal strName As String) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    For lIndex = 1 To Count
        If UCase(Item(lIndex).Name) = UCase(strName) Then
            strReturn = Item(lIndex).DefaultValue
            Exit For
        End If
    Next lIndex
    
    DefaultValueForName = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.DefaultValueForName.Get"
    
End Property
Public Property Let DefaultValueForName(ByVal strName As String, ByVal strDefaultValue As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        If UCase(Item(lIndex).Name) = UCase(strName) Then
            Item(lIndex).DefaultValue = strDefaultValue
        End If
    Next lIndex

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.DefaultValueForName.Let"
    
End Property

Public Property Get ValueForName(ByVal strName As String) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    For lIndex = 1 To Count
        If UCase(Item(lIndex).Name) = UCase(strName) Then
            strReturn = Item(lIndex).Value
            Exit For
        End If
    Next lIndex
    
    ValueForName = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.ValueForName.Get"
    
End Property
Public Property Let ValueForName(ByVal strName As String, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Count
        If UCase(Item(lIndex).Name) = UCase(strName) Then
            Item(lIndex).Value = strValue
        End If
    Next lIndex

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.ValueForName.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Build the tree from a delimited string
'' Inputs:      Inputs String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strInputs As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrInputs As cGdArray          ' Array of inputs
    Dim tsInput As cTradeSenseOrderInput ' Trade Sense order input object
    
    Set astrInputs = New cGdArray
    astrInputs.SplitFields strInputs, "|"
    
    m.Inputs.Clear
    For lIndex = 0 To astrInputs.Size - 1
        Set tsInput = New cTradeSenseOrderInput
        tsInput.FromString astrInputs(lIndex)
        
        Add tsInput
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a delimited string from the collection
'' Inputs:      None
'' Returns:     Inputs String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrInputs As cGdArray          ' Array of inputs
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrInputs = New cGdArray
    astrInputs.Create eGDARRAY_Strings
    
    For lIndex = 1 To Count
        astrInputs.Add Item(lIndex).ToString
    Next lIndex
    
    ToString = astrInputs.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateKey
'' Description: Update the key if necessary
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateKey(ByVal lIndex As Long)
On Error GoTo ErrSection:

    m.Inputs.Key(lIndex) = Item(lIndex).Key(m.bForGroups)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.UpdateKey"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order input object
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTradeSenseOrderInputs
On Error GoTo ErrSection:

    Dim aCopy As New cTradeSenseOrderInputs
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the input object
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsInput As cTradeSenseOrderInput ' Order input object

    m = mCopyFrom
    
    Set m.Inputs = New cGdTree
    For lIndex = 1 To mCopyFrom.Inputs.Count
        Set tsInput = mCopyFrom.Inputs(lIndex).MakeCopy
        Add tsInput
    Next lIndex
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.SetPrivateData"
    
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

    Set m.Inputs = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.Class_Initialize"
    
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

    Set m.Inputs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInputs.Class_Terminate"
    
End Sub

