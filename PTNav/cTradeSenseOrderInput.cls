VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrderInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrderInput.cls
'' Description: Class for a TradeSense order input
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strName As String                   ' Name of the input
    lParmType As Long                   ' Parameter type
    lOrderNumber As Long                ' Order number that the input is for
    strDefaultValue As String           ' Default value for the input
    strValue As String                  ' Value for the input
End Type
Private m As mPrivate

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get ParmType() As Long
    ParmType = m.lParmType
End Property
Public Property Let ParmType(ByVal lParmType As Long)
    m.lParmType = lParmType
End Property

Public Property Get OrderNumber() As Long
    OrderNumber = m.lOrderNumber
End Property
Public Property Let OrderNumber(ByVal lOrderNumber As Long)
    m.lOrderNumber = lOrderNumber
End Property

Public Property Get DefaultValue() As String
    DefaultValue = m.strDefaultValue
End Property
Public Property Let DefaultValue(ByVal strDefaultValue As String)
    m.strDefaultValue = strDefaultValue
End Property

Public Property Get Value() As String
    Value = m.strValue
End Property
Public Property Let Value(ByVal strValue As String)
    m.strValue = strValue
End Property

Public Property Get Key(ByVal bForGroups As Boolean) As String
    If bForGroups Then
        Key = Name & "|" & Str(OrderNumber)
    Else
        Key = Name
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Extract the values for the input from a delimited string
'' Inputs:      Info String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strInfo As String)
On Error GoTo ErrSection:

    Dim astrInput As cGdArray           ' Input information split out into an array
    
    Set astrInput = New cGdArray
    astrInput.SplitFields strInfo, ";"
    
    Name = astrInput(0)
    ParmType = CLng(Val(astrInput(1)))
    OrderNumber = CLng(Val(astrInput(2)))
    DefaultValue = astrInput(3)
    Value = astrInput(4)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInput.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a delimited string from the member variables
'' Inputs:      None
'' Returns:     Info String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrInput As cGdArray           ' Input information split out into an array
    
    Set astrInput = New cGdArray
    astrInput.Create eGDARRAY_Strings
    
    astrInput(0) = Name
    astrInput(1) = Str(ParmType)
    astrInput(2) = Str(OrderNumber)
    astrInput(3) = DefaultValue
    astrInput(4) = Value
    
    ToString = astrInput.JoinFields(";")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderInput.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order input object
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cTradeSenseOrderInput
On Error GoTo ErrSection:

    Dim aCopy As New cTradeSenseOrderInput
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderInput.MakeCopy"
    
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

    m = mCopyFrom

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderInput.SetPrivateData"
    
End Sub
