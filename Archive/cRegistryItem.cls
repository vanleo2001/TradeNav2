VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cRegistryItem.cls
'' Description: Class to handle a registry item
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
'' 02/17/2009   DAJ         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    nRoot As RKeyRoots                  ' Registry root
    strKeyName As String                ' Key name
    strValueName As String              ' Value name
    strDataType As String               ' Data type
    vDefaultValue As Variant            ' Default value
    vValue As Variant                   ' Value
End Type
Private m As mPrivate

Public Property Get Root() As RKeyRoots
    Root = m.nRoot
End Property
Public Property Let Root(ByVal nRoot As RKeyRoots)
    m.nRoot = nRoot
End Property

Public Property Get KeyName() As String
    KeyName = m.strKeyName
End Property
Public Property Let KeyName(ByVal strKeyName As String)
    m.strKeyName = strKeyName
End Property

Public Property Get ValueName() As String
    ValueName = m.strValueName
End Property
Public Property Let ValueName(ByVal strValueName As String)
    m.strValueName = strValueName
End Property

Public Property Get DataType() As String
    DataType = m.strDataType
End Property
Public Property Let DataType(ByVal strDataType As String)
    m.strDataType = strDataType
End Property

Public Property Get DefaultValue() As Variant
    DefaultValue = m.vDefaultValue
End Property
Public Property Let DefaultValue(ByVal vDefaultValue As Variant)
    m.vDefaultValue = vDefaultValue
End Property

Public Property Get Value() As Variant
    Value = m.vValue
End Property
Public Property Let Value(ByVal vValue As Variant)
    m.vValue = vValue
End Property

Public Property Get Key() As String
    Key = Str(m.nRoot) & "\" & m.strKeyName & "\" & m.strValueName
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Set the properties from the given string
'' Inputs:      String
'' Returns:     None
''
'' Fields:      Root, KeyName, ValueName, Data Type, Default Value, Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strRegistryItemString As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields broken out into an array
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strRegistryItemString, vbTab
    
    m.nRoot = CLng(Val(astrFields(0)))
    m.strKeyName = astrFields(1)
    m.strValueName = astrFields(2)
    m.strDataType = astrFields(3)
    
    Select Case UCase(m.strDataType)
        Case "STRING"
            m.vDefaultValue = astrFields(4)
            m.vValue = astrFields(5)
        Case "LONG"
            m.vDefaultValue = CLng(Val(astrFields(4)))
            m.vValue = CLng(Val(astrFields(5)))
        Case "DOUBLE"
            m.vDefaultValue = Val(astrFields(4))
            m.vValue = Val(astrFields(5))
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItem.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Return a string put together from the properties
'' Inputs:      None
'' Returns:     String
''
'' Fields:      Root, KeyName, ValueName, Data Type, Default Value, Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields broken out into an array
    
    Set astrFields = New cGdArray
    astrFields.Create eGDARRAY_Strings
    
    astrFields(0) = Str(m.nRoot)
    astrFields(1) = m.strKeyName
    astrFields(2) = m.strValueName
    astrFields(3) = m.strDataType
    astrFields(4) = Str(m.vDefaultValue)
    astrFields(5) = Str(m.vValue)
    
    ToString = astrFields.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRegistryItem.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromRegistry
'' Description: Get the value from the registry
'' Inputs:      None
'' Returns:     Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromRegistry() As Variant
On Error GoTo ErrSection:

    m.vValue = GetRegistryValue(m.nRoot, m.strKeyName, m.strValueName, m.vDefaultValue)
    FromRegistry = m.vValue

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cRegistryItem.FromRegistry"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToRegistry
'' Description: Set the value in the registry
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToRegistry()
On Error GoTo ErrSection:

    SetRegistryValue m.nRoot, m.strKeyName, m.strValueName, m.vValue

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cRegistryItem.ToRegistry"

End Sub
