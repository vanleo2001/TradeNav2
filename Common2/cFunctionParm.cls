VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctionParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type mPrivate
    lFunctionID As Long
    lParmID As Long
    bParmNumber As Byte
    strParmText As String
    lParmTypeID As Long
    strDefaultValue As String
    dFromValue As Double
    dToValue As Double
    bRequired As Boolean
    lListID As Long
    strFillPre As String
    strFillPost As String
    Lists As cLists
    ParmTypes As cParmTypes
    strError As String
End Type
Private m As mPrivate

'Error processing...
Const errParmText = vbObjectError + 1000
Const errParmType = vbObjectError + 1001
Const errList = vbObjectError + 1002
Const errFillWords = vbObjectError + 1003
Const errDefaultValue = vbObjectError + 1004

'Accept a reference to supporting tables (These are defined in cSystemsBridge)
Property Get ErrDesc() As String
    ErrDesc = m.strError
End Property
Property Let ParmTypesRef(pData As cParmTypes)
    Set m.ParmTypes = pData
End Property
Property Let ListsRef(pData As cLists)
    Set m.Lists = pData
End Property

Property Get FunctionID() As Long
    FunctionID = m.lFunctionID
End Property
Property Get ParmID() As Long
    ParmID = m.lParmID
End Property
Property Get ParmNbr() As Byte
    ParmNbr = m.bParmNumber
End Property
Property Get ParmText() As String
    ParmText = m.strParmText
End Property
Property Get ParmTypeID() As Long
    ParmTypeID = m.lParmTypeID
End Property
Property Get DefaultValue() As String
    DefaultValue = m.strDefaultValue
End Property
Property Get FromValue() As Double
    FromValue = m.dFromValue
End Property
Property Get ToValue() As Double
    ToValue = m.dToValue
End Property
Property Get Required() As Boolean
    Required = m.bRequired
End Property
Property Get ListID() As Long
    ListID = m.lListID
End Property
Property Get FillPre() As String
    FillPre = m.strFillPre
End Property
Property Get FillPost() As String
    FillPost = m.strFillPost
End Property


Property Let FunctionID(pData As Long)
    m.lFunctionID = pData
End Property
Property Let ParmID(pData As Long)
    m.lParmID = pData
End Property
Property Let ParmNbr(pData As Byte)
    m.bParmNumber = pData
End Property
Property Let ParmText(pData As String)
    m.strParmText = pData
End Property
Property Let ParmTypeID(pData As Long)
    m.lParmTypeID = pData
End Property
Property Let DefaultValue(pData As String)
    m.strDefaultValue = pData
End Property
Property Let FromValue(pData As Double)
    m.dFromValue = pData
End Property
Property Let ToValue(pData As Double)
    m.dToValue = pData
End Property
Property Let Required(pData As Boolean)
    m.bRequired = pData
End Property
Property Let ListID(pData As Long)
    m.lListID = pData
End Property
Property Let FillPre(pData As String)
    m.strFillPre = pData
End Property
Property Let FillPost(pData As String)
    m.strFillPost = pData
End Property

Public Sub Validate()
On Error GoTo ErrSection:

    'ParmText validation
    If Len(m.strParmText) = 0 Then
        m.strError = "Parameter text must be entered."
        Err.Raise errParmText, "cFunctionParm", m.strError
    End If
    If Len(m.strParmText) > 50 Then
        m.strError = "Parameter text cannot exceed 50 characters"
        Err.Raise errParmText, "cFunctionParm", m.strError
    End If

    'Parm type validation
    If Not m.ParmTypes.Found(m.lParmTypeID) Then
        m.strError = "Please select a Parameter type from the list"
        Err.Raise errParmType, "cFunctionParm", m.strError
    End If

    'List ID validation
    If m.lParmTypeID = 2 And ListID <> 0 Then
        If Not m.Lists.Found(m.lListID) Then
            m.strError = "Please select a List name"
            Err.Raise errList, "cFunctionParm", m.strError
        End If
    End If

    'Default value validation
    Select Case m.lParmTypeID
        Case 1, 4       'Single/Series Numeric values expected
            If IsNumeric(m.strDefaultValue) Then
                If m.dFromValue <> 0 And m.dToValue <> 0 Then
                    If Val(m.strDefaultValue) < m.dFromValue Or _
                       Val(m.strDefaultValue) > m.dToValue Then
                        m.strError = "The default value must be within the From/To range"
                        Err.Raise errDefaultValue, "cFunctionParm", m.strError
                    End If
                End If
            End If
        Case 5          'Price bars
            m.strDefaultValue = "Market1"
        
        Case 6, 3     'single True/False value
            If IsNumeric(m.strDefaultValue) Then
                If Val(m.strDefaultValue) <> 0 And Val(m.strDefaultValue) <> -1 Then
                    m.strError = "Please select a value of TRUE or FALSE"
                    Err.Raise errDefaultValue, "cFunctionParm", m.strError
                End If
            End If
        
        Case 7          'Trades
            m.strDefaultValue = "Trades"
            
    End Select
    
    'FillPre must be 30 or less characters
    If Len(m.strFillPre) > 30 Then
        m.strError = "Fill words before value must be 30 characters or less"
        Err.Raise errFillWords, "cFunctionParm", m.strError
    End If
    
    'FillPost must be 30 or less characters
    If Len(m.strFillPost) > 30 Then
        m.strError = "Fill words after value must be 30 characters or less"
        Err.Raise errParmText, "cFunctionParm", m.strError
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFunctionParm.Validate", eGDRaiseError_Raise, g.strAppPath

End Sub
