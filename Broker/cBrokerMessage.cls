VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerMessage.cls
'' Description: Class to handle a key/value pair style broker message
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/20/2011   DAJ         Added Symbol property
'' 12/09/2011   DAJ         Added LegSymbol property
'' 12/15/2011   DAJ         Add more keys for which to hide the value in the log
'' 09/11/2012   DAJ         Ability to make a copy of the object
'' 10/16/2012   DAJ         Added "Password" to keys for which to hide value in log
'' 03/07/2014   DAJ         Moved into NavBroker.DLL
'' 08/22/2014   DAJ         Added new keys to the "HideValueForLog" routine for E-Trade
'' 10/24/2014   DAJ         Core Application functions for DLL's
'' 01/22/2015   DAJ         Utilize method in cGdTree to split out key value string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Fields As cGdTree                   ' Collection of fields from the message
End Type
Private m As mPrivate
    
Public Property Get Count() As Long
    Count = m.Fields.Count
End Property

Public Property Get Item(ByVal vIndexOrKey As Variant) As String
Attribute Item.VB_UserMemId = 0
    If m.Fields.Exists(vIndexOrKey) Then
        Item = m.Fields(vIndexOrKey)
    Else
        Item = ""
    End If
End Property
Public Property Let Item(ByVal vIndexOrKey As Variant, ByVal strValue As String)
    If m.Fields.Exists(vIndexOrKey) Then
        m.Fields(vIndexOrKey) = strValue
    End If
End Property
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given key/value pair
'' Inputs:      Key, Value
'' Returns:     Index into the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(ByVal strKey As String, ByVal strValue As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    If m.Fields.Exists(strKey) Then
        m.Fields(strKey) = strValue
        lReturn = m.Fields.Index(strKey)
    Else
        lReturn = m.Fields.Add(strValue, strKey)
    End If
    
    Add = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Add"
    
End Function
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Remove
'' Description: Remove the item with the given index or key
'' Inputs:      Index or Key
'' Returns:     True if removed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    If m.Fields.Exists(vIndexOrKey) Then
        bReturn = m.Fields.Remove(vIndexOrKey)
    End If
    
    Remove = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Remove"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill the collection from the given message
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strMessage As String)
On Error GoTo ErrSection:

    m.Fields.FromKeyValueString strMessage, vbTab, "="

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.FromString"
    
End Sub
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string out of the collection
'' Inputs:      For Log?
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString(Optional ByVal bForLog As Boolean) As String
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array to build into the string
    Dim lIndex As Long                  ' Index into a for loop
    Dim strValue As String              ' Value for the key

    Set astrFields = New cGdArray
    For lIndex = 1 To m.Fields.Count
        If (HideValueForLog(m.Fields.Key(lIndex)) = True) And (bForLog = True) Then
            strValue = SetStrToChar(m.Fields(lIndex), "*")
        Else
            strValue = m.Fields(lIndex)
        End If
        
        astrFields.Add m.Fields.Key(lIndex) & "=" & strValue
    Next lIndex
    
    ToString = astrFields.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.ToString"
    
End Function
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Symbol
'' Description: Extract the symbol from the fields
'' Inputs:      None
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Symbol() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrFields As cGdArray          ' Fields to make into the symbol

    If Len(Item("Symbol")) > 0 Then
        strReturn = Item("Symbol")
    Else
        Set astrFields = New cGdArray
        
        astrFields(0) = Item("BaseSymbol")
        astrFields(1) = Item("Contract")
        astrFields(2) = Item("Exchange")
        astrFields(3) = Item("SecurityType")
        astrFields(4) = Item("StrikePrice")
        astrFields(5) = Item("PutCall")
        
        strReturn = astrFields.JoinFields("|")
    End If
    
    Symbol = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Symbol.Get"
    
End Property
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Symbol
'' Description: Set the field(s) from the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Symbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields to make into the symbol
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strSymbol, "|"
    
    If astrFields.Size = 1 Then
        Add "Symbol", astrFields(0)
    Else
        Add "BaseSymbol", astrFields(0)
        Add "Contract", astrFields(1)
        Add "Exchange", astrFields(2)
        Add "SecurityType", astrFields(3)
        Add "StrikePrice", astrFields(4)
        Add "PutCall", astrFields(5)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Symbol.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LegSymbol
'' Description: Extract the symbol for the given leg from the fields
'' Inputs:      None
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LegSymbol(ByVal lLeg As Long) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrFields As cGdArray          ' Fields to make into the symbol
    Dim strLeg As String                ' String version of leg number
    
    strLeg = Str(lLeg)

    If Len(Item("Symbol" & strLeg)) > 0 Then
        strReturn = Item("Symbol" & strLeg)
    Else
        Set astrFields = New cGdArray
        
        astrFields(0) = Item("BaseSymbol" & strLeg)
        astrFields(1) = Item("Contract" & strLeg)
        astrFields(2) = Item("Exchange" & strLeg)
        astrFields(3) = Item("SecurityType" & strLeg)
        astrFields(4) = Item("StrikePrice" & strLeg)
        astrFields(5) = Item("PutCall" & strLeg)
        
        strReturn = astrFields.JoinFields("|")
    End If
    
    LegSymbol = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.LegSymbol.Get"
    
End Property
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LegSymbol
'' Description: Set the field(s) from the given symbol
'' Inputs:      Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LegSymbol(ByVal lLeg As Long, ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields to make into the symbol
    Dim strLeg As String                ' String version of leg number
    
    strLeg = Str(lLeg)
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strSymbol, "|"
    
    If astrFields.Size = 1 Then
        Add "Symbol" & strLeg, astrFields(0)
    Else
        Add "BaseSymbol" & strLeg, astrFields(0)
        Add "Contract" & strLeg, astrFields(1)
        Add "Exchange" & strLeg, astrFields(2)
        Add "SecurityType" & strLeg, astrFields(3)
        Add "StrikePrice" & strLeg, astrFields(4)
        Add "PutCall" & strLeg, astrFields(5)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.LegSymbol.Let"
    
End Property
   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the broker message
'' Inputs:      None
'' Returns:     New Copy
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cBrokerMessage
On Error GoTo ErrSection:

    Dim aCopy As New cBrokerMessage
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.MakeCopy"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private properties for a copy of the trade
'' Inputs:      Copy From
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    Set m.Fields = mCopyFrom.Fields.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideValueForLog
'' Description: Determine whether to hide the value for the given key in the log
'' Inputs:      Key
'' Returns:     True if hide value, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HideValueForLog(ByVal strKey As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean

    Select Case UCase(strKey)
        Case "APPLICATIONID", "APPLICATIONLICENSE", "CONSUMERKEY", "CONSUMERSECRET", "NEWPASSWORD", "PASS", "PASSWORD"
            bReturn = True
            
        Case Else
            bReturn = False
    End Select
    
    HideValueForLog = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.HideValueForLog"
    
End Function
   
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

    Set m.Fields = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Class_Initialize"
    
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

    Set m.Fields = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cBrokerMessage.Class_Terminate"
    
End Sub
