VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBasketItemParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBasketItemParm.cls
'' Description: Wrapper class for a strategy basket item parameter override
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lStrategyBasketItemID As Long       ' ID for the strategy basket Item
    strRuleName As String               ' Name of the rule the parm belongs to ( if not linked inputs )
    strParmName As String               ' Parameter name
    strValue As String                  ' Overriden value for the strategy parameter
    bIsOverridden As Boolean            ' Is the parameter value overridden?
    bIsExposed As Boolean               ' Is the parameter exposed to the automated trading item?
End Type
Private m As mPrivate

Public Property Get StrategyBasketItemID() As Long
    StrategyBasketItemID = m.lStrategyBasketItemID
End Property
Public Property Let StrategyBasketItemID(ByVal lStrategyBasketItemID As Long)
    m.lStrategyBasketItemID = lStrategyBasketItemID
End Property

Public Property Get RuleName() As String
    RuleName = m.strRuleName
End Property
Public Property Let RuleName(ByVal strRuleName As String)
    m.strRuleName = strRuleName
End Property

Public Property Get ParmName() As String
    ParmName = m.strParmName
End Property
Public Property Let ParmName(ByVal strParmName As String)
    m.strParmName = strParmName
End Property

Public Property Get Value() As String
    Value = m.strValue
End Property
Public Property Let Value(ByVal strValue As String)
    m.strValue = strValue
End Property

Public Property Get IsOverridden() As Boolean
    IsOverridden = m.bIsOverridden
End Property
Public Property Let IsOverridden(ByVal bIsOverridden As Boolean)
    m.bIsOverridden = bIsOverridden
End Property

Public Property Get IsExposed() As Boolean
    IsExposed = m.bIsExposed
End Property
Public Property Let IsExposed(ByVal bIsExposed As Boolean)
    m.bIsExposed = bIsExposed
End Property

Public Property Get Key() As String
    If Len(m.strRuleName) > 0 Then
        Key = m.strRuleName & "." & m.strParmName
    Else
        Key = m.strParmName
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDbFromRs
'' Description: Load the strategy item paramater from the database
'' Inputs:      Strategy Basket Item ID, Rule Name, Parameter Name
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDb(ByVal lStrategyBasketItemID As Long, ByVal strRuleName As String, ByVal strParmName As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItemParms] " & _
                "WHERE [StrategyBasketItemID]=" & Str(lStrategyBasketItemID) & " AND [RuleName]='" & strRuleName & "' AND [ParmName]='" & strParmName & "';", dbOpenDynaset)
    LoadDb = LoadDbFromRs(rs)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.LoadDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDbFromRs
'' Description: Load the strategy item paramater from the database
'' Inputs:      Recordset
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDbFromRs(ByVal rs As Recordset) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        m.lStrategyBasketItemID = rs!StrategyBasketItemID
        m.strRuleName = rs!RuleName
        m.strParmName = rs!ParmName
        m.strValue = rs!Value
        m.bIsOverridden = NullChk(rs!IsOverridden, True)
        m.bIsExposed = NullChk(rs!IsExposed, False)
        
        bReturn = True
    End If
    
    LoadDbFromRs = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.LoadDbFromRs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveDb
'' Description: Save the strategy item parameter to the database
'' Inputs:      Database
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveDb(Optional ByVal DB As Database = Nothing)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    If DB Is Nothing Then
        Set DB = g.dbNav
    End If

    Set rs = DB.OpenRecordset("SELECT * FROM [tblStrategyBasketItemParms] " & _
                "WHERE [StrategyBasketItemID]=" & Str(m.lStrategyBasketItemID) & " AND [RuleName]='" & m.strRuleName & "' AND [ParmName]='" & m.strParmName & "';", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!StrategyBasketItemID = m.lStrategyBasketItemID
    rs!RuleName = m.strRuleName
    rs!ParmName = m.strParmName
    rs!Value = m.strValue
    rs!IsOverridden = m.bIsOverridden
    rs!IsExposed = m.bIsExposed
    
    rs.Update

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.SaveDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Load up the class members from the given string
'' Inputs:      Parameter Info
'' Returns:     None
''
'' Fields:      ParmName=Value ( linked inputs ) or RuleName.ParmName=Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strParm As String)
On Error GoTo ErrSection:

    Dim strKey As String                ' Key in the string
    
    strKey = Parse(strParm, "=", 1)
    
    If InStr(strKey, ".") <> 0 Then
        m.strRuleName = Parse(strKey, ".", 1)
        m.strParmName = Parse(strKey, ".", 2)
    Else
        m.strRuleName = ""
        m.strParmName = strKey
    End If
    m.strValue = Parse(strParm, "=", 2)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Build a string from the class members
'' Inputs:      Parameter Info
'' Returns:     None
''
'' Fields:      ParmName=Value ( linked inputs ) or RuleName.ParmName=Value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    ToString = Key & "=" & m.strValue

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsTheSameAs
'' Description: Is the given parm the same as this one?
'' Inputs:      Parm, Check Item ID?
'' Returns:     True if the same, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsTheSameAs(ByVal Parm As cStrategyBasketItemParm, Optional ByVal bCheckItemID As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If (bCheckItemID = True) And (m.lStrategyBasketItemID <> Parm.StrategyBasketItemID) Then
        bReturn = False
    ElseIf (Key <> Parm.Key) Then
        bReturn = False
    End If
    
    IsTheSameAs = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.IsTheSameAs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the strategy basket item parameter
'' Inputs:      None
'' Returns:     Copy of the strategy basket item parameter
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cStrategyBasketItemParm
On Error GoTo ErrSection:

    Dim aCopy As New cStrategyBasketItemParm ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    m = mCopyFrom
    m.lStrategyBasketItemID = 0&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.lStrategyBasketItemID = 0&
    m.strRuleName = ""
    m.strParmName = ""
    m.strValue = ""
    m.bIsOverridden = False
    m.bIsExposed = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParm.Class_Terminate"
    
End Sub
