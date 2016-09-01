VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemSecuritys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Type mPrivate
    Securities As Collection
End Type
Private m As mPrivate

Private Const eClass = "Class: cSystemSecurities"

Public Sub Add(SystemSecurity As cSystemSecurity, pKey As String)
    m.Securities.Add SystemSecurity, pKey
End Sub

Property Get Found(ByVal pKey As Variant) As Boolean
On Error GoTo ErrSection:
    
    Dim SystemSecurity As cSystemSecurity
    
    Found = True
    Set SystemSecurity = m.Securities(pKey)

ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Or Err.Number = 9 Then
        Found = False
        Resume Next
    Else
        RaiseError "cSystemSecuritys.Found.Get", eGDRaiseError_Raise
    End If

End Property

Property Get Item(ByVal pKey As Variant) As cSystemSecurity
    Set Item = m.Securities(pKey)
End Property

Property Get Count() As Long
    Count = m.Securities.Count
End Property

Public Sub Remove(ByVal pRuleID As Long)
On Error GoTo ErrSection:
    
    Dim X       As Integer
    
    'Remove rule's securities
    For X = Count To 1 Step -1
        If Item(X).RuleID = pRuleID Then
            m.Securities.Remove CStr(Item(X).ParmID)
        End If
    Next X
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSystemSecuritys.Remove", eGDRaiseError_Raise

End Sub

Public Sub Validate()
On Error GoTo ErrSection:
    
    Dim X           As Integer
    Dim rs          As Recordset
    Dim strSymbol   As String
    
    If Count = 0 Then
        Err.Raise vbObjectError + 1000, , _
            "At least one security must exist for a strategy."
    End If
    
    For X = 1 To Count
        With Item(X)
            
            'Error - Symbol not provided...
            If Len(Trim(.MarketSymbol)) = 0 Or Len(Trim(.Symbol)) = 0 Then
                Err.Raise vbObjectError + 1000, , "Must supply market information for " & .ParmName
            End If
    
            'Assign symbol (change for indexes and stocks)...
            Select Case .SecurityType
                Case "I": strSymbol = "$"
                Case "S": strSymbol = "!"
                Case Else
                    strSymbol = .MarketSymbol
            End Select
    
            If .Path = "Master" Or .Path = "\Qmaster" Then
                Err.Raise vbObjectError + 1000, , _
                    "Data location for " & Trim(UCase(.SecurityName)) & _
                        " not specified.  Please go to the DATA tab and " & _
                        "assign data to the security inputs."
            End If
    
        End With
    Next X
         
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cSystemSecuritys.Validate", eGDRaiseError_Raise

End Sub

Private Sub Class_Initialize()
    Set m.Securities = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.Securities = Nothing
End Sub
