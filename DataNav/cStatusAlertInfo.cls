VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStatusAlertInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStatusAlertInfo.cls
'' Description: Class of information for status alerts
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strSymbol As String                 ' Symbol for the status alert
    strAccount As String                ' Account for the status alert
    strAutoTradeItem As String          ' Automated trading item name
    strReason As String                 ' Reason for status alert
End Type
Private m As mPrivate

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
End Property

Public Property Get Account() As String
    Account = m.strAccount
End Property
Public Property Let Account(ByVal strAccount As String)
    m.strAccount = strAccount
End Property

Public Property Get AutoTradeItem() As String
    AutoTradeItem = m.strAutoTradeItem
End Property
Public Property Let AutoTradeItem(ByVal strAutoTradeItem As String)
    m.strAutoTradeItem = strAutoTradeItem
End Property

Public Property Get Reason() As String
    Reason = m.strReason
End Property
Public Property Let Reason(ByVal strReason As String)
    m.strReason = strReason
End Property

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

    m.strSymbol = ""
    m.strAccount = ""
    m.strAutoTradeItem = ""
    m.strReason = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatusAlertInfo.Class_Initialize"
    
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
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStatusAlertInfo.Class_Terminate"
    
End Sub

