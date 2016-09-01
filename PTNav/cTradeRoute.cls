VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeRoute.cls
'' Description: Class to hold information about a trade route
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    iTradeRouteType As Integer          ' Type of the trade route
    strExchange As String               ' Exchange for the trade route
    strFcmId As String                  ' FCM Identifier for the trade route
    strIbId As String                   ' Introducing Broker ID for the trade route
    strStatus As String                 ' Status of the trade route
    strTradeRoute As String             ' Name of the trade route
    
    strKey As String                    ' Unique key
End Type
Private m As mPrivate

Public Property Get TradeRouteType() As Integer
    TradeRouteType = m.iTradeRouteType
End Property
Public Property Let TradeRouteType(ByVal iTradeRouteType As Integer)
    m.iTradeRouteType = iTradeRouteType
End Property

Public Property Get Exchange() As String
    Exchange = m.strExchange
End Property
Public Property Let Exchange(ByVal strExchange As String)
    If strExchange <> m.strExchange Then
        m.strExchange = strExchange
        BuildKey
    End If
End Property

Public Property Get FcmId() As String
    FcmId = m.strFcmId
End Property
Public Property Let FcmId(ByVal strFcmId As String)
    If strFcmId <> m.strFcmId Then
        m.strFcmId = strFcmId
        BuildKey
    End If
End Property

Public Property Get IbID() As String
    IbID = m.strIbId
End Property
Public Property Let IbID(ByVal strIbId As String)
    If strIbId <> m.strIbId Then
        m.strIbId = strIbId
        BuildKey
    End If
End Property

Public Property Get Status() As String
    Status = m.strStatus
End Property
Public Property Let Status(ByVal strStatus As String)
    m.strStatus = strStatus
End Property

Public Property Get TradeRoute() As String
    TradeRoute = m.strTradeRoute
End Property
Public Property Let TradeRoute(ByVal strTradeRoute As String)
    m.strTradeRoute = strTradeRoute
End Property

Public Property Get Key() As String
    Key = m.strKey
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromRithmicString
'' Description: Assign class members from the given Rithmic trade route string
'' Inputs:      Rithmic Trade Route
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromRithmicString(ByVal strRithmicTradeRoute As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields out of the string
    
    If Len(strRithmicTradeRoute) = 0 Then
        Clear
    Else
        Set astrFields = New cGdArray
        astrFields.SplitFields strRithmicTradeRoute, vbTab
        
        m.iTradeRouteType = Int(Val(astrFields(0)))
        m.strExchange = astrFields(1)
        m.strFcmId = astrFields(2)
        m.strIbId = astrFields(3)
        m.strStatus = astrFields(4)
        m.strTradeRoute = astrFields(5)
        
        BuildKey
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRoute.FromRithmicString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildKey
'' Description: Build a unique key for the trade route
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildKey()
On Error GoTo ErrSection:

    m.strKey = m.strExchange & vbTab & m.strFcmId & vbTab & m.strIbId

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRoute.BuildKey"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Clear
'' Description: Clear the class members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Clear()
On Error GoTo ErrSection:

    m.iTradeRouteType = 0
    m.strExchange = ""
    m.strFcmId = ""
    m.strIbId = ""
    m.strStatus = ""
    m.strTradeRoute = ""
    
    m.strKey = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRoute.Clear"
    
End Sub

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

    Clear

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRoute.Class_Initialize"
    
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
    RaiseError "cTradeRoute.Class_Terminate"
    
End Sub

