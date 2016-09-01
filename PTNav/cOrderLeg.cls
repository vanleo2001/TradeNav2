VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLeg.cls
'' Description: Wrapper class for an order leg
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 03/10/2009   DAJ         Allowed for data passed back from Option Navigator
'' 12/09/2011   DAJ         Extract expiration date from Option Navigator
'' 02/14/2012   DAJ         Added multi-leg order support
'' 06/26/2012   DAJ         Get LotSize out of Option Navigator string
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 09/04/2014   DAJ         Pulled Option Navigator conversions out of trade objects
'' 10/23/2014   DAJ         Renamed mDataNav.SpreadComponents to mDataNav.SpreadComponentsForBars
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lOrderID As Long                    ' ID for the parent order for the leg
    lLegNumber As Long                  ' Number of the leg for the order
    lSymbolID As Long                   ' ID for the symbol of the leg of the order
    strSymbol As String                 ' Symbol for the leg of the order
    strExchange As String               ' Exchange for the leg of the order
    bIsBuy As Boolean                   ' Is this a Buy or a Sell?
    lMultiplier As Long                 ' Quantity multiplier of the leg of the order
    bIsEntry As Boolean                 ' Is this an Entry or an Exit?
    
    strLeadSymbol As String             ' Lead symbol for a calendar spread
    strSecondSymbol As String           ' Second symbol for a calendar spread
    
    strLegInfo As String                ' Leg information in a string
    
    Bars As cGdBars                     ' Bars object with properties for the symbol
End Type
Private m As mPrivate

Public Property Get OrderID() As Long
    OrderID = m.lOrderID
End Property
Public Property Let OrderID(ByVal lOrderID As Long)
    m.lOrderID = lOrderID
End Property

Public Property Get LegNumber() As Long
    LegNumber = m.lLegNumber
End Property
Public Property Let LegNumber(ByVal lLegNumber As Long)
    m.lLegNumber = lLegNumber
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    
    SetBarProperties m.Bars, SymbolOrSymbolID
    GetSpreadComponents
    
    BuildLegInfo
End Property

Public Property Get Exchange() As String
    Exchange = m.strExchange
End Property
Public Property Let Exchange(ByVal strExchange As String)
    m.strExchange = strExchange
End Property

Public Property Get IsBuy() As Boolean
    IsBuy = m.bIsBuy
End Property
Public Property Let IsBuy(ByVal bIsBuy As Boolean)
    m.bIsBuy = bIsBuy
    BuildLegInfo
End Property

Public Property Get Multiplier() As Long
    Multiplier = m.lMultiplier
End Property
Public Property Let Multiplier(ByVal lMultiplier As Long)
    m.lMultiplier = lMultiplier
    BuildLegInfo
End Property

Public Property Get IsEntry() As Boolean
    IsEntry = m.bIsEntry
End Property
Public Property Let IsEntry(ByVal bIsEntry As Boolean)
    m.bIsEntry = bIsEntry
End Property

Public Property Get LegInfo() As String
    LegInfo = m.strLegInfo
End Property

Public Property Get LeadSymbol() As String
    LeadSymbol = m.strLeadSymbol
End Property

Public Property Get SecondSymbol() As String
    SecondSymbol = m.strSecondSymbol
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the order leg from the database
'' Inputs:      Order ID, Leg Number, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lOrderID As Long, ByVal lLegNumber As Long, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLegs] " & _
                    "WHERE [OrderID]=" & Str(lOrderID) & " AND [LegNumber]=" & Str(lLegNumber) & ";", dbOpenDynaset)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        m.lOrderID = rs!OrderID
        m.lLegNumber = rs!LegNumber
        m.lSymbolID = rs!SymbolID
        m.strSymbol = rs!Symbol
        m.strExchange = rs!Exchange
        m.bIsBuy = rs!IsBuy
        m.lMultiplier = rs!Multiplier
        m.bIsEntry = rs!IsEntry
        
        SetBarProperties m.Bars, SymbolOrSymbolID
        GetSpreadComponents
        
        BuildLegInfo
        
        bReturn = True
    End If
    
    Load = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLeg.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the order leg to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLegs] " & _
                "WHERE [OrderID]=" & Str(m.lOrderID) & " AND [LegNumber]=" & Str(m.lLegNumber) & ";", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!OrderID = m.lOrderID
    rs!LegNumber = m.lLegNumber
    rs!SymbolID = m.lSymbolID
    rs!Symbol = m.strSymbol
    rs!Exchange = m.strExchange
    rs!IsBuy = m.bIsBuy
    rs!Multiplier = m.lMultiplier
    rs!IsEntry = m.bIsEntry
    
    rs.Update

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create a descriptive string from the class members
'' Inputs:      Include Symbol?
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString(Optional ByVal bIncludeSymbol As Boolean = True) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If m.bIsBuy Then
        strReturn = "Buy "
    Else
        strReturn = "Sell "
    End If
    
    strReturn = strReturn & Str(m.lMultiplier)
    
    If bIncludeSymbol Then
        strReturn = strReturn & " " & m.strSymbol
    End If
    
    ToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLeg.ToString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildLegInfo
'' Description: Build the leg information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildLegInfo()
On Error GoTo ErrSection:

    If m.bIsBuy Then
        m.strLegInfo = Str(m.lMultiplier) & ";" & m.strSymbol
    Else
        m.strLegInfo = Str(m.lMultiplier * -1&) & ";" & m.strSymbol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.BuildLegInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order
'' Inputs:      None
'' Returns:     Copy of the order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cOrderLeg
On Error GoTo ErrSection:

    Dim aCopy As New cOrderLeg          ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLeg.MakeCopy"

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
    
    Set m.Bars = mCopyFrom.Bars.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSpreadComponents
'' Description: Get the individual spread components (if applicable)
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSpreadComponents()
On Error GoTo ErrSection:

    Dim astrComponents As cGdArray      ' Spread components
    
    Set astrComponents = SpreadComponentsForBars(m.Bars)
    If astrComponents.Size > 0 Then
        m.strLeadSymbol = astrComponents(0)
        m.strSecondSymbol = astrComponents(1)
    Else
        m.strLeadSymbol = ""
        m.strSecondSymbol = ""
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.GetSpreadComponents"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Bars = New cGdBars
    
    m.lLegNumber = 1&
    m.lMultiplier = 1&
    m.strLeadSymbol = ""
    m.strSecondSymbol = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Bars = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLeg.Class_Terminate"
    
End Sub
