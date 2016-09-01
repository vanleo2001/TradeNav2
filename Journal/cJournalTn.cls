VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalTn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalTn.cls
'' Description: Application bridge for the Journal DLL to talk to Trade Navigator
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/08/2014   DAJ         Removed Picture16
'' 10/24/2014   DAJ         Core Application functions for DLL's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    AppBridge As Object                 ' Application bridge
End Type
Private m As mPrivate

Public Property Get AppBridge() As Object
    Set AppBridge = m.AppBridge
End Property
Public Property Let AppBridge(TheAppBridge As Object)
    Set m.AppBridge = TheAppBridge
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ActiveChartSymbol
'' Description: Get the symbol off of the active chart
'' Inputs:      None
'' Returns:     Symbol off of the active chart ( blank if no active chart )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ActiveChartSymbol() As String
    ActiveChartSymbol = AppBridge.ActiveChartSymbol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChartImageForHwnd
'' Description: Create a chart image for the given hWnd
'' Inputs:      Fill, Chart Caption ( out )
'' Returns:     Filename ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChartImageForHwnd(ByVal hWnd As Long, Optional strCaption As String) As String
    ChartImageForHwnd = AppBridge.ChartImageForHwnd(hWnd, strCaption)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPerformanceReport
'' Description: Get an image of a performance report
'' Inputs:      Capture File ( out ), Current Time ( out )
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPerformanceReport(strCaptureFile As String, dCurrentTime As Double) As Boolean
    GetPerformanceReport = AppBridge.GetPerformanceReport(strCaptureFile, dCurrentTime)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadChartsCombo
'' Description: Load a combo box with chart captions
'' Inputs:      Charts Combo, Image Caption
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadChartsCombo(cboCharts As ComboBox, Optional ByVal strImageCaption As String = "")
    AppBridge.LoadChartsCombo cboCharts, strImageCaption
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LookupSymbol
'' Description: Allow the user to look up a symbol
'' Inputs:      Key Pressed, Default Symbol
'' Returns:     Symbol selected ( Blank if cancelled out )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LookupSymbol(Optional ByVal KeyAscii As Long = 0&, Optional ByVal strDefaultSymbol As String = "") As String
    LookupSymbol = AppBridge.LookupSymbol(KeyAscii, strDefaultSymbol)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderForID
'' Description: Fill an order dictionary for the given ID
'' Inputs:      Order ID
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderForID(ByVal lOrderID As Long) As cBrokerMessage
    Set OrderForID = AppBridge.OrderForID(lOrderID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderTextForId
'' Description: Get the order text for the given Order ID
'' Inputs:      Order ID, Include Symbol?, Include Account?, Include ID's?
'' Returns:     Order Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderTextForId(ByVal lOrderID As Long, Optional ByVal bWithSymbol As Boolean = True, Optional ByVal bWithAccount As Boolean = False, Optional ByVal bWithID As Boolean = False) As String
    OrderTextForId = AppBridge.OrderTextForId(lOrderID, bWithSymbol, bWithAccount, bWithID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OwnerFormIcon
'' Description: Get the owner form icon from the main application
'' Inputs:      None
'' Returns:     Owner Form Icon
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OwnerFormIcon() As StdPicture
    Set OwnerFormIcon = AppBridge.OwnerFormIcon
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReportsPicture
'' Description: Grab the image for reports
'' Inputs:      None
'' Returns:     Image
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReportsPicture() As Object
    Set ReportsPicture = AppBridge.ReportsPicture
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTradeFilter
'' Description: Show the trade filter form
'' Inputs:      Account, Symbol, Auto Trade Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowTradeFilter(Optional ByVal vAccountNumberOrID As Variant = kNullData, Optional ByVal vSymbolOrSymbolID As Variant = kNullData, Optional ByVal lAutoTradeItemID As Long = kNullData)
    AppBridge.ShowTradeFilter vAccountNumberOrID, vSymbolOrSymbolID, lAutoTradeItemID
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateJournal
'' Description: Update the given journal entry
'' Inputs:      Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateJournal(ByVal lJournalID As Long)
    AppBridge.UpdateJournal lJournalID
End Sub

