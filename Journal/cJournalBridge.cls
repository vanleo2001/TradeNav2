VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalBridge.cls
'' Description: Bridge for the Journal DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/08/2014   DAJ         Initialize/Terminate NavCore Bridge; Fix for Journal for Transaction
'' 10/24/2014   DAJ         Core Application functions for DLL's; Trade Tracker Database object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Property Get AltGridRowColor() As Long
    AltGridRowColor = ALT_GRID_ROW_COLOR
End Property
Public Property Let AltGridRowColor(ByVal lAltGridRowColor As Long)
    ALT_GRID_ROW_COLOR = lAltGridRowColor
End Property

Public Property Get AppBridge() As Object
    Set AppBridge = g.AppBridge.AppBridge
End Property
Public Property Let AppBridge(TheAppBridge As Object)
    g.AppBridge.AppBridge = TheAppBridge
End Property

Public Property Get AppPath() As String
    AppPath = g.strAppPath
End Property
Public Property Let AppPath(ByVal strAppPath As String)
    g.strAppPath = strAppPath
End Property

Public Property Get IniFile() As String
    IniFile = g.strIniFile
End Property
Public Property Let IniFile(ByVal strIniFile As String)
    g.strIniFile = strIniFile
End Property

Public Property Get MainForm() As Object
    Set MainForm = g.frmMain
End Property
Public Property Let MainForm(frmMain As Object)
    Set g.frmMain = frmMain
End Property

Public Property Get TnCore() As Object
    Set TnCore = g.TnCore.CoreBridge
End Property
Public Property Let TnCore(TheTnCore As Object)
    g.TnCore.CoreBridge = TheTnCore
End Property

Public Property Get TradingDatabase() As Database
    Set TradingDatabase = g.TradeTrackerDB.Database
End Property
Public Property Let TradingDatabase(TradeTrackerDatabase As Database)
    g.TradeTrackerDB.Database = TradeTrackerDatabase
End Property

Public Property Get AppIsIde() As Boolean
    AppIsIde = g.bAppIsIde
End Property
Public Property Let AppIsIde(ByVal bAppIsIde As Boolean)
    g.bAppIsIde = bAppIsIde
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the bridge and global variables for the DLL
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init()
On Error GoTo ErrSection:

    Set g.JournalCategories = New cJournalCategories
    g.JournalCategories.Load AddSlash(g.strAppPath) & "Provided\JournalCategories.TXT"
    
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.Init"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Shutdown
'' Description: Clean up the bridge and global variables for the DLL
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Shutdown()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.Shutdown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowDateJournals
'' Description: Show the date journals form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowDateJournals()
On Error GoTo ErrSection:

    frmDateJournals.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.ShowDateJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowJournals
'' Description: Show the journals form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowJournals()
On Error GoTo ErrSection:

    frmJournals.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.ShowJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowOrderJournal
'' Description: Show the order journal form for the given order
'' Inputs:      Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowOrderJournal(ByVal Order As cBrokerMessage) As Boolean
On Error GoTo ErrSection:

    ShowOrderJournal = frmJournal.ShowMe(Order)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.ShowOrderJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowJournalForTransaction
'' Description: Show the order journal form for the given order
'' Inputs:      Order
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowJournalForTransaction(ByVal Order As cBrokerMessage) As Boolean
On Error GoTo ErrSection:
    
    Dim Journal As cJournal             ' Journal for the order
    
    Set Journal = New cJournal
    If g.JournalDB.LoadOrderJournalForOrderID(CLng(Val(Order("OrderID"))), Journal) = True Then
        ShowJournalForTransaction = frmJournal.ShowMe(Order, Journal.JournalID, Journal)
    Else
        ShowJournalForTransaction = frmJournal.ShowMe(Order)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.ShowJournalForTransaction"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrderJournal
'' Description: Create an order journal entry with the given note
'' Inputs:      Order ID, Note Date, Note
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateOrderJournal(ByVal lOrderID As Long, ByVal dNoteDate As Double, ByVal strNote As String)
On Error GoTo ErrSection:

    Dim OrderJournal As cJournal        ' Order journal object
    
    Set OrderJournal = New cJournal
    
    OrderJournal.OrderID = lOrderID
    OrderJournal.EmotionNumber = -1&
    OrderJournal.NoteDate = dNoteDate
    OrderJournal.JournalDate = CDbl(Int(dNoteDate))
    OrderJournal.Note = strNote
    g.JournalDB.SaveOrderJournal OrderJournal
    
    If FormIsLoaded("frmDateJournals") Then
        frmDateJournals.UpdateOrderJournal OrderJournal
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.CreateOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AutoJournalForFill
'' Description: Create a fill journal for the user based on the given fill
'' Inputs:      Current Time, Text, Chart Caption, Chart Image
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoJournalForFill(ByVal dCurrentTime As Double, ByVal strText As String, ByVal strChartCaption As String, ByVal strChartImage As String)
On Error GoTo ErrSection:

    Dim journalEntry As cDateJournal    ' Date Journal entry
    Dim JournalImage As cJournalImage   ' Journal image
        
    Set journalEntry = New cDateJournal
    
    journalEntry.JournalCategoryID = g.JournalCategories.IdForName("Fill")
    If journalEntry.JournalCategoryID > 0& Then
        journalEntry.JournalDate = CDbl(Int(dCurrentTime))
        journalEntry.JournalTime = dCurrentTime - journalEntry.JournalDate
        journalEntry.SymbolOrSymbolID = ""
        journalEntry.Text = strText
            
        If Len(strChartImage) > 0 Then
            Set JournalImage = New cJournalImage
            
            JournalImage.FileName = strChartImage
            If Len(JournalImage.FileName) > 0 Then
                JournalImage.Caption = strChartCaption
            End If
            JournalImage.DateJournalID = journalEntry.DateJournalID
            journalEntry.JournalImage(eGDJournalImageType_Chart) = JournalImage
        End If
        
        g.JournalDB.SaveDateJournal journalEntry
        
        If FormIsLoaded("frmDateJournals") Then
            frmDateJournals.UpdateDateJournal journalEntry
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.AutoJournalForFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderJournal
'' Description: Delete the order journal with the given ID
'' Inputs:      Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteOrderJournal(ByVal lJournalID As Long)
On Error GoTo ErrSection:

    Dim Journal As cJournal             ' Journal object
    
    Set Journal = New cJournal
    If g.JournalDB.LoadOrderJournal(lJournalID, Journal) = True Then
        g.JournalDB.DeleteOrderJournal Journal
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.DeleteOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderIDChanged
'' Description: Handle the order ID changing on an order
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub OrderIDChanged(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    If FormIsLoaded("frmJournal") Then
        If frmJournal.OrderID = lOldOrderID Then
            frmJournal.NewOrderID = lNewOrderID
        Else
            g.JournalDB.CopyOrderJournal lOldOrderID, lNewOrderID
        End If
    Else
        g.JournalDB.CopyOrderJournal lOldOrderID, lNewOrderID
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.OrderIDChanged"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set g.AppBridge = New cJournalTn
    Set g.JournalDB = New cJournalDatabase
    Set g.CoreBridge = New cCoreBridge
    Set g.TnCore = New cCoreTn
    Set g.TradeTrackerDB = New cTradeTrackerDb
    
    g.bAppIsIde = False

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set g.JournalCategories = Nothing
    Set g.JournalDB = Nothing
    Set g.CoreBridge = Nothing
    Set g.TnCore = Nothing
    Set g.TradeTrackerDB = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalBridge.Class_Terminate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAppBackColor (copied from NavReports2)
'' Description: Set app back color for TradeNav theme
'' Inputs:      nBackColor
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAppBackColor(ByVal nBackColor As Long)
    On Error Resume Next
    mGenesis.SetAppBackColor nBackColor
End Sub
