VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeTrackerDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeTrackerDb.cls
'' Description: Class for accessing the Trade Tracker database
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/10/2014   DAJ         Created
'' 10/24/2014   DAJ         Core Application functions for DLL's; Trade Tracker database object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    TradeTrackerDB As Database          ' Database object
End Type
Private m As mPrivate

Public Property Get Database() As Database
    Set Database = m.TradeTrackerDB
End Property
Public Property Let Database(TradeTrackerDB As Database)
    Set m.TradeTrackerDB = TradeTrackerDB
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteRecord
'' Description: Delete the record from the database
'' Inputs:      Recordset
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteRecord(rs As Recordset) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        bReturn = True
    End If
    
    DeleteRecord = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cTradeTrackerDb.DeleteRecord"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountForID
'' Description: Get the record out of the accounts table for the given account ID
'' Inputs:      Account ID
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountForID(ByVal lAccountID As Long) As Recordset
    Set AccountForID = OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AccountsForAccountType
'' Description: Get the records out of the accounts table for the given account type
'' Inputs:      Account ID
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AccountsForAccountType(ByVal nAccountType As eTT_AccountType) As Recordset
    Set AccountsForAccountType = OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountType]=" & Str(nAccountType) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllDateJournals
'' Description: Get all date journals
'' Inputs:      None
'' Returns:     Date Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllDateJournals() As Recordset
    Set AllDateJournals = OpenRecordset("SELECT * FROM [tblDateJournals];", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllOrderJournals
'' Description: Get all order journals
'' Inputs:      None
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllOrderJournals() As Recordset
    Set AllOrderJournals = OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllOrderJournalsByDateDesc
'' Description: Get all order journals sorted by date descending
'' Inputs:      None
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AllOrderJournalsByDateDesc() As Recordset
    Set AllOrderJournalsByDateDesc = OpenRecordset("SELECT * FROM [tblOrderJournal] ORDER BY [NoteDate] DESC;", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalDates
'' Description: Get the distinct journal dates for date journals
'' Inputs:      None
'' Returns:     Date Journal Dates
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalDates() As Recordset
    Set DateJournalDates = OpenRecordset("SELECT DISTINCT JournalDate FROM [tblDateJournals];", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalForId
'' Description: Get the date journal record for the given ID
'' Inputs:      Date Journal ID
'' Returns:     Date Journal Record
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalForId(ByVal lDateJournalID As Long) As Recordset
    Set DateJournalForId = OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [DateJournalID]=" & Str(lDateJournalID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalImagesForId
'' Description: Get the date journal images for the given ID
'' Inputs:      Date Journal ID
'' Returns:     Date Journal Images
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalImagesForId(ByVal lDateJournalID As Long) As Recordset
    Set DateJournalImagesForId = OpenRecordset("SELECT * FROM [tblJournalImages] WHERE [DateJournalID]=" & Str(lDateJournalID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalsForDate
'' Description: Get all date journals for the given date
'' Inputs:      Date
'' Returns:     Date Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalsForDate(ByVal lDate As Long) As Recordset
    Set DateJournalsForDate = OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [JournalDate]=" & Str(lDate) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalsForSymbol
'' Description: Get all date journals for the given symbol
'' Inputs:      Symbol
'' Returns:     Date Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalsForSymbol(ByVal strSymbol As String) As Recordset
    Set DateJournalsForSymbol = OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [Symbol]='" & strSymbol & "';", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DateJournalsForSymbolID
'' Description: Get all date journals for the given symbol ID
'' Inputs:      Symbol ID
'' Returns:     Date Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DateJournalsForSymbolID(ByVal lSymbolID As Long) As Recordset
    Set DateJournalsForSymbolID = OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [SymbolID]=" & Str(lSymbolID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillsForAccountID
'' Description: Get the records out of the fills table for the given account ID
'' Inputs:      Account ID
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FillsForAccountID(ByVal lAccountID As Long) As Recordset
    Set FillsForAccountID = OpenRecordset("SELECT * FROM [tblFills] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    JournalImageForId
'' Description: Get the journal image for the given ID
'' Inputs:      Journal Image ID
'' Returns:     Journal Image Record
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function JournalImageForId(ByVal lJournalImageID As Long) As Recordset
    Set JournalImageForId = OpenRecordset("SELECT * FROM [tblJournalImages] WHERE [ID]=" & Str(lJournalImageID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoneyCodeJournals
'' Description: Get money code journals
'' Inputs:      None
'' Returns:     Money Code Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MoneyCodeJournals() As Recordset
    Set MoneyCodeJournals = OpenRecordset("SELECT * FROM [tblDateJournals] WHERE [JournalCategoryID]=4;", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalDates
'' Description: Get the distinct journal dates for order journals
'' Inputs:      None
'' Returns:     Order Journal Dates
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalDates() As Recordset
    Set OrderJournalDates = OpenRecordset("SELECT DISTINCT JournalDate FROM [tblOrderJournal];", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalForId
'' Description: Get the order journal record for the given ID
'' Inputs:      Order Journal ID
'' Returns:     Order Journal Record
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalForId(ByVal lOrderJournalID As Long) As Recordset
    Set OrderJournalForId = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [JournalID]=" & Str(lOrderJournalID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalImagesForId
'' Description: Get the order journal images for the given ID
'' Inputs:      Order Journal ID
'' Returns:     Order Journal Images
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalImagesForId(ByVal lOrderJournalID As Long) As Recordset
    Set OrderJournalImagesForId = OpenRecordset("SELECT * FROM [tblJournalImages] WHERE [OrderJournalID]=" & Str(lOrderJournalID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalsForAccountId
'' Description: Get all order journals for the given account ID
'' Inputs:      Account ID
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalsForAccountId(ByVal lAccountID As Long) As Recordset
    Set OrderJournalsForAccountId = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalsForDate
'' Description: Get all order journals for the given date
'' Inputs:      Date
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalsForDate(ByVal lDate As Long) As Recordset
    Set OrderJournalsForDate = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [JournalDate]=" & Str(lDate) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalForOrderId
'' Description: Get the order journal record for the given Order ID
'' Inputs:      Order ID
'' Returns:     Order Journal record
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalForOrderId(ByVal lOrderID As Long) As Recordset
    Set OrderJournalForOrderId = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [OrderID]=" & Str(lOrderID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalsForSymbol
'' Description: Get all order journals for the given symbol
'' Inputs:      Symbol
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalsForSymbol(ByVal strSymbol As String) As Recordset
    Set OrderJournalsForSymbol = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [Symbol]='" & strSymbol & "';", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderJournalsForSymbolID
'' Description: Get all order journals for the given symbol ID
'' Inputs:      Symbol ID
'' Returns:     Order Journals
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OrderJournalsForSymbolID(ByVal lSymbolID As Long) As Recordset
    Set OrderJournalsForSymbolID = OpenRecordset("SELECT * FROM [tblOrderJournal] WHERE [SymbolID]=" & Str(lSymbolID) & ";", dbOpenDynaset)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OpenRecordset
'' Description: Open a recordset in the database
'' Inputs:      Query, Type
'' Returns:     Recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OpenRecordset(ByVal strQuery As String, Optional ByVal vType As Variant) As Recordset
    Set OpenRecordset = m.TradeTrackerDB.OpenRecordset(strQuery, vType)
End Function

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

    Set m.TradeTrackerDB = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cTradeTrackerDb.Class_Initialize"
    
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

    Set m.TradeTrackerDB = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cTradeTrackerDb.Class_Terminate"
    
End Sub
