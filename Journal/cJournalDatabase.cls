VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJournalDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cJournalDatabase.cls
'' Description: Class to handle database interaction for journal stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/08/2014   DAJ         Added LoadOrderJournalForOrderID
'' 10/24/2014   DAJ         Core Application functions for DLL's; Trade Tracker Database object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyOrderJournal
'' Description: Copy the order journal entries from one order to another
'' Inputs:      Old Order ID, New Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CopyOrderJournal(ByVal lOldOrderID As Long, ByVal lNewOrderID As Long)
On Error GoTo ErrSection:

    Dim rsOld As Recordset              ' Recordset into the database
    Dim rsNew As Recordset              ' Recordset into the database
    
    Set rsOld = g.TradeTrackerDB.OrderJournalForOrderId(lOldOrderID)
    Set rsNew = g.TradeTrackerDB.OrderJournalForOrderId(lNewOrderID)
    Do While Not rsOld.EOF
        rsNew.AddNew
        rsNew!OrderID = lNewOrderID
        rsNew!NoteDate = rsOld!NoteDate
        rsNew!Note = rsOld!Note
        rsNew.Update
        
        rsOld.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.CopyOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetJournalDates
'' Description: Get an array of dates that have a journal entry
'' Inputs:      None
'' Returns:     Journal Dates
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetJournalDates() As cGdArray
On Error GoTo ErrSection:

    Dim alReturn As cGdArray            ' Array of journal dates to return from the function
    Dim rs As Recordset                 ' Recordset into the database
    Dim lPos As Long                    ' Position of date in the array
    
    Set alReturn = New cGdArray
    alReturn.Create eGDARRAY_Longs
    
    Set rs = g.TradeTrackerDB.DateJournalDates
    Do While Not rs.EOF
        If alReturn.BinarySearch(CLng(rs!JournalDate), lPos) = False Then
            alReturn.Add CLng(rs!JournalDate), lPos
        End If
        
        rs.MoveNext
    Loop

    Set rs = g.TradeTrackerDB.OrderJournalDates
    Do While Not rs.EOF
        If alReturn.BinarySearch(CLng(rs!JournalDate), lPos) = False Then
            alReturn.Add CLng(rs!JournalDate), lPos
        End If
        
        rs.MoveNext
    Loop
    
    Set GetJournalDates = alReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetJournalDates"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolsForOrderJournals
'' Description: Get a list of unique symbols for order journals
'' Inputs:      None
'' Returns:     Order Journal symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbolsForOrderJournals() As cGdArray
On Error GoTo ErrSection:

    Set GetSymbolsForOrderJournals = GetSymbolsForRecordset(g.TradeTrackerDB.AllOrderJournals)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetSymbolsForOrderJournals"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolsForMoneyCodeJournals
'' Description: Get a list of unique symbols for money code journals
'' Inputs:      None
'' Returns:     Money Code Journal symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbolsForMoneyCodeJournals() As cGdArray
On Error GoTo ErrSection:

    Set GetSymbolsForMoneyCodeJournals = GetSymbolsForRecordset(g.TradeTrackerDB.MoneyCodeJournals)

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetSymbolsForMoneyCodeJournals"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDateJournals
'' Description: Load the date journals from the database
'' Inputs:      Date Journals
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDateJournals(DateJournals As cDateJournals)
On Error GoTo ErrSection:

    GetDateJournalsForRecordset DateJournals, g.TradeTrackerDB.AllDateJournals

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadDateJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournals
'' Description: Load the order journals from the database
'' Inputs:      Order Journals
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadOrderJournals(OrderJournals As cJournals)
On Error GoTo ErrSection:

    GetOrderJournalsForRecordset OrderJournals, g.TradeTrackerDB.AllOrderJournalsByDateDesc

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadOrderJournals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDateJournalsForSymbol
'' Description: Load the date journals from the database for the given symbol
'' Inputs:      Date Journals, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDateJournalsForSymbol(DateJournals As cDateJournals, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID from the value passed in
    Dim strSymbol As String             ' Symbol from the value passed in
    
    lSymbolID = g.TnCore.GetSymbolID(vSymbolOrSymbolID)
    strSymbol = g.TnCore.GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID = 0 Then
        GetDateJournalsForRecordset DateJournals, g.TradeTrackerDB.DateJournalsForSymbol(strSymbol)
    Else
        GetDateJournalsForRecordset DateJournals, g.TradeTrackerDB.DateJournalsForSymbolID(lSymbolID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadDateJournalsForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournalsForSymbol
'' Description: Load the order journals from the database for the given symbol
'' Inputs:      Order Journals, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadOrderJournalsForSymbol(OrderJournals As cJournals, ByVal vSymbolOrSymbolID As Variant)
On Error GoTo ErrSection:

    Dim lSymbolID As Long               ' Symbol ID from the value passed in
    Dim strSymbol As String             ' Symbol from the value passed in
    
    lSymbolID = g.TnCore.GetSymbolID(vSymbolOrSymbolID)
    strSymbol = g.TnCore.GetSymbol(vSymbolOrSymbolID)
    
    If lSymbolID = 0 Then
        GetOrderJournalsForRecordset OrderJournals, g.TradeTrackerDB.OrderJournalsForSymbol(strSymbol)
    Else
        GetOrderJournalsForRecordset OrderJournals, g.TradeTrackerDB.OrderJournalsForSymbolID(lSymbolID)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadDateJournalsForSymbol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDateJournalsForDate
'' Description: Load the date journals from the database for the given date
'' Inputs:      Date Journals, Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadDateJournalsForDate(DateJournals As cDateJournals, ByVal lDate As Long)
On Error GoTo ErrSection:

    GetDateJournalsForRecordset DateJournals, g.TradeTrackerDB.DateJournalsForDate(lDate)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadDateJournalsForDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournalsForDate
'' Description: Load the order journals from the database for the given date
'' Inputs:      Order Journals, Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadOrderJournalsForDate(OrderJournals As cJournals, ByVal lDate As Long)
On Error GoTo ErrSection:

    GetOrderJournalsForRecordset OrderJournals, g.TradeTrackerDB.OrderJournalsForDate(lDate)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadOrderJournalsForDate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournalsForAccountID
'' Description: Load the order journals from the database for the account ID
'' Inputs:      Order Journals, Account ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadOrderJournalsForAccountID(OrderJournals As cJournals, ByVal lAccountID As Long)
On Error GoTo ErrSection:

    GetOrderJournalsForRecordset OrderJournals, g.TradeTrackerDB.OrderJournalsForAccountId(lAccountID)

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cDateJournals.LoadOrderJournalsForAccountID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDateJournal
'' Description: Load the date journal
'' Inputs:      Date Journal ID, Date Journal object, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDateJournal(ByVal lDateJournalID As Long, DateJournal As cDateJournal, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rsImages As Recordset           ' Recordset for the date journal images
    Dim JournalImage As cJournalImage   ' Journal image information
    Dim lSymbolID As Long               ' Symbol ID
    
    bReturn = False
    
    If rs Is Nothing Then
        Set rs = g.TradeTrackerDB.DateJournalForId(lDateJournalID)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        DateJournal.DateJournalID = rs!DateJournalID
        DateJournal.JournalDate = rs!JournalDate
        DateJournal.JournalCategoryID = rs!JournalCategoryID
        DateJournal.JournalTime = rs!JournalTime
        DateJournal.Text = rs!Text
        
        lSymbolID = NullChk(rs!SymbolID, 0&)
        If lSymbolID = 0& Then
            DateJournal.SymbolOrSymbolID = NullChk(rs!Symbol, "")
        Else
            DateJournal.SymbolOrSymbolID = lSymbolID
        End If
        
        DateJournal.JournalImages.Clear
        Set rsImages = g.TradeTrackerDB.DateJournalImagesForId(lDateJournalID)
        Do While Not rsImages.EOF
            Set JournalImage = New cJournalImage
            If LoadJournalImage(rsImages!ID, JournalImage, rsImages) Then
                DateJournal.JournalImages.Add JournalImage, Str(JournalImage.ImageType)
            End If
            
            rsImages.MoveNext
        Loop
        
        bReturn = True
    End If

    LoadDateJournal = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.LoadDateJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournal
'' Description: Load the order journal
'' Inputs:      Order Journal ID, Order Journal object, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadOrderJournal(ByVal lOrderJournalID As Long, OrderJournal As cJournal, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rsImages As Recordset           ' Recordset for the date journal images
    Dim JournalImage As cJournalImage   ' Journal image information
    Dim lSymbolID As Long               ' Symbol ID
    
    bReturn = False
    
    If rs Is Nothing Then
        Set rs = g.TradeTrackerDB.OrderJournalForId(lOrderJournalID)
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        OrderJournal.JournalID = NullChk(rs!JournalID, 0&)
        OrderJournal.OrderID = NullChk(rs!OrderID, 0&)
        OrderJournal.NoteDate = NullChk(rs!NoteDate, 0#)
        OrderJournal.Note = NullChk(rs!Note, "")
        OrderJournal.Action = NullChk(rs!Action, "")
        OrderJournal.Feelings = NullChk(rs!Feelings, "")
        OrderJournal.WhyTrade = NullChk(rs!WhyTrade, "")
        OrderJournal.Thoughts = NullChk(rs!Thoughts, "")
        OrderJournal.EmotionNumber = NullChk(rs!EmotionNumber, -1&)
        OrderJournal.AccountID = NullChk(rs!AccountID, 0&)
        OrderJournal.JournalDate = NullChk(rs!JournalDate, 0#)
        
        lSymbolID = NullChk(rs!SymbolID, 0&)
        If lSymbolID = 0& Then
            OrderJournal.SymbolOrSymbolID = NullChk(rs!Symbol, "")
        Else
            OrderJournal.SymbolOrSymbolID = lSymbolID
        End If
        
        OrderJournal.JournalImages.Clear
        Set rsImages = g.TradeTrackerDB.OrderJournalImagesForId(lOrderJournalID)
        Do While Not rsImages.EOF
            Set JournalImage = New cJournalImage
            If LoadJournalImage(rsImages!ID, JournalImage, rsImages) Then
                OrderJournal.JournalImages.Add JournalImage, Str(JournalImage.ImageType)
            End If
            
            rsImages.MoveNext
        Loop
        
        bReturn = True
    End If

    LoadOrderJournal = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.LoadOrderJournal"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOrderJournalForOrderID
'' Description: Load the order journal for the given Order ID
'' Inputs:      Order ID, Order Journal object, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadOrderJournalForOrderID(ByVal lOrderID As Long, OrderJournal As cJournal) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    Set rs = g.TradeTrackerDB.OrderJournalForOrderId(lOrderID)
    If Not (rs.BOF And rs.EOF) Then
        bReturn = LoadOrderJournal(rs!JournalID, OrderJournal, rs)
    End If
    
    LoadOrderJournalForOrderID = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.LoadOrderJournalForOrderID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveDateJournal
'' Description: Save the date journal to the database
'' Inputs:      Date Journal object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveDateJournal(DateJournal As cDateJournal)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    ' Save the date journal...
    Set rs = g.TradeTrackerDB.DateJournalForId(DateJournal.DateJournalID)
    If rs.BOF And rs.EOF Then
        rs.AddNew
        DateJournal.DateJournalID = rs!DateJournalID
    Else
        rs.Edit
    End If
    
    rs!JournalDate = DateJournal.JournalDate
    rs!JournalCategoryID = DateJournal.JournalCategoryID
    rs!JournalTime = DateJournal.JournalTime
    rs!Text = DateJournal.Text
    rs!SymbolID = DateJournal.SymbolID
    rs!Symbol = DateJournal.Symbol
    
    rs.Update
    
    ' Save the journal images...
    SaveJournalImages DateJournal.JournalImages, DateJournal.DateJournalID

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.SaveDateJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveOrderJournal
'' Description: Save the order journal to the database
'' Inputs:      Date Order object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveOrderJournal(OrderJournal As cJournal)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.TradeTrackerDB.OrderJournalForId(OrderJournal.JournalID)
    If rs.BOF And rs.EOF Then
        rs.AddNew
        OrderJournal.JournalID = rs!JournalID
    Else
        rs.Edit
    End If
    
    rs!OrderID = OrderJournal.OrderID
    rs!NoteDate = OrderJournal.NoteDate
    rs!Note = OrderJournal.Note
    rs!Action = OrderJournal.Action
    rs!Feelings = OrderJournal.Feelings
    rs!WhyTrade = OrderJournal.WhyTrade
    rs!Thoughts = OrderJournal.Thoughts
    rs!EmotionNumber = OrderJournal.EmotionNumber
    rs!AccountID = OrderJournal.AccountID
    rs!JournalDate = OrderJournal.JournalDate
    rs!SymbolID = OrderJournal.SymbolID
    rs!Symbol = OrderJournal.Symbol
    
    rs.Update
    
    SaveJournalImages OrderJournal.JournalImages, , OrderJournal.JournalID
                
ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournal.SaveOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteDateJournal
'' Description: Delete the date journal out of the database
'' Inputs:      Date Journal object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteDateJournal(DateJournal As cDateJournal)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim JournalImage As cJournalImage   ' Journal image
    
    Set rs = g.TradeTrackerDB.DateJournalForId(DateJournal.DateJournalID)
    If g.TradeTrackerDB.DeleteRecord(rs) Then
        DeleteJournalImages DateJournal.JournalImages
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.DeleteDateJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderJournal
'' Description: Delete the order journal out of the database
'' Inputs:      Order Journal object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteOrderJournal(OrderJournal As cJournal)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.TradeTrackerDB.OrderJournalForId(OrderJournal.JournalID)
    If g.TradeTrackerDB.DeleteRecord(rs) Then
        DeleteJournalImages OrderJournal.JournalImages
    End If

    If FormIsLoaded("frmDateJournals") Then
        frmDateJournals.DeleteOrderJournal OrderJournal
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.DeleteOrderJournal"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadJournalImage
'' Description: Load the journal image
'' Inputs:      Journal Image ID, Journal Image object, Recordset
'' Returns:     True if found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadJournalImage(ByVal lJournalImageID As Long, JournalImage As cJournalImage, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If rs Is Nothing Then
        Set rs = g.TradeTrackerDB.JournalImageForId(lJournalImageID)
    End If

    If Not (rs.BOF And rs.EOF) Then
        JournalImage.ID = rs!ID
        JournalImage.ImageType = rs!Type
        JournalImage.FileName = rs!FileName
        JournalImage.Caption = rs!Caption
        JournalImage.OrderJournalID = rs!OrderJournalID
        JournalImage.DateJournalID = rs!DateJournalID
        
        bReturn = True
    End If

    LoadJournalImage = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.LoadJournalImage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveJournalImage
'' Description: Save the journal image
'' Inputs:      Journal Image object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveJournalImage(JournalImage As cJournalImage)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.TradeTrackerDB.JournalImageForId(JournalImage.ID)
    If rs.BOF And rs.EOF Then
        rs.AddNew
        JournalImage.ID = rs!ID
    Else
        rs.Edit
    End If
    
    rs!Type = JournalImage.ImageType
    rs!FileName = JournalImage.FileName
    rs!Caption = JournalImage.Caption
    rs!OrderJournalID = JournalImage.OrderJournalID
    rs!DateJournalID = JournalImage.DateJournalID
    
    rs.Update

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.SaveJournalImage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteJournalImage
'' Description: Delete the journal image
'' Inputs:      Journal Image object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteJournalImage(JournalImage As cJournalImage)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.TradeTrackerDB.JournalImageForId(JournalImage.ID)
    g.TradeTrackerDB.DeleteRecord rs

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.DeleteJournalImage"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveJournalImages
'' Description: Save the given journal images to the database
'' Inputs:      Journal Images, Date Journal ID, Order Journal ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveJournalImages(JournalImages As cGdTree, Optional ByVal lDateJournalID As Long = kNullData, Optional ByVal lOrderJournalID As Long = kNullData)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    Dim JournalImage As cJournalImage   ' Journal image
    Dim alJournalIds As cGdArray        ' Array of journal ID's
    Dim lPos As Long                    ' Position in the array

    Set alJournalIds = New cGdArray
    alJournalIds.Create eGDARRAY_Longs
    
    ' Save the journal images...
    For lIndex = 1 To JournalImages.Count
        Set JournalImage = JournalImages(lIndex)
        If lDateJournalID <> kNullData Then
            JournalImage.DateJournalID = lDateJournalID
        End If
        If lOrderJournalID <> kNullData Then
            JournalImage.OrderJournalID = lOrderJournalID
        End If
        
        SaveJournalImage JournalImage
        If alJournalIds.BinarySearch(JournalImage.ID, lPos) = False Then
            alJournalIds.Add JournalImage.ID, lPos
        End If
        
        Set JournalImages(lIndex) = JournalImage
    Next lIndex
        
    ' Remove any old journal images...
    If lDateJournalID <> kNullData Then
        Set rs = g.TradeTrackerDB.DateJournalImagesForId(lDateJournalID)
    ElseIf lOrderJournalID <> kNullData Then
        Set rs = g.TradeTrackerDB.OrderJournalImagesForId(lOrderJournalID)
    End If
    
    Do While Not rs.EOF
        If alJournalIds.BinarySearch(rs!ID) = False Then
            rs.Delete
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.SaveJournalImages"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolsForRecordset
'' Description: Get a list of symbols for the given recordset
'' Inputs:      Recordset
'' Returns:     Array of symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetSymbolsForRecordset(rs As Recordset) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of symbols to return from the function
    Dim lPos As Long                    ' Position in the array
    Dim strSymbol As String             ' Symbol from the recordset
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings

    Do While Not rs.EOF
        If rs!SymbolID <> 0 Then
            strSymbol = g.TnCore.GetSymbol(rs!SymbolID)
        Else
            strSymbol = rs!Symbol
        End If
        
        If Len(strSymbol) > 0 Then
            If astrReturn.BinarySearch(strSymbol, lPos) = False Then
                astrReturn.Add strSymbol, lPos
            End If
        End If
        
        rs.MoveNext
    Loop
    
    Set GetSymbolsForRecordset = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetSymbolsForRecordset"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDateJournalsForRecordset
'' Description: Fill the given date journals collection from the given recordset
'' Inputs:      Date Journals object, Recordset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetDateJournalsForRecordset(DateJournals As cDateJournals, rs As Recordset)
On Error GoTo ErrSection:

    Dim DateJournal As cDateJournal     ' Date Journal object

    DateJournals.Clear

    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set DateJournal = New cDateJournal
            If LoadDateJournal(rs!DateJournalID, DateJournal, rs) Then
                DateJournals.Add DateJournal
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetDateJournalsForRecordset"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOrderJournalsForRecordset
'' Description: Fill the given order journals collection from the given recordset
'' Inputs:      Order Journals object, Recordset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetOrderJournalsForRecordset(OrderJournals As cJournals, rs As Recordset)
On Error GoTo ErrSection:

    Dim OrderJournal As cJournal        ' Order Journal object

    OrderJournals.Clear

    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            Set OrderJournal = New cJournal
            If LoadOrderJournal(rs!JournalID, OrderJournal, rs) = True Then
                OrderJournals.Add OrderJournal
            End If
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.GetOrderJournalsForRecordset"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteJournalImages
'' Description: Delete the journal images
'' Inputs:      Journal Images collection
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteJournalImages(JournalImages As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim JournalImage As cJournalImage   ' Journal image
    
    For lIndex = 1 To JournalImages.Count
        Set JournalImage = JournalImages(lIndex)
        
        If FileExist(JournalImage.FileName) Then
            KillFile JournalImage.FileName
        End If
        
        DeleteJournalImage JournalImage
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.DeleteJournalImages"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    g.TnCore.RaiseError "cJournalDatabase.Class_Terminate"
    
End Sub

