VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTTUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTTUpdates.cls
'' Description: Handles the structural updates to the Trade Tracker database
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 01/21/2009   DAJ         Added the Max Profit/Max Loss fields to the database
'' 04/20/2009   DAJ         Added emotion number to the order journal table
'' 04/23/2009   DAJ         Added AccountID to order journal, fix nulls in data
'' 06/01/2009   DAJ         Added FO and SO to security type mask for trading
'' 08/12/2009   DAJ         Added S and SO to security type mask for PFG
'' 08/13/2009   DAJ         Added the IsAutoExit flag for orders
'' 09/01/2009   DAJ         Use new Parked order status
'' 10/07/2009   DAJ         Added support for Linked Orders at broker
'' 12/01/2009   DAJ         Added support for automatic commissions on fills
'' 12/01/2009   DAJ         Enhancements for Auto Exits held at broker
'' 12/04/2009   DAJ         Changed automatic commissions for stocks
'' 05/27/2011   DAJ         Added IsAutomated flag for orders
'' 06/21/2011   DAJ         Separate out Simulated trading types
'' 06/22/2011   DAJ         For now, don't auto convert TSU accounts to SimBroker
'' 07/28/2011   DAJ         Convert all non-Stream Replay simulated accounts to SimBroker
'' 09/20/2011   DAJ         Added date journal stuff, added journal date to order journals
'' 09/21/2011   DAJ         Changed conversion of simulated accounts
'' 09/22/2011   DAJ         Added filename and caption for attached chart image on journal entries
'' 09/23/2011   DAJ         Change date journal text to memo, add indexes on journal date fields
'' 12/09/2011   DAJ         Fix Interactive Brokers Broker IDs, remove order and fill duplicates
'' 01/30/2012   DAJ         Option Nav Journal Image
'' 02/14/2012   DAJ         Move fields from the Order Legs table up to the Orders table
'' 02/14/2012   DAJ         Added multi-leg order support
'' 03/19/2012   DAJ         Added Symbol/Symbol ID to the order journal and date journal tables
'' 06/26/2012   DAJ         Added symbol info table
'' 08/03/2012   DAJ         Remove Alaron
'' 08/21/2012   DAJ         Added FcmAccountNumber for accounts
'' 12/11/2012   DAJ         Contingency Orders, OcaGroup name for orders
'' 01/31/2013   DAJ         Simulated/CQG Trading for Calendar Spread Symbols
'' 10/16/2013   DAJ         Removed PFG and Xpress
'' 09/02/2014   DAJ         Move Journal stuff into Journal DLL
'' 10/17/2014   DAJ         Replaced File System Object calls
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kCurrentVersion = 66

Private Type mPrivate
    DB As Database                      ' Database object
End Type
Private m As mPrivate

Private Sub Class_Terminate()
    Set m.DB = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Version (Get)
'' Description: Get the version number from the database, or if the database
''              table does not exist, create it and return Version 1
'' Inputs:      None
'' Returns:     Version of the TradeTracker database
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Version() As Long
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset to get the version
    
    ' Get database "Header" recordset (one record table)...
    If ItemExists(m.DB.TableDefs, "tblDatabase") Then
        Set rs = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
        Version = rs!Version
        
    ' Otherwise create one if it does not exist...
    Else
        Dim t As TableDef
        Set t = New TableDef
        With t
            .Name = "tblDatabase"
            .Fields.Append .CreateField("Version", dbLong)
        End With
        m.DB.TableDefs.Append t
        
        Set rs = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
        With rs
            .AddNew
            !Version = 1
            .Update
            .MoveFirst
        End With
        
        ' Also delete obsolete table...
        If ItemExists(m.DB.TableDefs, "tblVersion") Then
            m.DB.TableDefs.Delete "tblVersion"
        End If
        Set t = Nothing
        Version = 1
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Property
    
ErrSection:
    Set rs = Nothing
    RaiseError "cTTUpdates.Version.Get", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Version (Let)
'' Description: Set the database version to the one passed in
'' Inputs:      New Version Number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let Version(ByVal lNewVersion As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
    With rs
        .Edit
        !Version = lNewVersion
        .Update
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Property
    
ErrSection:
    Set rs = Nothing
    RaiseError "cTTUpdates.Version.Let", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DB (Get)
'' Description: Return the database
'' Inputs:      None
'' Returns:     Database
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get DB() As Database
On Error GoTo ErrSection:

    Set DB = m.DB
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTTUpdates.DB.Get", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DB (Let)
'' Description: Set the local database to the one passed in
'' Inputs:      Database to update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let DB(pData As Database)
On Error GoTo ErrSection:

    Dim lVersion As Long                ' Version of the database
    
    Set m.DB = pData
    
    ' Check version now just to make sure version table exists...
    lVersion = Version
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTTUpdates.DB.Let", eGDRaiseError_Raise
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Upgrade
'' Description: Upgrade the database to the new version
'' Inputs:      None
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Upgrade() As Boolean
On Error GoTo ErrSection:
    
    Dim lVersion As Long                ' Index into a for loop
    Dim lStartVersion As Long           ' Current version of the database
    Dim strBackup As String             ' Name of the backup database
    Dim strUpgrading As String          ' Name of the temporary backup of the database

    ' See if need to upgrade...
    lStartVersion = Version
    If lStartVersion < kCurrentVersion And lStartVersion > 0 Then
        ' Log message regarding upgrade...
        MakeDir App.Path & "\Backup\", True
        FileFromString App.Path & "\Backup\UpgradeTT.log", _
            "Your database will now be upgraded (from v." _
            & CStr(lStartVersion) & " to v." & CStr(kCurrentVersion) & ").", _
            True, True
        frmSplash.Message -1, "Applying update to TradeTracker.MDB ..."
        
        ' Make a backup before any upgrading...
        strUpgrading = ReplaceFileExt(m.DB.Name, "_!_")
        KillFile strUpgrading
        mGenesis.CopyFiles m.DB.Name, strUpgrading
        
        ' Apply each version of MDB changes...
        For lVersion = lStartVersion + 1 To kCurrentVersion
            ApplyVersion lVersion
            Version = lVersion
        Next

        ' Rename upgrading to backup after upgrade is complete...
        strBackup = App.Path & "\Backup\" _
            & FileBase(m.DB.Name) & "." _
            & Format(lStartVersion, "000")
        KillFile strBackup
        mGenesis.MoveFiles strUpgrading, strBackup
        Upgrade = True
    Else
        Upgrade = False
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTTUpdates.Upgrade", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeDataType
'' Description: Change the data type on a field in a table
'' Inputs:      Table Name, Field Name, New Data Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeDataType(ByVal strTable As String, ByVal strField As String, Optional ByVal nNewDataType As DataTypeEnum = dbLongBinary)
On Error GoTo ErrSection:

    ChangeAccessFieldDataType m.DB, strTable, strField, nNewDataType

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ChangeDataType"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion
'' Description: Apply appropriate version updates
'' Inputs:      Version to upgrade to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion(ByVal lVersion As Long)
On Error GoTo ErrSection:
    
    Dim q As QueryDef
    Dim i As Index
    Dim r As Relation
    Dim t As TableDef
    Dim rs As Recordset

    Select Case lVersion
        Case 2:
            ' Need to make sure that text fields can be zero length...
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                .Fields("AccountNumber").AllowZeroLength = True
            End With
            
            ' Need to make sure that text fields can be zero length...
            Set t = m.DB.TableDefs("tblPositions")
            With t
                .Fields("Symbol").AllowZeroLength = True
                .Fields("SecurityType").AllowZeroLength = True
                .Fields("Notes").AllowZeroLength = True
            End With
                        
        Case 3:
            ' Undo original version 2 additions (not ready for them yet)...
            Set t = m.DB.TableDefs("tblPositions")
            With t
                If ItemExists(.Fields, "Position") Then
                    .Fields.Delete "Position"
                End If
            End With
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If ItemExists(.Fields, "Fees") Then
                    .Fields.Delete "Fees"
                End If
                If ItemExists(.Fields, "Symbol") Then
                    .Fields.Delete "Symbol"
                End If
                If ItemExists(.Fields, "SymbolID") Then
                    .Fields.Delete "SymbolID"
                End If
                If ItemExists(.Fields, "Status") Then
                    .Fields.Delete "Status"
                End If
            End With
                        
        Case 4:
            ' Add UserName, Password, Account Type, Broker, and Commissions to tblAccounts...
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                If Not ItemExists(.Fields, "UserName") Then
                    .Fields.Append .CreateField("UserName", dbText, 50)
                    .Fields("UserName").Required = False
                    .Fields("UserName").AllowZeroLength = True
                    .Fields("UserName").DefaultValue = ""
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblAccounts SET tblAccounts.UserName = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Password") Then
                    .Fields.Append .CreateField("Password", dbText, 50)
                    .Fields("Password").Required = False
                    .Fields("Password").AllowZeroLength = True
                    .Fields("Password").DefaultValue = ""
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblAccounts SET tblAccounts.Password = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "AccountType") Then
                    .Fields.Append .CreateField("AccountType", dbLong)
                    .Fields("UserName").DefaultValue = 0&
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblAccounts SET tblAccounts.AccountType = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Broker") Then
                    .Fields.Append .CreateField("Broker", dbText, 50)
                    .Fields("Broker").Required = False
                    .Fields("Broker").AllowZeroLength = True
                    .Fields("Broker").DefaultValue = ""
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblAccounts SET tblAccounts.Broker = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Comms") Then
                    .Fields.Append .CreateField("Comms", dbDouble)
                    .Fields("Comms").DefaultValue = 0
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblAccounts SET tblAccounts.Comms = 0;"
                    q.Execute
                End If
            End With
            
            ' Add AdjustAmount to tblActivities...
            Set t = m.DB.TableDefs("tblActivities")
            With t
                If Not ItemExists(.Fields, "AdjustAmount") Then
                    .Fields.Append .CreateField("AdjustAmount", dbDouble)
                    .Fields("AdjustAmount").DefaultValue = 0
                End If
            End With
            
            ' Add Position and Open to tblPositions...
            Set t = m.DB.TableDefs("tblPositions")
            With t
                If Not ItemExists(.Fields, "Position") Then
                    .Fields.Append .CreateField("Position", dbLong)
                    .Fields("Position").DefaultValue = 0
                End If
                If Not ItemExists(.Fields, "Open") Then
                    .Fields.Append .CreateField("Open", dbBoolean)
                    .Fields("Open").DefaultValue = 0
                End If
                If Not ItemExists(.Fields, "OpenSymbol") Then
                    .Fields.Append .CreateField("OpenSymbol", dbText, 50)
                    .Fields("OpenSymbol").Required = False
                    .Fields("OpenSymbol").AllowZeroLength = True
                    .Fields("OpenSymbol").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblPositions SET tblPositions.OpenSymbol='';"
                    q.Execute
                End If
            End With
            
            ' Add Fees, Symbol, SymbolID, Status, Status Date, BrokerOrderID,
            ' ExchangeOrderID, Expiration, AccountID, DateSent, DateHostRecd,
            ' DateExchRecd, and DateExchAckn to tblOrders...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "Fees") Then
                    .Fields.Append .CreateField("Fees", dbDouble)
                    .Fields("Fees").DefaultValue = 0
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.Fees = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Symbol") Then
                    .Fields.Append .CreateField("Symbol", dbText)
                    .Fields("Symbol").AllowZeroLength = True
                    .Fields("Symbol").DefaultValue = ""
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.Symbol = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "SymbolID") Then
                    .Fields.Append .CreateField("SymbolID", dbLong)
                    .Fields("SymbolID").DefaultValue = 0
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.SymbolID = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Status") Then
                    .Fields.Append .CreateField("Status", dbLong)
                    .Fields("Status").DefaultValue = 0
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.Status = " & Str(eTT_OrderStatus_Filled) & ";"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "StatusDate") Then
                    .Fields.Append .CreateField("StatusDate", dbDate)
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.StatusDate = tblOrders.OrderDate;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "BrokerOrderID") Then
                    .Fields.Append .CreateField("BrokerOrderID", dbText, 10)
                    .Fields("BrokerOrderID").Required = False
                    .Fields("BrokerOrderID").AllowZeroLength = True
                    .Fields("BrokerOrderID").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.BrokerOrderID = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "ExchangeOrderID") Then
                    .Fields.Append .CreateField("ExchangeOrderID", dbText, 30)
                    .Fields("ExchangeOrderID").Required = False
                    .Fields("ExchangeOrderID").AllowZeroLength = True
                    .Fields("ExchangeOrderID").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.ExchangeOrderID = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "Expiration") Then
                    .Fields.Append .CreateField("Expiration", dbLong)
                    .Fields("Expiration").DefaultValue = 0
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.Expiration=0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "AccountID") Then
                    .Fields.Append .CreateField("AccountID", dbLong)
                    .Fields("AccountID").DefaultValue = 0
                End If
                If Not ItemExists(.Fields, "DateSent") Then
                    .Fields.Append .CreateField("DateSent", dbDate)
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.DateSent=0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "DateHostRecd") Then
                    .Fields.Append .CreateField("DateHostRecd", dbDate)
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.DateHostRecd=0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "DateExchRecd") Then
                    .Fields.Append .CreateField("DateExchRecd", dbDate)
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.DateExchRecd=0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "DateExchAckn") Then
                    .Fields.Append .CreateField("DateExchAckn", dbDate)
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.DateExchAckn=0;"
                    q.Execute
                End If
            End With
            
            ' Add BrokerFillID and BuySell to tblFills...
            Set t = m.DB.TableDefs("tblFills")
            With t
                If Not ItemExists(.Fields, "BrokerFillID") Then
                    .Fields.Append .CreateField("BrokerFillID", dbText, 70)
                    .Fields("BrokerFillID").Required = False
                    .Fields("BrokerFillID").AllowZeroLength = True
                    .Fields("BrokerFillID").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblFills SET tblFills.BrokerFillID = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "BuySell") Then
                    .Fields.Append .CreateField("BuySell", dbBoolean)
                    .Fields("BuySell").DefaultValue = False
                End If
            End With
            
            FillTablesForVersion4
            
            ' Create tblOrderHistory...
            If Not ItemExists(m.DB.TableDefs, "tblOrderHistory") Then
                Set t = m.DB.CreateTableDef("tblOrderHistory")
                With t
                    If Not ItemExists(.Fields, "HistoryID") Then
                        .Fields.Append .CreateField("HistoryID", dbLong)
                        .Fields("HistoryID").Attributes = dbAutoIncrField
                    End If
                    If Not ItemExists(.Fields, "OrderID") Then
                        .Fields.Append .CreateField("OrderID", dbLong)
                    End If
                    If Not ItemExists(.Fields, "Seq") Then
                        .Fields.Append .CreateField("Seq", dbLong)
                        .Fields("Seq").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "StatusDate") Then
                        .Fields.Append .CreateField("StatusDate", dbDate)
                    End If
                    If Not ItemExists(.Fields, "Status") Then
                        .Fields.Append .CreateField("Status", dbText)
                        .Fields("Status").Required = False
                        .Fields("Status").AllowZeroLength = True
                        .Fields("Status").DefaultValue = ""
                    End If
                    If Not ItemExists(.Fields, "Message") Then
                        .Fields.Append .CreateField("Message", dbMemo)
                        .Fields("Message").Required = False
                        .Fields("Message").AllowZeroLength = True
                        .Fields("Message").DefaultValue = ""
                    End If
                    
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("OrderID")
                            .Fields.Append .CreateField("HistoryID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                End With
                m.DB.TableDefs.Append t
            End If
        
            If Not ItemExists(m.DB.Relations, "OrderHistory") Then
                Set r = m.DB.CreateRelation("OrderHistory", "tblOrders", "tblOrderHistory", _
                    dbRelationDeleteCascade)
                With r
                    .Fields.Append .CreateField("OrderID")
                    .Fields.Item("OrderID").ForeignName = "OrderID"
                End With
                m.DB.Relations.Append r
            End If
            
            Set t = m.DB.TableDefs("tblFills")
            With t
                If ItemExists(.Indexes, "PrimaryKey") Then
                    .Indexes.Delete "PrimaryKey"
                    Set i = .CreateIndex("PrimaryKey")
                    With i
                        .Fields.Append .CreateField("FillID")
                        .Primary = True
                    End With
                    .Indexes.Append i
                End If
                
            End With
                        
            If ItemExists(m.DB.Relations, "tblPositionstblOrders") Then
                m.DB.Relations.Delete "tblPositionstblOrders"
            End If
            
            ' Create tblPositionFills...
            If Not ItemExists(m.DB.TableDefs, "tblPositionFills") Then
                Set t = m.DB.CreateTableDef("tblPositionFills")
                With t
                    If Not ItemExists(.Fields, "PositionID") Then
                        .Fields.Append .CreateField("PositionID", dbLong)
                    End If
                    If Not ItemExists(.Fields, "FillID") Then
                        .Fields.Append .CreateField("FillID", dbLong)
                    End If
                    If Not ItemExists(.Fields, "Quantity") Then
                        .Fields.Append .CreateField("Quantity", dbLong)
                    End If
                    If Not ItemExists(.Fields, "Position") Then
                        .Fields.Append .CreateField("Position", dbLong)
                    End If
                                        
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("PositionID")
                            .Fields.Append .CreateField("FillID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                    
                    m.DB.TableDefs.Append t
                End With
            End If
        
            If Not ItemExists(m.DB.Relations, "tblPositionstblPositionFills") Then
                Set r = m.DB.CreateRelation("tblPositionstblPositionFills", "tblPositions", _
                    "tblPositionFills", dbRelationDeleteCascade Or dbRelationUpdateCascade)
                With r
                    .Fields.Append .CreateField("PositionID")
                    .Fields.Item("PositionID").ForeignName = "PositionID"
                End With
                
                m.DB.Relations.Append r
            End If
    
            If Not ItemExists(m.DB.Relations, "tblFillstblPositionFills") Then
                Set r = m.DB.CreateRelation("tblFillstblPositionFills", "tblFills", _
                    "tblPositionFills", dbRelationDeleteCascade Or dbRelationUpdateCascade)
                With r
                    .Fields.Append .CreateField("FillID")
                    .Fields.Item("FillID").ForeignName = "FillID"
                End With
                
                m.DB.Relations.Append r
            End If
    
            FillPositionFills
            
            ' Remove PositionID from tblOrders...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If ItemExists(.Indexes, "PositionID") Then
                    .Indexes.Delete "PositionID"
                End If
                If ItemExists(.Indexes, "PrimaryKey") Then
                    .Indexes.Delete "PrimaryKey"
                    Set i = .CreateIndex("PrimaryKey")
                    With i
                        .Fields.Append .CreateField("OrderID")
                        .Primary = True
                    End With
                    .Indexes.Append i
                End If
                If ItemExists(.Fields, "PositionID") Then
                    .Fields.Delete "PositionID"
                End If
            End With
            
        Case 5:
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "Message") Then
                    .Fields.Append .CreateField("Message", dbText, 100)
                        
                    .Fields("Message").Required = False
                    .Fields("Message").AllowZeroLength = True
                    .Fields("Message").DefaultValue = ""
                
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.Message = '';"
                    q.Execute
                End If
            End With
            
        Case 6:
            ChangeFieldSize "tblOrders", "BrokerOrderID", 100
            ChangeFieldSize "tblOrders", "ExchangeOrderID", 100
            
        Case 7:
            ' Create tblSymbolInfo to hold last known information about a symbol/id...
            If Not ItemExists(m.DB.TableDefs, "tblSymbolInfo") Then
                Set t = m.DB.CreateTableDef("tblSymbolInfo")
                With t
                    If Not ItemExists(.Fields, "SymbolID") Then
                        .Fields.Append .CreateField("SymbolID", dbLong)
                        .Fields("SymbolID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "Symbol") Then
                        .Fields.Append .CreateField("Symbol", dbText, 50)
                        .Fields("Symbol").AllowZeroLength = True
                        .Fields("Symbol").Required = False
                        .Fields("Symbol").DefaultValue = ""
                    End If
                    If Not ItemExists(.Fields, "TickMove") Then
                        .Fields.Append .CreateField("TickMove", dbDouble)
                        .Fields("TickMove").DefaultValue = 0#
                    End If
                    If Not ItemExists(.Fields, "TickValue") Then
                        .Fields.Append .CreateField("TickValue", dbDouble)
                        .Fields("TickValue").DefaultValue = 0#
                    End If
                                        
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("SymbolID")
                            .Fields.Append .CreateField("Symbol")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                    
                    m.DB.TableDefs.Append t
                End With
            End If

            ' Add OpenSymbolID to the Positions table...
            Set t = m.DB.TableDefs("tblPositions")
            With t
                If Not ItemExists(.Fields, "OpenSymbolID") Then
                    .Fields.Append .CreateField("OpenSymbolID", dbLong)
                    .Fields("OpenSymbolID").DefaultValue = 0&
                End If
            End With
            
        Case 8:
            ' Add GenesisOrderID field to the orders table...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "GenesisOrderID") Then
                    .Fields.Append .CreateField("GenesisOrderID", dbText, 50)
                    .Fields("GenesisOrderID").Required = False
                    .Fields("GenesisOrderID").AllowZeroLength = True
                    .Fields("GenesisOrderID").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.GenesisOrderID = tblOrders.OrderID;"
                    q.Execute
                End If
            End With
            
        Case 9:
            ' Add AutoTradeItemID field to the orders table...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "AutoTradeItemID") Then
                    .Fields.Append .CreateField("AutoTradeItemID", dbLong)
                    .Fields("AutoTradeItemID").DefaultValue = 0
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.AutoTradeItemID = 0;"
                    q.Execute
                End If
            End With
            
            ' Add AutoTradeItemID field to the positions table...
            Set t = m.DB.TableDefs("tblPositions")
            With t
                If Not ItemExists(.Fields, "AutoTradeItemID") Then
                    .Fields.Append .CreateField("AutoTradeItemID", dbLong)
                    .Fields("AutoTradeItemID").DefaultValue = 0
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblPositions SET tblPositions.AutoTradeItemID = 0;"
                    q.Execute
                End If
            End With
            
        Case 10:
            ' Create Order Journal table to hold journal entries for orders...
            If Not ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
                Set t = m.DB.CreateTableDef("tblOrderJournal")
                With t
                    If Not ItemExists(.Fields, "JournalID") Then
                        .Fields.Append .CreateField("JournalID", dbLong)
                        .Fields("JournalID").Attributes = dbAutoIncrField
                    End If
                    If Not ItemExists(.Fields, "OrderID") Then
                        .Fields.Append .CreateField("OrderID", dbLong)
                        .Fields("OrderID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "NoteDate") Then
                        .Fields.Append .CreateField("NoteDate", dbDate)
                        .Fields("NoteDate").DefaultValue = 0#
                    End If
                    If Not ItemExists(.Fields, "Note") Then
                        .Fields.Append .CreateField("Note", dbMemo)
                        .Fields("Note").AllowZeroLength = True
                        .Fields("Note").Required = False
                        .Fields("Note").DefaultValue = ""
                    End If
                                        
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("JournalID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                    
                    m.DB.TableDefs.Append t
                End With
            End If
        
            If Not ItemExists(m.DB.Relations, "tblOrderstblOrderJournal") Then
                Set r = m.DB.CreateRelation("tblOrderstblOrderJournal", "tblOrders", _
                    "tblOrderJournal", dbRelationDeleteCascade Or dbRelationUpdateCascade)
                With r
                    .Fields.Append .CreateField("OrderID")
                    .Fields.Item("OrderID").ForeignName = "OrderID"
                End With
                
                m.DB.Relations.Append r
            End If
            
        Case 11:
            ' Create Order Groups table and Order Group Items table...
            If Not ItemExists(m.DB.TableDefs, "tblOrderGroups") Then
                Set t = m.DB.CreateTableDef("tblOrderGroups")
                With t
                    If Not ItemExists(.Fields, "OrderGroupID") Then
                        .Fields.Append .CreateField("OrderGroupID", dbLong)
                        .Fields("OrderGroupID").Attributes = dbAutoIncrField
                    End If
                    If Not ItemExists(.Fields, "Name") Then
                        .Fields.Append .CreateField("Name", dbText, 50)
                        .Fields("Name").AllowZeroLength = True
                        .Fields("Name").Required = False
                        .Fields("Name").DefaultValue = ""
                    End If
                    If Not ItemExists(.Fields, "Description") Then
                        .Fields.Append .CreateField("Description", dbText, 255)
                        .Fields("Description").AllowZeroLength = True
                        .Fields("Description").Required = False
                        .Fields("Description").DefaultValue = ""
                    End If
                
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("OrderGroupID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                    
                    m.DB.TableDefs.Append t
                End With
            End If
            
            If Not ItemExists(m.DB.TableDefs, "tblOrderGroupItems") Then
                Set t = m.DB.CreateTableDef("tblOrderGroupItems")
                With t
                    If Not ItemExists(.Fields, "OrderGroupItemID") Then
                        .Fields.Append .CreateField("OrderGroupItemID", dbLong)
                        .Fields("OrderGroupItemID").Attributes = dbAutoIncrField
                    End If
                    If Not ItemExists(.Fields, "OrderGroupID") Then
                        .Fields.Append .CreateField("OrderGroupID", dbLong)
                        .Fields("OrderGroupID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "OrderText") Then
                        .Fields.Append .CreateField("OrderText", dbText, 255)
                        .Fields("OrderText").AllowZeroLength = True
                        .Fields("OrderText").Required = False
                        .Fields("OrderText").DefaultValue = ""
                    End If
                
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("OrderGroupItemID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                    
                    m.DB.TableDefs.Append t
                End With
            End If
    
            If Not ItemExists(m.DB.Relations, "tblOrderGroupstblOrderGroupItems") Then
                Set r = m.DB.CreateRelation("tblOrderGroupstblOrderGroupItems", "tblOrderGroups", _
                    "tblOrderGroupItems", dbRelationDeleteCascade Or dbRelationUpdateCascade)
                With r
                    .Fields.Append .CreateField("OrderGroupID")
                    .Fields.Item("OrderGroupID").ForeignName = "OrderGroupID"
                End With
                
                m.DB.Relations.Append r
            End If
            
        Case 12:
            ' Add the Security Type bitmask to the Accounts table...
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                If Not ItemExists(.Fields, "SecTypeMask") Then
                    .Fields.Append .CreateField("SecTypeMask", dbLong)
                    .Fields("SecTypeMask").DefaultValue = 0&
                    
                    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
                    Do While Not rs.EOF
                        rs.Edit
                        Select Case rs!AccountType
                            Case eTT_AccountType_SimTrade, eTT_AccountType_Standard
                                rs!SecTypeMask = 7
                            Case Else
                                rs!SecTypeMask = 1
                        End Select
                        rs.Update
                        
                        rs.MoveNext
                    Loop
                End If
            End With
            
        Case 13:
            ' Add fields to the order table for Conditional/Trigger/OCO/Trail orders...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "CancelOrderID") Then
                    .Fields.Append .CreateField("CancelOrderID", dbLong)
                    .Fields("CancelOrderID").DefaultValue = 0&
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.CancelOrderID = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "TriggerOrderID") Then
                    .Fields.Append .CreateField("TriggerOrderID", dbLong)
                    .Fields("TriggerOrderID").DefaultValue = 0&
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.TriggerOrderID = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "TriggerOptions") Then
                    .Fields.Append .CreateField("TriggerOptions", dbText, 255)
                    .Fields("TriggerOptions").AllowZeroLength = True
                    .Fields("TriggerOptions").Required = False
                    .Fields("TriggerOptions").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.TriggerOptions = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "ConditionOptions") Then
                    .Fields.Append .CreateField("ConditionOptions", dbMemo)
                    .Fields("ConditionOptions").AllowZeroLength = True
                    .Fields("ConditionOptions").Required = False
                    .Fields("ConditionOptions").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.ConditionOptions = '';"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "TrailAmount") Then
                    .Fields.Append .CreateField("TrailAmount", dbDouble)
                    .Fields("TrailAmount").DefaultValue = 0#
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.TrailAmount = 0;"
                    q.Execute
                End If
                If Not ItemExists(.Fields, "TrailOptions") Then
                    .Fields.Append .CreateField("TrailOptions", dbText, 255)
                    .Fields("TrailOptions").AllowZeroLength = True
                    .Fields("TrailOptions").Required = False
                    .Fields("TrailOptions").DefaultValue = ""
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.TrailOptions = '';"
                    q.Execute
                End If
            End With
            
        Case 14:
            ' Add the ExitPos field to the Orders table...
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Fields, "ExitPos") Then
                    .Fields.Append .CreateField("ExitPos", dbLong)
                    .Fields("ExitPos").DefaultValue = 0&
                    
                    Set q = m.DB.CreateQueryDef("")
                    q.SQL = "UPDATE tblOrders SET tblOrders.ExitPos = 0;"
                    q.Execute
                End If
            End With
            
        Case 15:
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                If Not ItemExists(.Fields, "FillRT") Then
                    .Fields.Append .CreateField("FillRT", dbBoolean)
                    .Fields("FillRT").DefaultValue = True
                    
                    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
                    Do While Not rs.EOF
                        rs.Edit
                        If (rs!AccountType = eTT_AccountType_SimTrade) Or (rs!AccountType = eTT_AccountType_Standard) Then
                            rs!FillRT = True
                        Else
                            rs!FillRT = False
                        End If
                        rs.Update
                        
                        rs.MoveNext
                    Loop
                End If
            End With
            
        Case 16:
            Set t = m.DB.TableDefs("tblPositions")
            With t
                If Not ItemExists(.Fields, "ClosedProfit") Then
                    .Fields.Append .CreateField("ClosedProfit", dbDouble)
                    .Fields("ClosedProfit").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "AverageEntry") Then
                    .Fields.Append .CreateField("AverageEntry", dbDouble)
                    .Fields("AverageEntry").DefaultValue = 0#
                End If
            End With
            
            Set t = m.DB.TableDefs("tblActivities")
            With t
                If Not ItemExists(.Fields, "ClosedProfit") Then
                    .Fields.Append .CreateField("ClosedProfit", dbDouble)
                    .Fields("ClosedProfit").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "IsOpen") Then
                    .Fields.Append .CreateField("IsOpen", dbBoolean)
                    .Fields("IsOpen").DefaultValue = False
                End If
            End With
            
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                If Not ItemExists(.Fields, "ClosedProfit") Then
                    .Fields.Append .CreateField("ClosedProfit", dbDouble)
                    .Fields("ClosedProfit").DefaultValue = 0#
                End If
            End With
            
        Case 17:
            Set t = m.DB.TableDefs("tblFills")
            With t
                If Not ItemExists(.Fields, "ClosedProfit") Then
                    .Fields.Append .CreateField("ClosedProfit", dbDouble)
                    .Fields("ClosedProfit").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "CurrentPosition") Then
                    .Fields.Append .CreateField("CurrentPosition", dbLong)
                    .Fields("CurrentPosition").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "SessionDate") Then
                    .Fields.Append .CreateField("SessionDate", dbDouble)
                    .Fields("SessionDate").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "NumOpen") Then
                    .Fields.Append .CreateField("NumOpen", dbLong)
                    .Fields("NumOpen").DefaultValue = 0&
                End If
                If Not ItemExists(.Fields, "SessionQuantity") Then
                    .Fields.Append .CreateField("SessionQuantity", dbLong)
                    .Fields("SessionQuantity").DefaultValue = 0&
                End If
                If Not ItemExists(.Fields, "SessionProfit") Then
                    .Fields.Append .CreateField("SessionProfit", dbDouble)
                    .Fields("SessionProfit").DefaultValue = 0#
                End If
                
                If Not ItemExists(.Indexes, "SessionDate") Then
                    Set i = .CreateIndex("SessionDate")
                    With i
                        .Fields.Append .CreateField("SessionDate")
                        .Primary = False
                    End With
                    .Indexes.Append i
                End If
                If Not ItemExists(.Indexes, "NumOpen") Then
                    Set i = .CreateIndex("NumOpen")
                    With i
                        .Fields.Append .CreateField("NumOpen")
                        .Primary = False
                    End With
                    .Indexes.Append i
                End If
            End With
            
        Case 18:
            Set t = m.DB.TableDefs("tblAccounts")
            With t
                If Not ItemExists(.Fields, "TotalFees") Then
                    .Fields.Append .CreateField("TotalFees", dbDouble)
                    .Fields("TotalFees").DefaultValue = 0#
                End If
                If Not ItemExists(.Fields, "TotalAdjustments") Then
                    .Fields.Append .CreateField("TotalAdjustments", dbDouble)
                    .Fields("TotalAdjustments").DefaultValue = 0#
                End If
            
                If Not ItemExists(.Indexes, "PrimaryKey") Then
                    Set i = .CreateIndex("PrimaryKey")
                    With i
                        .Fields.Append .CreateField("AccountID")
                        .Primary = True
                    End With
                    .Indexes.Append i
                End If
            End With
            
            If Not ItemExists(m.DB.TableDefs, "tblAccountPositions") Then
                Set t = m.DB.CreateTableDef("tblAccountPositions")
                With t
                    If Not ItemExists(.Fields, "AccountID") Then
                        .Fields.Append .CreateField("AccountID", dbLong)
                        .Fields("AccountID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "Symbol") Then
                        .Fields.Append .CreateField("Symbol", dbText, 50)
                        .Fields("Symbol").DefaultValue = ""
                        .Fields("Symbol").Required = False
                        .Fields("Symbol").AllowZeroLength = True
                    End If
                    If Not ItemExists(.Fields, "SymbolID") Then
                        .Fields.Append .CreateField("SymbolID", dbLong)
                        .Fields("SymbolID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "AutoTradeItemID") Then
                        .Fields.Append .CreateField("AutoTradeItemID", dbLong)
                        .Fields("AutoTradeItemID").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "CurrentPosition") Then
                        .Fields.Append .CreateField("CurrentPosition", dbLong)
                        .Fields("CurrentPosition").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "SessionDate") Then
                        .Fields.Append .CreateField("SessionDate", dbLong)
                        .Fields("SessionDate").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "SessionQuantity") Then
                        .Fields.Append .CreateField("SessionQuantity", dbLong)
                        .Fields("SessionQuantity").DefaultValue = 0&
                    End If
                    If Not ItemExists(.Fields, "SessionProfit") Then
                        .Fields.Append .CreateField("SessionProfit", dbDouble)
                        .Fields("AccountID").DefaultValue = 0#
                    End If
                    If Not ItemExists(.Fields, "AverageEntry") Then
                        .Fields.Append .CreateField("AverageEntry", dbDouble)
                        .Fields("AccountID").DefaultValue = 0#
                    End If
                    If Not ItemExists(.Fields, "LastTraded") Then
                        .Fields.Append .CreateField("LastTraded", dbDouble)
                        .Fields("LastTraded").DefaultValue = 0#
                    End If
                
                    If Not ItemExists(.Indexes, "Primary") Then
                        Set i = .CreateIndex("Primary")
                        With i
                            .Fields.Append .CreateField("AccountID")
                            .Fields.Append .CreateField("Symbol")
                            .Fields.Append .CreateField("SymbolID")
                            .Fields.Append .CreateField("AutoTradeItemID")
                            .Primary = True
                        End With
                        .Indexes.Append i
                    End If
                End With
            
                m.DB.TableDefs.Append t
            End If
            
            If Not ItemExists(m.DB.Relations, "tblAccountstblAccountPositions") Then
                Set r = m.DB.CreateRelation("tblAccountstblAccountPositions", "tblAccounts", _
                    "tblAccountPositions", dbRelationDeleteCascade Or dbRelationUpdateCascade)
                With r
                    .Fields.Append .CreateField("AccountID")
                    .Fields.Item("AccountID").ForeignName = "AccountID"
                End With
                
                m.DB.Relations.Append r
            End If
            
            FillTablesForVersion18
            
        Case 19:
            ApplyVersion19
            
        Case 20:
            ' Clean out the account positions table to force the fill information to be
            ' regenerated because of a bug with NumOpen (03/20/2006 - DAJ)...
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Delete
                rs.MoveNext
            Loop
            
        Case 21:
            Set t = m.DB.TableDefs("tblOrders")
            With t
                If Not ItemExists(.Indexes, "GenesisLookup") Then
                    Set i = .CreateIndex("GenesisLookup")
                    With i
                        .Fields.Append .CreateField("GenesisOrderID")
                        .Fields.Append .CreateField("AccountID")
                        .Primary = False
                    End With
                    .Indexes.Append i
                End If
                If Not ItemExists(.Indexes, "BrokerLookup") Then
                    Set i = .CreateIndex("BrokerLookup")
                    With i
                        .Fields.Append .CreateField("BrokerOrderID")
                        .Fields.Append .CreateField("AccountID")
                        .Primary = False
                    End With
                    .Indexes.Append i
                End If
                If Not ItemExists(.Indexes, "TriggerLookup") Then
                    Set i = .CreateIndex("TriggerLookup")
                    With i
                        .Fields.Append .CreateField("TriggerOrderID")
                        .Primary = False
                    End With
                    .Indexes.Append i
                End If
            End With
        
        Case 22:
            CleanOutTransActAccounts
            
        Case 23:
            ApplyVersion23
            
        Case 24:
            ApplyVersion24
            
        Case 25:
            ApplyVersion25
            
        Case 26:
            ApplyVersion26
            
        Case 27:
            ' Change the security type mask for PFG accounts from just Futures to Futures and Forex...
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] " & _
                    "WHERE [AccountType]=" & Str(10) & ";", dbOpenDynaset)
'                    "WHERE [AccountType]=" & Str(eTT_AccountType_PFG) & ";", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Edit
                rs!SecTypeMask = 5
                rs.Update
    
                rs.MoveNext
            Loop
            
        Case 28:
            ApplyVersion28
            
        Case 29:
            FillVersion29
            
        Case 30:
            ApplyVersion30
            
        Case 31:
            ApplyVersion31
            
        Case 32:
            ApplyVersion32
            
        Case 33:
            ApplyVersion33
            
        Case 34:
            ApplyVersion34
            
        Case 35:
            ApplyVersion35
            
        ' Version 36: Change the length of the message field in the orders table to the maximum 255 characters...
        Case 36:
            ChangeFieldSize "tblOrders", "Message", 255
            
        ' Version 37: Change the BrokerFillID for PFG accounts to BrokerOrderID-BrokerFillID...
        Case 37:
            ApplyVersion37
            
        ' Version 38: Additions to OrderJournal table (Entry, Feeling, Why Making Trade, Thoughts)...
        Case 38:
            ApplyVersion38
            
        ' Version 39: Add the OrderLegs table, populate it, then remove the leg fields from the orders table...
        Case 39:
            ApplyVersion39
            
        ' Version 40: Add the Group ID and Group Name fields to the orders table...
        Case 40:
            ApplyVersion40
            
        ' Version 41: Add Max Profit and Max Loss to the account position trades table...
        Case 41:
            ApplyVersion41
            
        ' Version 42: Add EmotionNumber field to the journal table...
        Case 42:
            ApplyVersion42
            
        ' Version 43: Add AccountID to the order journal table...
        Case 43:
            ApplyVersion43
            
        ' Version 44: Change security type masks to accomodate options...
        Case 44:
            ApplyVersion44
            
        ' Version 45: Change security type masks for PFG accounts to turn on Stocks and Stock Options...
        Case 45:
            ApplyVersion45
            
        ' Version 46: Add the IsAutoExit flag to the orders table...
        Case 46:
            ApplyVersion46
            
        ' Version 47: Change any orders with status "Open" to status "Parked"
        Case 47:
            ApplyVersion47
            
        ' Version 48: Add the BrokerCancelOrderID field to the orders table...
        Case 48:
            ApplyVersion48
            
        ' Version 49: Add the Fees to tblAccountPositionTrades and StockFees to tblAccounts...
        Case 49:
            ApplyVersion49
            
        ' Version 50: Add the Stop2OrderID and Stop3OrderID fields to tblActiveExit...
        Case 50:
            ApplyVersion50
            
        ' Version 51: Change the default value for the fees on fills to kNullData...
        Case 51:
            ApplyVersion51
            
        ' Version 52: Add the IsAutomated flag in the orders table...
        Case 52:
            ApplyVersion52
            
        ' Version 53: Convert SimTrade accounts over to SimBroker, SimStream, or SimReplay...
        Case 53:
            ApplyVersion53
            
        ' Version 54: Add table for date journals...
        Case 54:
            ApplyVersion54
            
        ' Version 55: Add image filename field to the date journals and order journals tables...
        Case 55:
            ApplyVersion55
            
        ' Version 56: Change date journals text field to memo, add indexes for journal date fields...
        Case 56:
            ApplyVersion56
            
        ' Version 57: Fix Broker ID's on orders and fills for Interactive Brokers and I-Deal
        Case 57:
            ApplyVersion57
            
        ' Version 58: Add expiration date to the account positions table...
        Case 58:
            ApplyVersion58
            
        ' Version 59: Add the journal images table and move data around...
        Case 59:
            ApplyVersion59
            
        ' Version 60: Move fields from the Order Legs table up to the Orders table...
        Case 60:
            ApplyVersion60
            
        ' Version 61: Add symbol/symbol id fields to the two journal tables...
        Case 61:
            ApplyVersion61
            
        ' Version 62: Add the symbol info table to hold expiration date and multiplier...
        Case 62:
            ApplyVersion62
            
        ' Version 63: Add 'FcmAccountNumber' to the accounts table...
        Case 63:
            ApplyVersion63
            
        ' Version 64: Add 'ContingencyOptions' to the orders table...
        Case 64:
            ApplyVersion64
            
        ' Version 65: Add 'OcaGroup' to the orders table...
        Case 65:
            ApplyVersion65
            
        ' Version 66: Add 'IsSpread' to the fills table...
        Case 66:
            ApplyVersion66
            
    End Select

ErrExit:
    Set q = Nothing
    Set i = Nothing
    Set r = Nothing
    Set t = Nothing
    Exit Sub
    
ErrSection:
    Set q = Nothing
    Set i = Nothing
    Set r = Nothing
    Set t = Nothing
    RaiseError "cTTUpdates.ApplyVersion", eGDRaiseError_Raise

End Sub

Private Sub FillTablesForVersion4()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim lPosition As Long
    Dim lCurPos As Long
    
    ' Set the AdjustAmount in tblActivities...
    Set rs = m.DB.OpenRecordset("SELECT tblActivities.*, tblFills.Fees " & _
                "FROM (tblActivities INNER JOIN tblPositions ON tblActivities.ActivityID = tblPositions.ActivityID) INNER JOIN (tblOrders INNER JOIN tblFills ON tblOrders.OrderID = tblFills.OrderID) ON tblPositions.PositionID = tblOrders.PositionID " & _
                "WHERE (((tblActivities.ActivityType)=1));", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!AdjustAmount = rs!Fees
        rs.Update
        
        ' Delete any Positions that are adjustments since all of the info will
        ' now be stored in the Activity table...
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblPositions] " & _
                "WHERE [ActivityID]=" & Str(rs!ActivityID) & ";", dbOpenDynaset)
        If Not (rs2.EOF And rs2.BOF) Then
            rs2.Delete
        End If
        
        rs.MoveNext
    Loop
    
    ' Set BuySell in tblFills...
    Set rs = m.DB.OpenRecordset("SELECT tblFills.*, tblOrders.BuySell AS OrderBuy " & _
                "FROM tblOrders INNER JOIN tblFills ON tblOrders.OrderID = tblFills.OrderID;", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!BuySell = rs!OrderBuy
        rs.Update
        
        rs.MoveNext
    Loop
    
    ' Set Position and Open in tblPositions...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        lPosition = 0
        lCurPos = 0&
        
        Set rs2 = m.DB.OpenRecordset("SELECT tblPositions.PositionID, tblOrders.EnterExit, tblFills.* " & _
                "FROM (tblPositions INNER JOIN tblOrders ON tblPositions.PositionID = tblOrders.PositionID) INNER JOIN tblFills ON tblOrders.OrderID = tblFills.OrderID " & _
                "WHERE tblPositions.PositionID=" & Str(rs!PositionID) & ";", dbOpenDynaset)
        Do While Not rs2.EOF
            If rs2!BuySell = True Then
                If rs2!EnterExit = True Then lPosition = 1
                lCurPos = lCurPos + rs2!FillQuantity
            Else
                If rs2!EnterExit = True Then lPosition = 2
                lCurPos = lCurPos - rs2!FillQuantity
            End If
            
            rs2.MoveNext
        Loop
        
        rs.Edit
        If Val(rs!SecurityType) <> 0 Then rs!SecurityType = Chr(rs!SecurityType)
        rs!Position = lPosition
        rs!Open = (lCurPos <> 0)
        rs.Update
        
        rs.MoveNext
    Loop
    
    ' Set Symbol, SymbolID, and AccountID in tblOrders...
    Set rs = m.DB.OpenRecordset("SELECT tblActivities.AccountID AS AccID, tblPositions.Symbol AS PosSym, tblPositions.SymbolID AS PosSymID, tblOrders.* " & _
                "FROM (tblActivities INNER JOIN tblPositions ON tblActivities.ActivityID = tblPositions.ActivityID) INNER JOIN tblOrders ON tblPositions.PositionID = tblOrders.PositionID;", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!Symbol = rs!PosSym
        rs!SymbolID = rs!PosSymId
        rs!AccountID = rs!AccID
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set rs2 = Nothing
    RaiseError "cTTUpdates.FillTablesForVersion4", eGDRaiseError_Raise
    
End Sub
    
Private Sub FillPositionFills()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim rs2 As Recordset

    Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblPositionFills];", dbOpenDynaset)
    Set rs = m.DB.OpenRecordset("SELECT tblPositions.PositionID, tblFills.FillID, tblFills.FillQuantity, tblOrders.BuySell " & _
                "FROM tblPositions INNER JOIN (tblOrders INNER JOIN tblFills ON tblOrders.OrderID = tblFills.OrderID) ON tblPositions.PositionID = tblOrders.PositionID;", dbOpenDynaset)
    Do While Not rs.EOF
        rs2.AddNew
        rs2!PositionID = rs!PositionID
        rs2!FillID = rs!FillID
        rs2!Quantity = rs!FillQuantity
        If rs!BuySell = True Then
            rs2!Position = 1
        Else
            rs2!Position = 2
        End If
        rs2.Update
        
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set rs2 = Nothing
    RaiseError "cTTUpdates.FillPositionFills", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeFieldSize(ByVal strTable As String, ByVal strField As String, ByVal lNewSize As Long)
On Error GoTo ErrSection:

    Dim t As TableDef
    Dim fld As Field
    Dim q As QueryDef
    Dim rs As Recordset
    Dim mb As New cMemBuffer

    Set t = m.DB.TableDefs(strTable)
    With t
        If .Fields(strField).Size <> lNewSize Then
            ' create a new field of the correct data type
            ' and insert into table after field being replaced
            Set fld = .CreateField("temp_new", .Fields(strField).Type, lNewSize)
            fld.OrdinalPosition = .Fields(strField).OrdinalPosition
            .Fields.Append fld
            
            ' duplicate some of the properties
            On Error Resume Next
            With .Fields(strField)
                fld.DefaultValue = .DefaultValue
                fld.Required = .Required
                fld.AllowZeroLength = True ' .AllowZeroLength
            End With
            On Error GoTo ErrSection:
            
            Set q = m.DB.CreateQueryDef("") '(temp update query)
            q.SQL = "UPDATE " & strTable & " SET [" & strTable & "].[temp_new] = [" & strTable & "].[" & strField & "];"
            q.Execute
            
            ' delete the old field and rename the new field
            .Fields.Delete strField
            .Fields("temp_new").Name = strField
        End If
    End With

ErrExit:
    Set q = Nothing
    Set fld = Nothing
    Set t = Nothing
    Exit Sub
    
ErrSection:
    Set q = Nothing
    Set fld = Nothing
    Set t = Nothing
    RaiseError "cTTUpdates.ChangeFieldSize", eGDRaiseError_Raise

End Sub

Private Sub FillTablesForVersion18()
On Error GoTo ErrSection:

    Dim rs1 As Recordset                ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim dTotalFees As Double            ' Total fees
    Dim dTotalAdjustments As Double     ' Total adjustments
    
    Set rs1 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    Do While Not rs1.EOF
        dTotalFees = 0#
        dTotalAdjustments = 0#
        
        Set rs2 = m.DB.OpenRecordset("SELECT tblFills.* " & _
                    "FROM tblFills INNER JOIN tblOrders ON tblFills.OrderID=tblOrders.OrderID " & _
                    "WHERE tblOrders.AccountID=" & Str(rs1!AccountID) & ";", dbOpenDynaset)
        Do While Not rs2.EOF
            dTotalFees = dTotalFees + rs2!Fees
            rs2.MoveNext
        Loop
        
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblActivities] " & _
                    "WHERE [AccountID]=" & Str(rs1!AccountID) & " AND [ActivityType]=1;", dbOpenDynaset)
        Do While Not rs2.EOF
            dTotalAdjustments = dTotalAdjustments + rs2!AdjustAmount
            rs2.MoveNext
        Loop
        
        rs1.Edit
        rs1!TotalFees = dTotalFees
        rs1!TotalAdjustments = dTotalAdjustments
        rs1.Update
                
        rs1.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.FillTablesForVersion18"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion19
'' Description: Apply version 19 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion19()
On Error GoTo ErrSection:

    Dim t As TableDef
    Dim rs As Recordset
    Dim astrPasswords As New cGdArray
    Dim lIndex As Long

    astrPasswords.Create eGDARRAY_Strings

    Set t = m.DB.TableDefs("tblAccounts")
    With t
        If .Fields("Password").Type = dbText Then
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
            Do While Not rs.EOF
                astrPasswords.Add rs!Password
                rs.MoveNext
            Loop
            Set rs = Nothing
        
            ChangeDataType "tblAccounts", "Password", dbLongBinary
            
            lIndex = 0&
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Edit
                EncryptField rs!Password, astrPasswords(lIndex)
                rs.Update
                
                lIndex = lIndex + 1
                rs.MoveNext
            Loop
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion19"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CleanOutTransActAccounts
'' Description: Clean out everything out of any TransAct accounts
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CleanOutTransActAccounts()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
        
    ' Clean out all orders and fills for any TransAct accounts...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(rs!AccountID) & ";", dbOpenDynaset)
        If Not (rs2.EOF And rs2.BOF) Then
            If (rs2!AccountType = eTT_AccountType_TransAct) Then
                rs.Delete
            End If
        End If
        
        rs.MoveNext
    Loop
    
    ' Clean out any activities for any TransAct accounts...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblActivities];", dbOpenDynaset)
    Do While Not rs.EOF
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(rs!AccountID) & ";", dbOpenDynaset)
        If Not (rs2.EOF And rs2.BOF) Then
            If (rs2!AccountType = eTT_AccountType_TransAct) Then
                rs.Delete
            End If
        End If
        
        rs.MoveNext
    Loop
    
    ' Clean out any account positions for any TransAct accounts...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(rs!AccountID) & ";", dbOpenDynaset)
        If Not (rs2.EOF And rs2.BOF) Then
            If (rs2!AccountType = eTT_AccountType_TransAct) Then
                rs.Delete
            End If
        End If
        
        rs.MoveNext
    Loop
    
    ' Start the TransAct accounts over from scratch...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    Do While Not rs.EOF
        If (rs!AccountType = eTT_AccountType_TransAct) Then
            rs.Edit
            rs!CurrentBalance = rs!StartingBalance
            rs!ClosedProfit = 0#
            rs!TotalFees = 0#
            rs!TotalAdjustments = 0#
            rs.Update
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.CleanOutTransActAccounts"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion23
'' Description: Apply version 23 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion23()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim rs3 As Recordset                ' Recordset into the database
    Dim rs4 As Recordset                ' Recordset into the database
    Dim lAccountID As Long              ' Account ID

    ' Clean out any orders for account ID's that no longer exist...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        lAccountID = -1&
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(rs!AccountID) & ";", dbOpenDynaset)
        If (rs2.BOF And rs2.EOF) Then
            Set rs3 = m.DB.OpenRecordset("SELECT * FROM [tblFills] WHERE [OrderID]=" & Str(rs!OrderID) & ";", dbOpenDynaset)
            If Not (rs3.BOF And rs3.EOF) Then
                Set rs4 = m.DB.OpenRecordset("SELECT tblActivities.AccountID " & _
                        "FROM (tblActivities INNER JOIN tblPositions ON tblActivities.ActivityID=tblPositions.ActivityID) INNER JOIN tblPositionFills ON tblPositions.PositionID=tblPositionFills.PositionID " & _
                        "WHERE tblPositionFills.FillID=" & Str(rs3!FillID) & ";", dbOpenDynaset)
                If Not (rs4.BOF And rs4.EOF) Then
                    lAccountID = rs4!AccountID
                End If
            End If
            
            If lAccountID = -1& Then
                rs.Delete
            Else
                rs.Edit
                rs!AccountID = lAccountID
                rs.Update
            End If
        End If
        
        rs.MoveNext
    Loop

    ' Clean out the account positions table to force the fill information to be
    ' regenerated because of a bug with NumOpen (03/20/2006 - DAJ)...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion23"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion24
'' Description: Apply version 24 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion24()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim bOpen As Boolean                ' Is the activity open?
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblActivities];", dbOpenDynaset)
    Do While Not rs.EOF
        bOpen = False
        
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblPositions] " & _
                    "WHERE [ActivityID]=" & Str(rs!ActivityID) & ";", dbOpenDynaset)
        Do While Not rs2.EOF
            If rs2!Open <> 0 Then
                bOpen = True
                Exit Do
            End If
            
            rs2.MoveNext
        Loop
        
        rs.Edit
        rs!IsOpen = bOpen
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion24"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion25
'' Description: Apply version 25 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion25()
On Error GoTo ErrSection:

    Dim q As QueryDef
    Dim i As Index
    Dim r As Relation
    Dim t As TableDef
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim lID As Long
    Dim lIndex As Long

    ' Create the Active Exits table...
    If Not ItemExists(m.DB.TableDefs, "tblActiveExits") Then
        Set t = m.DB.CreateTableDef("tblActiveExits")
        With t
            If Not ItemExists(.Fields, "AccountID") Then
                .Fields.Append .CreateField("AccountID", dbLong)
                .Fields("AccountID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").DefaultValue = ""
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
            End If
            If Not ItemExists(.Fields, "ExitStrategyFile") Then
                .Fields.Append .CreateField("ExitStrategyFile", dbText, 255)
                .Fields("ExitStrategyFile").DefaultValue = ""
                .Fields("ExitStrategyFile").AllowZeroLength = True
                .Fields("ExitStrategyFile").Required = False
            End If
            If Not ItemExists(.Fields, "Active") Then
                .Fields.Append .CreateField("Active", dbBoolean)
                .Fields("Active").DefaultValue = 0
            End If
            If Not ItemExists(.Fields, "CurrentPosition") Then
                .Fields.Append .CreateField("CurrentPosition", dbDouble)
                .Fields("CurrentPosition").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "EntryPrice") Then
                .Fields.Append .CreateField("EntryPrice", dbDouble)
                .Fields("EntryPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "Target1OrderID") Then
                .Fields.Append .CreateField("Target1OrderID", dbLong)
                .Fields("Target1OrderID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Target2OrderID") Then
                .Fields.Append .CreateField("Target2OrderID", dbLong)
                .Fields("Target2OrderID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Target3OrderID") Then
                .Fields.Append .CreateField("Target3OrderID", dbLong)
                .Fields("Target3OrderID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "StopOrderID") Then
                .Fields.Append .CreateField("StopOrderID", dbLong)
                .Fields("StopOrderID").DefaultValue = 0&
            End If
            
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AccountID")
                    .Fields.Append .CreateField("SymbolID")
                    .Fields.Append .CreateField("Symbol")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If
    
    ' Delete all of the records from the Account Positions table, remove all relationships to it,
    ' then delete it so that we can recreate it...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    Set rs = Nothing
    For lIndex = m.DB.Relations.Count - 1 To 0 Step -1
        If (m.DB.Relations(lIndex).Table = "tblAccountPositions") Or (m.DB.Relations(lIndex).ForeignTable = "tblAccountPositions") Then
            m.DB.Relations.Delete m.DB.Relations(lIndex).Name
        End If
    Next lIndex
    m.DB.TableDefs.Delete "tblAccountPositions"
    
    If Not ItemExists(m.DB.TableDefs, "tblAccountPositions") Then
        Set t = m.DB.CreateTableDef("tblAccountPositions")
        With t
            If Not ItemExists(.Fields, "AccountPositionID") Then
                .Fields.Append .CreateField("AccountPositionID", dbLong)
                .Fields("AccountPositionID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "AccountID") Then
                .Fields.Append .CreateField("AccountID", dbLong)
                .Fields("AccountID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").DefaultValue = ""
                .Fields("Symbol").Required = False
                .Fields("Symbol").AllowZeroLength = True
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "AutoTradeItemID") Then
                .Fields.Append .CreateField("AutoTradeItemID", dbLong)
                .Fields("AutoTradeItemID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "CurrentPosition") Then
                .Fields.Append .CreateField("CurrentPosition", dbLong)
                .Fields("CurrentPosition").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SessionDate") Then
                .Fields.Append .CreateField("SessionDate", dbLong)
                .Fields("SessionDate").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SessionQuantity") Then
                .Fields.Append .CreateField("SessionQuantity", dbLong)
                .Fields("SessionQuantity").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SessionProfit") Then
                .Fields.Append .CreateField("SessionProfit", dbDouble)
                .Fields("SessionProfit").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "AverageEntry") Then
                .Fields.Append .CreateField("AverageEntry", dbDouble)
                .Fields("AverageEntry").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "LastTraded") Then
                .Fields.Append .CreateField("LastTraded", dbDouble)
                .Fields("LastTraded").DefaultValue = 0#
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AccountPositionID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
    
        m.DB.TableDefs.Append t
    End If
    
    If Not ItemExists(m.DB.Relations, "tblAccountstblAccountPositions") Then
        Set r = m.DB.CreateRelation("tblAccountstblAccountPositions", "tblAccounts", _
            "tblAccountPositions", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountID")
            .Fields.Item("AccountID").ForeignName = "AccountID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    ' Make sure that the tblAccountPositionFills table exists...
    If Not ItemExists(m.DB.TableDefs, "tblAccountPositionFills") Then
        Set t = m.DB.CreateTableDef("tblAccountPositionFills")
        With t
            If Not ItemExists(.Fields, "AccountPositionID") Then
                .Fields.Append .CreateField("AccountPositionID", dbLong)
                .Fields("AccountPositionID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillID") Then
                .Fields.Append .CreateField("FillID", dbLong)
                .Fields("FillID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillDate") Then
                .Fields.Append .CreateField("FillDate", dbDouble)
                .Fields("FillDate").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "Flat") Then
                .Fields.Append .CreateField("Flat", dbBoolean)
                .Fields("Flat").DefaultValue = False
            End If
            If Not ItemExists(.Fields, "Reverse") Then
                .Fields.Append .CreateField("Reverse", dbBoolean)
                .Fields("Reverse").DefaultValue = False
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AccountPositionID")
                    .Fields.Append .CreateField("FillID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            If Not ItemExists(.Indexes, "Lookup") Then
                Set i = .CreateIndex("Lookup")
                With i
                    .Fields.Append .CreateField("AccountPositionID")
                    .Fields.Append .CreateField("FillID")
                    .Fields.Append .CreateField("FillDate")
                    .Primary = False
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If

    If Not ItemExists(m.DB.Relations, "tblAccountPositionstblAccountPositionFills") Then
        Set r = m.DB.CreateRelation("tblAccountPositionstblAccountPositionFills", "tblAccountPositions", _
            "tblAccountPositionFills", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountPositionID")
            .Fields.Item("AccountPositionID").ForeignName = "AccountPositionID"
        End With
        
        m.DB.Relations.Append r
    End If

    If Not ItemExists(m.DB.Relations, "tblFillstblAccountPositionFills") Then
        Set r = m.DB.CreateRelation("tblFillstblAccountPositionFills", "tblFills", _
            "tblAccountPositionFills", dbRelationUnique Or dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("FillID")
            .Fields.Item("FillID").ForeignName = "FillID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    Set t = m.DB.TableDefs("tblFills")
    With t
        If Not ItemExists(.Fields, "AverageEntry") Then
            .Fields.Append .CreateField("AverageEntry", dbDouble)
            .Fields("AverageEntry").DefaultValue = 0#
        End If
    End With
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM " & _
                "tblFills INNER JOIN tblOrders ON tblFills.OrderID=tblOrders.OrderID;", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SymbolID <> 0 Then
            Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                        "WHERE [AccountID]=" & Str(rs!AccountID) & " AND [SymbolID]=" & Str(rs!SymbolID) & " AND [AutoTradeItemID]=" & Str(rs!AutoTradeItemID) & ";", dbOpenDynaset)
        Else
            Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                        "WHERE [AccountID]=" & Str(rs!AccountID) & " AND [Symbol]='" & rs!Symbol & "' AND [AutoTradeItemID]=" & Str(rs!AutoTradeItemID) & ";", dbOpenDynaset)
        End If
        If (rs2.BOF And rs2.EOF) Then
            rs2.AddNew
            rs2!AccountID = rs!AccountID
            rs2!SymbolID = rs!SymbolID
            rs2!Symbol = rs!Symbol
            rs2!AutoTradeItemID = rs!AutoTradeItemID
            lID = rs2!AccountPositionID
            rs2.Update
        Else
            lID = rs2!AccountPositionID
        End If
        
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositionFills] " & _
                    "WHERE [AccountPositionID]=" & Str(lID) & " AND [FillID]=" & Str(rs!FillID) & ";", dbOpenDynaset)
        If (rs2.BOF And rs2.EOF) Then
            rs2.AddNew
        Else
            rs2.Edit
        End If
        
        rs2!AccountPositionID = lID
        rs2!FillID = rs!FillID
        rs2!FillDate = rs!FillDate
        
        rs2.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion25"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion26
'' Description: Apply version 26 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion26()
On Error GoTo ErrSection:

    Dim t As TableDef
    Dim rs As Recordset
    Dim rs2 As Recordset
    
    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "Session") Then
            .Fields.Append .CreateField("Session", dbText, 1)
            .Fields("Session").DefaultValue = ""
            .Fields("Session").Required = False
            .Fields("Session").AllowZeroLength = True
        End If
    End With
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & rs!AccountID & ";", dbOpenDynaset)
        If Not (rs2.BOF And rs2.EOF) Then
            If rs2!AccountType = 8 Then ' eTT_AccountType_Alaron Then
                rs!Session = "P"
            Else
                rs!Session = " "
            End If
        Else
            rs!Session = " "
        End If
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion26"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion28
'' Description: Apply version 28 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion28()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rs As Recordset                 ' Recordset into the database
    Dim i As Index                      ' Index for a table in the database
    Dim r As Relation                   ' Relation object in the database
    Dim q As QueryDef                   ' Query def into the database
    
    ' Update the fills table with new fields and rearrange the field order...
    Set t = m.DB.TableDefs("tblFills")
    With t
        If Not ItemExists(.Fields, "AccountID") Then
            .Fields.Append .CreateField("AccountID", dbLong)
            .Fields("AccountID").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.AccountID=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SymbolID") Then
            .Fields.Append .CreateField("SymbolID", dbLong)
            .Fields("SymbolID").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.SymbolID=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Symbol") Then
            .Fields.Append .CreateField("Symbol", dbText, 50)
            .Fields("Symbol").DefaultValue = ""
            .Fields("Symbol").Required = False
            .Fields("Symbol").AllowZeroLength = True
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.Symbol='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "AutoTradeID") Then
            .Fields.Append .CreateField("AutoTradeID", dbLong)
            .Fields("AutoTradeID").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.AutoTradeID=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "BrokerOrderID") Then
            .Fields.Append .CreateField("BrokerOrderID", dbText, 50)
            .Fields("BrokerOrderID").DefaultValue = ""
            .Fields("BrokerOrderID").Required = False
            .Fields("BrokerOrderID").AllowZeroLength = True
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.BrokerOrderID='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "IsManual") Then
            .Fields.Append .CreateField("IsManual", dbBoolean)
            .Fields("IsManual").DefaultValue = False
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.IsManual=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "IsSnapshot") Then
            .Fields.Append .CreateField("IsSnapshot", dbBoolean)
            .Fields("IsSnapshot").DefaultValue = False
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.IsSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Notes") Then
            .Fields.Append .CreateField("Notes", dbMemo)
            .Fields("Notes").AllowZeroLength = True
            .Fields("Notes").Required = False
            .Fields("Notes").DefaultValue = ""
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.Notes='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Sequence") Then
            .Fields.Append .CreateField("Sequence", dbLong)
            .Fields("Sequence").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.Sequence=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Action") Then
            .Fields.Append .CreateField("Action", dbText, 1)
            .Fields("Action").DefaultValue = ""
            .Fields("Action").Required = False
            .Fields("Action").AllowZeroLength = True
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.Action='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "TotalProfit") Then
            .Fields.Append .CreateField("TotalProfit", dbDouble)
            .Fields("TotalProfit").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.TotalProfit=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SequenceCategory") Then
            .Fields.Append .CreateField("SequenceCategory", dbLong)
            .Fields("SequenceCategory").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.SequenceCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "ActionCategory") Then
            .Fields.Append .CreateField("ActionCategory", dbText, 1)
            .Fields("ActionCategory").DefaultValue = ""
            .Fields("ActionCategory").Required = False
            .Fields("ActionCategory").AllowZeroLength = True
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.ActionCategory='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "ClosedProfitCategory") Then
            .Fields.Append .CreateField("ClosedProfitCategory", dbDouble)
            .Fields("ClosedProfitCategory").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.ClosedProfitCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SessionProfitCategory") Then
            .Fields.Append .CreateField("SessionProfitCategory", dbDouble)
            .Fields("SessionProfitCategory").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.SessionProfitCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "TotalProfitCategory") Then
            .Fields.Append .CreateField("TotalProfitCategory", dbDouble)
            .Fields("TotalProfitCategory").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.TotalProfitCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SessionQuantityCategory") Then
            .Fields.Append .CreateField("SessionQuantityCategory", dbLong)
            .Fields("SessionQuantityCategory").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.SessionQuantityCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "CurrentPositionCategory") Then
            .Fields.Append .CreateField("CurrentPositionCategory", dbLong)
            .Fields("CurrentPositionCategory").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.CurrentPositionCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "AverageEntryCategory") Then
            .Fields.Append .CreateField("AverageEntryCategory", dbDouble)
            .Fields("AverageEntryCategory").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.AverageEntryCategory=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumOpenCategory") Then
            .Fields.Append .CreateField("NumOpenCategory", dbLong)
            .Fields("NumOpenCategory").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblFills SET tblFills.NumOpenCategory=0;"
            q.Execute
        End If
    
        .Fields("AccountID").OrdinalPosition = 3
        .Fields("SymbolID").OrdinalPosition = 4
        .Fields("Symbol").OrdinalPosition = 5
        .Fields("AutoTradeID").OrdinalPosition = 6
        .Fields("BrokerFillID").OrdinalPosition = 7
        .Fields("BrokerOrderID").OrdinalPosition = 8
        .Fields("BuySell").OrdinalPosition = 9
        .Fields("FillQuantity").OrdinalPosition = 10
        .Fields("FillPrice").OrdinalPosition = 11
        .Fields("FillDate").OrdinalPosition = 12
        .Fields("SessionDate").OrdinalPosition = 13
        .Fields("Fees").OrdinalPosition = 14
        .Fields("IsManual").OrdinalPosition = 15
        .Fields("IsSnapshot").OrdinalPosition = 16
        .Fields("Notes").OrdinalPosition = 17
        .Fields("Sequence").OrdinalPosition = 18
        .Fields("Action").OrdinalPosition = 19
        .Fields("ClosedProfit").OrdinalPosition = 20
        .Fields("SessionProfit").OrdinalPosition = 21
        .Fields("TotalProfit").OrdinalPosition = 22
        .Fields("SessionQuantity").OrdinalPosition = 23
        .Fields("CurrentPosition").OrdinalPosition = 24
        .Fields("AverageEntry").OrdinalPosition = 25
        .Fields("NumOpen").OrdinalPosition = 26
        .Fields("SequenceCategory").OrdinalPosition = 27
        .Fields("ActionCategory").OrdinalPosition = 28
        .Fields("ClosedProfitCategory").OrdinalPosition = 29
        .Fields("SessionProfitCategory").OrdinalPosition = 30
        .Fields("TotalProfitCategory").OrdinalPosition = 31
        .Fields("SessionQuantityCategory").OrdinalPosition = 32
        .Fields("CurrentPositionCategory").OrdinalPosition = 33
        .Fields("AverageEntryCategory").OrdinalPosition = 34
        .Fields("NumOpenCategory").OrdinalPosition = 35
    End With
    
    ' Update the account position table with new fields and rearrange the field order...
    Set t = m.DB.TableDefs("tblAccountPositions")
    With t
        If Not ItemExists(.Fields, "NumBuys") Then
            .Fields.Append .CreateField("NumBuys", dbLong)
            .Fields("NumBuys").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumBuys=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumSells") Then
            .Fields.Append .CreateField("NumSells", dbLong)
            .Fields("NumSells").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumSells=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumTotal") Then
            .Fields.Append .CreateField("NumTotal", dbLong)
            .Fields("NumTotal").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumTotal=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "RollingSum") Then
            .Fields.Append .CreateField("RollingSum", dbDouble)
            .Fields("RollingSum").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.RollingSum=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "ClosedProfit") Then
            .Fields.Append .CreateField("ClosedProfit", dbDouble)
            .Fields("ClosedProfit").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.ClosedProfit=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "StartFillTime") Then
            .Fields.Append .CreateField("StartFillTime", dbDouble)
            .Fields("StartFillTime").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.StartFillTime=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "TotalFees") Then
            .Fields.Append .CreateField("TotalFees", dbDouble)
            .Fields("TotalFees").DefaultValue = 0#
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.TotalFees=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "LastFillSeq") Then
            .Fields.Append .CreateField("LastFillSeq", dbLong)
            .Fields("LastFillSeq").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.LastFillSeq=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SessionDateSnapshot") Then
            .Fields.Append .CreateField("SessionDateSnapshot", dbLong)
            .Fields("SessionDateSnapshot").DefaultValue = 0&
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.SessionDateSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SessionQuantitySnapshot") Then
            .Fields.Append .CreateField("SessionQuantitySnapshot", dbLong)
            .Fields("SessionQuantitySnapshot").DefaultValue = 0&
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.SessionQuantitySnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "SessionProfitSnapshot") Then
            .Fields.Append .CreateField("SessionProfitSnapshot", dbDouble)
            .Fields("SessionProfitSnapshot").DefaultValue = 0#
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.SessionProfitSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumBuysSnapshot") Then
            .Fields.Append .CreateField("NumBuysSnapshot", dbLong)
            .Fields("NumBuysSnapshot").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumBuysSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumSellsSnapshot") Then
            .Fields.Append .CreateField("NumSellsSnapshot", dbLong)
            .Fields("NumSellsSnapshot").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumSellsSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "CurrentPositionSnapshot") Then
            .Fields.Append .CreateField("CurrentPositionSnapshot", dbLong)
            .Fields("CurrentPositionSnapshot").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.CurrentPositionSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "NumTotalSnapshot") Then
            .Fields.Append .CreateField("NumTotalSnapshot", dbLong)
            .Fields("NumTotalSnapshot").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.NumTotalSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "RollingSumSnapshot") Then
            .Fields.Append .CreateField("RollingSumSnapshot", dbDouble)
            .Fields("RollingSumSnapshot").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.RollingSumSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "ClosedProfitSnapshot") Then
            .Fields.Append .CreateField("ClosedProfitSnapshot", dbDouble)
            .Fields("ClosedProfitSnapshot").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.ClosedProfitSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "AverageEntrySnapshot") Then
            .Fields.Append .CreateField("AverageEntrySnapshot", dbDouble)
            .Fields("AverageEntrySnapshot").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.AverageEntrySnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "LastTradedSnapshot") Then
            .Fields.Append .CreateField("LastTradedSnapshot", dbDouble)
            .Fields("LastTradedSnapshot").DefaultValue = 0#
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.LastTradedSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "TotalFeesSnapshot") Then
            .Fields.Append .CreateField("TotalFeesSnapshot", dbDouble)
            .Fields("TotalFeesSnapshot").DefaultValue = 0#
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.TotalFeesSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "LastFillSeqSnapshot") Then
            .Fields.Append .CreateField("LastFillSeqSnapshot", dbLong)
            .Fields("LastFillSeqSnapshot").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.LastFillSeqSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "FillMatchMode") Then
            .Fields.Append .CreateField("FillMatchMode", dbByte)
            .Fields.Item("FillMatchMode").DefaultValue = 0
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.FillMatchMode=0;"
            q.Execute
        End If
                
        .Fields("FillMatchMode").OrdinalPosition = 5
        .Fields("StartFillTime").OrdinalPosition = 6
        .Fields("SessionDate").OrdinalPosition = 7
        .Fields("SessionQuantity").OrdinalPosition = 8
        .Fields("SessionProfit").OrdinalPosition = 9
        .Fields("NumBuys").OrdinalPosition = 10
        .Fields("NumSells").OrdinalPosition = 11
        .Fields("CurrentPosition").OrdinalPosition = 12
        .Fields("NumTotal").OrdinalPosition = 13
        .Fields("RollingSum").OrdinalPosition = 14
        .Fields("ClosedProfit").OrdinalPosition = 15
        .Fields("AverageEntry").OrdinalPosition = 16
        .Fields("LastTraded").OrdinalPosition = 17
        .Fields("TotalFees").OrdinalPosition = 18
        .Fields("LastFillSeq").OrdinalPosition = 19
        .Fields("SessionDateSnapshot").OrdinalPosition = 20
        .Fields("SessionQuantitySnapshot").OrdinalPosition = 21
        .Fields("SessionProfitSnapshot").OrdinalPosition = 22
        .Fields("NumBuysSnapshot").OrdinalPosition = 23
        .Fields("NumSellsSnapshot").OrdinalPosition = 24
        .Fields("CurrentPositionSnapshot").OrdinalPosition = 25
        .Fields("NumTotalSnapshot").OrdinalPosition = 26
        .Fields("RollingSumSnapshot").OrdinalPosition = 27
        .Fields("ClosedProfitSnapshot").OrdinalPosition = 28
        .Fields("AverageEntrySnapshot").OrdinalPosition = 29
        .Fields("LastTradedSnapshot").OrdinalPosition = 30
        .Fields("TotalFeesSnapshot").OrdinalPosition = 31
        .Fields("LastFillSeqSnapshot").OrdinalPosition = 32
    End With
    
    ' Create the account position entries table if it doesn't already exist...
    If Not ItemExists(m.DB.TableDefs, "tblAccountPositionEntries") Then
        Set t = m.DB.CreateTableDef("tblAccountPositionEntries")
        With t
            If Not ItemExists(.Fields, "AccountPositionID") Then
                .Fields.Append .CreateField("AccountPositionID", dbLong)
                .Fields("AccountPositionID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillID") Then
                .Fields.Append .CreateField("FillID", dbLong)
                .Fields("FillID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillSequence") Then
                .Fields.Append .CreateField("FillSequence", dbLong)
                .Fields("FillSequence").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillQuantity") Then
                .Fields.Append .CreateField("FillQuantity", dbLong)
                .Fields("FillQuantity").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "FillPrice") Then
                .Fields.Append .CreateField("FillPrice", dbDouble)
                .Fields("FillPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "IsSnapshot") Then
                .Fields.Append .CreateField("IsSnapshot", dbBoolean)
                .Fields("IsSnapshot").DefaultValue = False
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AccountPositionID")
                    .Fields.Append .CreateField("FillID")
                    .Fields.Append .CreateField("IsSnapshot")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If
    
    ' Create a relationship between the account position table and the account position entry table...
    If Not ItemExists(m.DB.Relations, "tblAccountPositionstblAccountPositionEntries") Then
        Set r = m.DB.CreateRelation("tblAccountPositionstblAccountPositionEntries", "tblAccountPositions", _
            "tblAccountPositionEntries", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountPositionID")
            .Fields.Item("AccountPositionID").ForeignName = "AccountPositionID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    ' Update the orders table with new fields and rearrange the field order...
    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "SessionDate") Then
            .Fields.Append .CreateField("SessionDate", dbLong)
            .Fields("SessionDate").DefaultValue = 0&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrders SET tblOrders.SessionDate=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "IsSnapshot") Then
            .Fields.Append .CreateField("IsSnapshot", dbBoolean)
            .Fields("IsSnapshot").DefaultValue = False
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrders SET tblOrders.IsSnapshot=0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "PreviousBrokerID") Then
            .Fields.Append .CreateField("PreviousBrokerID", dbText, 100)
            .Fields("PreviousBrokerID").AllowZeroLength = True
            .Fields("PreviousBrokerID").Required = False
            .Fields("PreviousBrokerID").DefaultValue = ""
        End If
        
        .Fields("AccountID").OrdinalPosition = 2
        .Fields("SymbolID").OrdinalPosition = 3
        .Fields("Symbol").OrdinalPosition = 4
        .Fields("AutoTradeItemID").OrdinalPosition = 5
        .Fields("GenesisOrderID").OrdinalPosition = 6
        .Fields("BrokerOrderID").OrdinalPosition = 7
        .Fields("ExchangeOrderID").OrdinalPosition = 8
        .Fields("PreviousBrokerID").OrdinalPosition = 9
        .Fields("BuySell").OrdinalPosition = 10
        .Fields("Quantity").OrdinalPosition = 11
        .Fields("LimitPrice").OrdinalPosition = 12
        .Fields("StopPrice").OrdinalPosition = 13
        .Fields("OrderType").OrdinalPosition = 14
        .Fields("Expiration").OrdinalPosition = 15
        .Fields("Session").OrdinalPosition = 16
        .Fields("OrderDate").OrdinalPosition = 17
        .Fields("SessionDate").OrdinalPosition = 18
        .Fields("EnterExit").OrdinalPosition = 19
        .Fields("Fees").OrdinalPosition = 20
        .Fields("IsSnapshot").OrdinalPosition = 21
        .Fields("Status").OrdinalPosition = 22
        .Fields("StatusDate").OrdinalPosition = 23
        .Fields("Message").OrdinalPosition = 24
        .Fields("DateSent").OrdinalPosition = 25
        .Fields("DateHostRecd").OrdinalPosition = 26
        .Fields("DateExchRecd").OrdinalPosition = 27
        .Fields("DateExchAckn").OrdinalPosition = 28
        .Fields("CancelOrderID").OrdinalPosition = 29
        .Fields("TriggerOrderID").OrdinalPosition = 30
        .Fields("TriggerOptions").OrdinalPosition = 31
        .Fields("ConditionOptions").OrdinalPosition = 32
        .Fields("TrailAmount").OrdinalPosition = 33
        .Fields("TrailOptions").OrdinalPosition = 34
        .Fields("ExitPos").OrdinalPosition = 35
    End With

    ' Create the account position trades table if it doesn't already exist...
    If Not ItemExists(m.DB.TableDefs, "tblAccountPositionTrades") Then
        Set t = m.DB.CreateTableDef("tblAccountPositionTrades")
        With t
            If Not ItemExists(.Fields, "AccountPositionID") Then
                .Fields.Append .CreateField("AccountPositionID", dbLong)
                .Fields("AccountPositionID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "TradeNumber") Then
                .Fields.Append .CreateField("TradeNumber", dbLong)
                .Fields("TradeNumber").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Direction") Then
                .Fields.Append .CreateField("Direction", dbText, 1)
                .Fields("Direction").DefaultValue = ""
                .Fields("Direction").Required = False
                .Fields("Direction").AllowZeroLength = True
            End If
            If Not ItemExists(.Fields, "Quantity") Then
                .Fields.Append .CreateField("Quantity", dbLong)
                .Fields("Quantity").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "EntryBrokerFillID") Then
                .Fields.Append .CreateField("EntryBrokerFillID", dbText, 70)
                .Fields("EntryBrokerFillID").DefaultValue = ""
                .Fields("EntryBrokerFillID").Required = False
                .Fields("EntryBrokerFillID").AllowZeroLength = True
            End If
            If Not ItemExists(.Fields, "EntryPrice") Then
                .Fields.Append .CreateField("EntryPrice", dbDouble)
                .Fields("EntryPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "EntryTime") Then
                .Fields.Append .CreateField("EntryTime", dbDouble)
                .Fields("EntryTime").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "ExitBrokerFillID") Then
                .Fields.Append .CreateField("ExitBrokerFillID", dbText, 70)
                .Fields("ExitBrokerFillID").DefaultValue = ""
                .Fields("ExitBrokerFillID").Required = False
                .Fields("ExitBrokerFillID").AllowZeroLength = True
            End If
            If Not ItemExists(.Fields, "ExitPrice") Then
                .Fields.Append .CreateField("ExitPrice", dbDouble)
                .Fields("ExitPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "ExitTime") Then
                .Fields.Append .CreateField("ExitTime", dbDouble)
                .Fields("ExitTime").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "ClosedProfit") Then
                .Fields.Append .CreateField("ClosedProfit", dbDouble)
                .Fields("ClosedProfit").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "IsOpen") Then
                .Fields.Append .CreateField("IsOpen", dbBoolean)
                .Fields("IsOpen").DefaultValue = False
            End If
            If Not ItemExists(.Fields, "IsSnapshot") Then
                .Fields.Append .CreateField("IsSnapshot", dbBoolean)
                .Fields("IsSnapshot").DefaultValue = False
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AccountPositionID")
                    .Fields.Append .CreateField("TradeNumber")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If
    
    ' Create a relationship between the account position table and the account position trades table...
    If Not ItemExists(m.DB.Relations, "tblAccountPositionstblAccountPositionTrades") Then
        Set r = m.DB.CreateRelation("tblAccountPositionstblAccountPositionTrades", "tblAccountPositions", _
            "tblAccountPositionTrades", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountPositionID")
            .Fields.Item("AccountPositionID").ForeignName = "AccountPositionID"
        End With
        
        m.DB.Relations.Append r
    End If

    ' Remove the account position fills table and all of its relationships...
    If ItemExists(m.DB.Relations, "tblAccountPositionstblAccountPositionFills") Then
        m.DB.Relations.Delete "tblAccountPositionstblAccountPositionFills"
    End If
    If ItemExists(m.DB.Relations, "tblFillstblAccountPositionFills") Then
        m.DB.Relations.Delete "tblFillstblAccountPositionFills"
    End If
    If ItemExists(m.DB.TableDefs, "tblAccountPositionFills") Then
        m.DB.TableDefs.Delete "tblAccountPositionFills"
    End If
    
    ' Create the account adjustments table if it doesn't already exist...
    If Not ItemExists(m.DB.TableDefs, "tblAccountAdjustments") Then
        Set t = m.DB.CreateTableDef("tblAccountAdjustments")
        With t
            If Not ItemExists(.Fields, "AdjustmentID") Then
                .Fields.Append .CreateField("AdjustmentID", dbLong)
                .Fields("AdjustmentID").DefaultValue = 0&
                .Fields("AdjustmentID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "AccountID") Then
                .Fields.Append .CreateField("AccountID", dbLong)
                .Fields("AccountID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Description") Then
                .Fields.Append .CreateField("Description", dbText, 50)
                .Fields("Description").Required = False
                .Fields("Description").AllowZeroLength = True
                .Fields("Description").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "AdjustmentTime") Then
                .Fields.Append .CreateField("AdjustmentTime", dbDouble)
                .Fields("AdjustmentTime").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "AdjustmentAmount") Then
                .Fields.Append .CreateField("AdjustmentAmount", dbDouble)
                .Fields("AdjustmentAmount").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "Notes") Then
                .Fields.Append .CreateField("Notes", dbMemo)
                .Fields("Notes").Required = False
                .Fields("Notes").AllowZeroLength = True
                .Fields("Notes").DefaultValue = ""
            End If
            
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("AdjustmentID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If

    ' Move any position notes to the entry fills for that position...
    MovePositionNotesToFills
    
    ' Fill the new fields and tables for version 28 before we remove the old tables and
    ' relationships...
    FillVersion28
    
    ' Create a relationship between the accounts table and the account adjustments table...
    If Not ItemExists(m.DB.Relations, "tblAccountstblAccountAdjustments") Then
        Set r = m.DB.CreateRelation("tblAccountstblAccountAdjustments", "tblAccounts", _
            "tblAccountAdjustments", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountID")
            .Fields.Item("AccountID").ForeignName = "AccountID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    ' Remove the symbol info table if it exists...
    If ItemExists(m.DB.TableDefs, "tblSymbolInfo") Then
        m.DB.TableDefs.Delete "tblSymbolInfo"
    End If
    
    ' Remove the position fills table and all of its relationships...
    If ItemExists(m.DB.Relations, "tblFillstblPositionFills") Then
        m.DB.Relations.Delete "tblFillstblPositionFills"
    End If
    If ItemExists(m.DB.Relations, "tblPositionstblPositionFills") Then
        m.DB.Relations.Delete "tblPositionstblPositionFills"
    End If
    If ItemExists(m.DB.TableDefs, "tblPositionFills") Then
        m.DB.TableDefs.Delete "tblPositionFills"
    End If
    
    ' Remove the position table and all of its relationships...
    If ItemExists(m.DB.Relations, "tblActivitiestblPositions") Then
        m.DB.Relations.Delete "tblActivitiestblPositions"
    End If
    If ItemExists(m.DB.TableDefs, "tblPositions") Then
        m.DB.TableDefs.Delete "tblPositions"
    End If
    
    ' Remove the activities table and all of its relationships...
    If ItemExists(m.DB.Relations, "tblAccountstblActivities") Then
        m.DB.Relations.Delete "tblAccountstblActivities"
    End If
    If ItemExists(m.DB.TableDefs, "tblActivities") Then
        m.DB.TableDefs.Delete "tblActivities"
    End If

    ' Create a relationship between the accounts table and the fills table...
    If Not ItemExists(m.DB.Relations, "tblAccountstblFills") Then
        Set r = m.DB.CreateRelation("tblAccountstblFills", "tblAccounts", _
            "tblFills", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountID")
            .Fields.Item("AccountID").ForeignName = "AccountID"
        End With
        
        m.DB.Relations.Append r
    End If

    ' Remove the relationship between the fills table and the orders table...
    If ItemExists(m.DB.Relations, "tblOrderstblFills") Then
        m.DB.Relations.Delete "tblOrderstblFills"
    End If
    
    ' Create a relationship between the accounts table and the orders table...
    If Not ItemExists(m.DB.Relations, "tblAccountstblOrders") Then
        Set r = m.DB.CreateRelation("tblAccountstblOrders", "tblAccounts", _
            "tblOrders", dbRelationDeleteCascade Or dbRelationUpdateCascade)
        With r
            .Fields.Append .CreateField("AccountID")
            .Fields.Item("AccountID").ForeignName = "AccountID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    ' Create the Fifo/Lifo option field in the accounts table...
    Set t = m.DB.TableDefs("tblAccounts")
    With t
        If Not ItemExists(.Fields, "FillMatchMode") Then
            .Fields.Append .CreateField("FillMatchMode", dbByte)
            .Fields.Item("FillMatchMode").DefaultValue = 0
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccounts SET tblAccounts.FillMatchMode=0;"
            q.Execute
        End If
        
        .Fields("AccountID").OrdinalPosition = 0
        .Fields("AccountNumber").OrdinalPosition = 1
        .Fields("Name").OrdinalPosition = 2
        .Fields("AccountType").OrdinalPosition = 3
        .Fields("StartingDate").OrdinalPosition = 4
        .Fields("Broker").OrdinalPosition = 5
        .Fields("Comms").OrdinalPosition = 6
        .Fields("SecTypeMask").OrdinalPosition = 7
        .Fields("FillRT").OrdinalPosition = 8
        .Fields("FillMatchMode").OrdinalPosition = 9
        .Fields("UserName").OrdinalPosition = 10
        .Fields("Password").OrdinalPosition = 11
        .Fields("StartingBalance").OrdinalPosition = 12
        .Fields("TotalFees").OrdinalPosition = 13
        .Fields("TotalAdjustments").OrdinalPosition = 14
        .Fields("ClosedProfit").OrdinalPosition = 15
        .Fields("CurrentBalance").OrdinalPosition = 16
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion28"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillVersion28
'' Description: Fill version 28 data in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillVersion28()
On Error GoTo ErrSection:

    Dim rsAcctPos As Recordset          ' Recordset into the database
    Dim rsAcctPos2 As Recordset         ' Recordset into the database
    Dim rsFills As Recordset            ' Recordset into the database
    Dim rsOrders As Recordset           ' Recordset into the database
    Dim rs As Recordset                 ' Recordset into the database
    Dim rsAdj As Recordset              ' Recordset into the database
    Dim lPrevSession As Long            ' Previous session to the session date
    Dim lIndex As Long                  ' Index into a for loop
    
    ' 1) Remove any orders that have an account ID that does not exist in the
    '    database.  This will also remove the fills that were for the order...
    Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rsOrders.EOF
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] WHERE [AccountID]=" & Str(rsOrders!AccountID) & ";", dbOpenDynaset)
        If (rs.BOF And rs.EOF) Then
            rsOrders.Delete
        End If
        
        rsOrders.MoveNext
    Loop
    
    ' 2) Pass 1: Fill in new fields in the Fills table from existing fields in the Orders table,
    '    make sure that all applicable account position records exist and that the session
    '    date is set correctly on the account position records...
    Set rsFills = m.DB.OpenRecordset("SELECT * " & _
                "FROM tblFills INNER JOIN tblOrders ON tblFills.OrderID=tblOrders.OrderID " & _
                "ORDER BY [FillDate];", dbOpenDynaset)
    Do While Not rsFills.EOF
        rsFills.Edit
        rsFills![tblFills.AccountID] = rsFills![tblOrders.AccountID]
        rsFills![tblFills.SymbolID] = rsFills![tblOrders.SymbolID]
        rsFills![tblFills.Symbol] = rsFills![tblOrders.Symbol]
        rsFills!AutoTradeID = rsFills!AutoTradeItemID
        rsFills![tblFills.BrokerOrderID] = rsFills![tblOrders.BrokerOrderID]
        rsFills!IsManual = (Len(rsFills!BrokerFillID) = 0)
        rsFills![tblFills.IsSnapshot] = True
        rsFills.Update
        
        ' Make sure that the account position record exists for the fill key (category)...
        Set rsAcctPos = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountID]=" & Str(rsFills![tblFills.AccountID]) & " AND [Symbol]='" & rsFills![tblFills.Symbol] & "' AND [SymbolID]=" & Str(rsFills![tblFills.SymbolID]) & " AND [AutoTradeItemID]=" & Str(rsFills!AutoTradeID) & ";", dbOpenDynaset)
        If rsAcctPos.BOF And rsAcctPos.EOF Then
            rsAcctPos.AddNew
            rsAcctPos!AccountID = rsFills![tblFills.AccountID]
            rsAcctPos!Symbol = rsFills![tblFills.Symbol]
            rsAcctPos!SymbolID = rsFills![tblFills.SymbolID]
            rsAcctPos!AutoTradeItemID = rsFills!AutoTradeID
            rsAcctPos!SessionDate = rsFills![tblFills.SessionDate]
            rsAcctPos!LastTraded = -1#
            rsAcctPos!LastTradedSnapshot = -1#
            rsAcctPos.Update
        Else
            rsAcctPos.Edit
            If rsFills![tblFills.SessionDate] > rsAcctPos!SessionDate Then
                rsAcctPos!SessionDate = rsFills![tblFills.SessionDate]
            End If
            rsAcctPos!LastTraded = -1#
            rsAcctPos!LastTradedSnapshot = -1#
            rsAcctPos.Update
        End If
        
        ' Make sure that the account position record exists for the fill key (total)...
        Set rsAcctPos = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountID]=" & Str(rsFills![tblFills.AccountID]) & " AND [Symbol]='" & rsFills![tblFills.Symbol] & "' AND [SymbolID]=" & Str(rsFills![tblFills.SymbolID]) & " AND [AutoTradeItemID]=-1;", dbOpenDynaset)
        If rsAcctPos.BOF And rsAcctPos.EOF Then
            rsAcctPos.AddNew
            rsAcctPos!AccountID = rsFills![tblFills.AccountID]
            rsAcctPos!Symbol = rsFills![tblFills.Symbol]
            rsAcctPos!SymbolID = rsFills![tblFills.SymbolID]
            rsAcctPos!AutoTradeItemID = -1&
            rsAcctPos!SessionDate = rsFills![tblFills.SessionDate]
            rsAcctPos!LastTraded = -1#
            rsAcctPos!LastTradedSnapshot = -1#
            rsAcctPos.Update
        Else
            rsAcctPos.Edit
            If rsFills![tblFills.SessionDate] > rsAcctPos!SessionDate Then
                rsAcctPos!SessionDate = rsFills![tblFills.SessionDate]
            End If
            rsAcctPos!LastTraded = -1#
            rsAcctPos!LastTradedSnapshot = -1#
            rsAcctPos.Update
        End If
        
        rsFills.MoveNext
    Loop
    
    ' 3) Make sure that a "total" line exists for the account position (auto trade item ID: -1)...
    Set rsAcctPos = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [AutoTradeItemID]<>-1;", dbOpenDynaset)
    Do While Not rsAcctPos.EOF
        Set rsAcctPos2 = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountID]=" & Str(rsAcctPos!AccountID) & " AND [SymbolID]=" & Str(rsAcctPos!SymbolID) & " AND [Symbol]='" & rsAcctPos!Symbol & "' AND [AutoTradeItemID]=-1;", dbOpenDynaset)
        If rsAcctPos2.BOF And rsAcctPos2.EOF Then
            rsAcctPos2.AddNew
            For lIndex = 1 To rsAcctPos.Fields.Count - 1
                rsAcctPos2.Fields(lIndex).Value = rsAcctPos.Fields(lIndex).Value
            Next lIndex
            rsAcctPos2!AutoTradeItemID = -1&
            rsAcctPos2.Update
        End If
        
        rsAcctPos.MoveNext
    Loop
    
    ' 4) Pass 2: Walk back through the fills and change the snapshot flag such that fills before
    '    the session date of the account position are non-snapshot and fills that are on
    '    the session date of the account position are snapshot...
    Set rsAcctPos = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [AutoTradeItemID]=-1;", dbOpenDynaset)
    Do While Not rsAcctPos.EOF
        If rsAcctPos!SessionDate > 0& Then
            lPrevSession = rsAcctPos!SessionDate - 1&
            Do While IsWeekday(lPrevSession) = False
                lPrevSession = lPrevSession - 1&
            Loop
            
            Set rsFills = m.DB.OpenRecordset("SELECT * FROM [tblFills] " & _
                        "WHERE [AccountID]=" & Str(rsAcctPos!AccountID) & " AND [SymbolID]=" & Str(rsAcctPos!SymbolID) & " AND [Symbol]='" & rsAcctPos!Symbol & "';", dbOpenDynaset)
            Do While Not rsFills.EOF
                If rsFills!SessionDate < lPrevSession Then
                    rsFills.Edit
                    rsFills!IsSnapshot = False
                    rsFills.Update
                End If
                
                rsFills.MoveNext
            Loop
        End If
        
        rsAcctPos.MoveNext
    Loop
    
    ' 5) Set some of the new fields in the orders table and mark them so that the post
    '    process will pick them up and fix them some more...
    Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rsOrders.EOF
        rsOrders.Edit
        rsOrders!SessionDate = -1&
        rsOrders!IsSnapshot = True
        rsOrders!PreviousBrokerID = ""
        rsOrders.Update
        
        rsOrders.MoveNext
    Loop
        
    ' 6) Move the adjustments from the activities table to the new adjustments table...
    If ItemExists(m.DB.TableDefs, "tblActivities") Then
        Set rsAdj = m.DB.OpenRecordset("SELECT * FROM [tblAccountAdjustments];", dbOpenDynaset)
        If (rsAdj.BOF) And (rsAdj.EOF) Then
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblActivities] WHERE [ActivityType]=1;", dbOpenDynaset)
            Do While Not rs.EOF
                rsAdj.AddNew
                rsAdj!AccountID = rs!AccountID
                rsAdj!Description = rs!ActivityName
                rsAdj!AdjustmentTime = rs!StartingDate
                rsAdj!AdjustmentAmount = rs!AdjustAmount
                rsAdj!Notes = rs!Journal
                rsAdj.Update
                
                rs.MoveNext
            Loop
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.FillVersion28"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MovePositionNotesToFills
'' Description: Move the position notes to the entry fills for the position
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MovePositionNotesToFills()
On Error GoTo ErrSection:

    Dim rsPositions As Recordset        ' Recordset of positions in the database
    Dim rsPosFills As Recordset         ' Recordset of position fills in the database
    Dim rsFills As Recordset            ' Recordset of fills in the database
    
    If ItemExists(m.DB.TableDefs, "tblPositions") Then
        Set rsPositions = m.DB.OpenRecordset("SELECT * FROM [tblPositions];", dbOpenDynaset)
        Do While Not rsPositions.EOF
            If Len(rsPositions!Notes) > 0 Then
                Set rsPosFills = m.DB.OpenRecordset("SELECT * FROM [tblPositionFills] " & _
                            "WHERE [PositionID]=" & Str(rsPositions!PositionID) & ";", dbOpenDynaset)
                Do While Not rsPosFills.EOF
                    Set rsFills = m.DB.OpenRecordset("SELECT * FROM [tblFills] " & _
                                "WHERE [FillID]=" & Str(rsPosFills!FillID) & ";", dbOpenDynaset)
                    If Not (rsFills.BOF And rsFills.EOF) Then
                        If ((rsPositions!Position = 1) And (rsFills!BuySell = True)) Or ((rsPositions!Position = 2) And (rsFills!BuySell = False)) Then
                            rsFills.Edit
                            If Len(rsFills!Notes) = 0 Then
                                rsFills!Notes = rsPositions!Notes
                            Else
                                rsFills!Notes = rsFills!Notes & vbCrLf & rsPositions!Notes
                            End If
                            rsFills.Update
                        End If
                    End If
                    
                    rsPosFills.MoveNext
                Loop
            End If
            
            rsPositions.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.MovePositionNotesToFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillVersion29
'' Description: Fill version 29 data in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillVersion29()
On Error GoTo ErrSection:

    Dim rsAcctPos As Recordset          ' Recordset into the database
    Dim rsAcctPos2 As Recordset         ' Recordset into the database
    
    ' Make sure that a "manual" line exists for the account position (auto trade item ID: 0)...
    Set rsAcctPos = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [AutoTradeItemID]=-1;", dbOpenDynaset)
    Do While Not rsAcctPos.EOF
        Set rsAcctPos2 = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] " & _
                    "WHERE [AccountID]=" & Str(rsAcctPos!AccountID) & " AND [SymbolID]=" & Str(rsAcctPos!SymbolID) & " AND [Symbol]='" & rsAcctPos!Symbol & "' AND [AutoTradeItemID]=0;", dbOpenDynaset)
        If rsAcctPos2.BOF And rsAcctPos2.EOF Then
            rsAcctPos2.AddNew
            rsAcctPos2!AccountID = rsAcctPos!AccountID
            rsAcctPos2!Symbol = rsAcctPos!Symbol
            rsAcctPos2!SymbolID = rsAcctPos!SymbolID
            rsAcctPos2!AutoTradeItemID = 0&
            rsAcctPos2!FillMatchMode = rsAcctPos!FillMatchMode
            rsAcctPos2.Update
        End If
        
        rsAcctPos.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.FillVersion29"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion30
'' Description: Apply version 30 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion30()
On Error GoTo ErrSection:

    Dim rsOrders As Recordset           ' Recordset into the database
    Dim astrOldOptions As cGdArray      ' Array of conditional options
    Dim astrNewOptions As cGdArray      ' Array of conditional options
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrOldOptions = New cGdArray
    astrOldOptions.Create eGDARRAY_Strings
    Set astrNewOptions = New cGdArray
    astrNewOptions.Create eGDARRAY_Strings
    
    Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rsOrders.EOF
        If Len(rsOrders!ConditionOptions) > 0 Then
            If (InStr(rsOrders!ConditionOptions, vbTab) = 0) And (InStr(rsOrders!ConditionOptions, "|") <> 0) Then
                rsOrders.Edit
                rsOrders!ConditionOptions = Replace(rsOrders!ConditionOptions, "|", vbTab)
                rsOrders.Update
            ElseIf (InStr(rsOrders!ConditionOptions, vbTab) = 0) And (InStr(rsOrders!ConditionOptions, ",") <> 0) Then
                astrOldOptions.SplitFields rsOrders!ConditionOptions, ","
                For lIndex = 0 To 7
                    astrNewOptions(lIndex) = astrOldOptions(lIndex)
                Next lIndex
                For lIndex = 8 To astrOldOptions.Size - 2
                    astrNewOptions(7) = astrNewOptions(7) & "," & astrOldOptions(lIndex)
                Next lIndex
                astrNewOptions(8) = astrOldOptions(astrOldOptions.Size - 1)
                
                rsOrders.Edit
                rsOrders!ConditionOptions = astrNewOptions.JoinFields(vbTab)
                rsOrders.Update
            End If
        End If
        
        rsOrders.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion30"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion31
'' Description: Apply version 31 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion31()
On Error GoTo ErrSection:

    Dim rsFills As Recordset            ' Recordset into the database
    
    Set rsFills = m.DB.OpenRecordset("SELECT tblFills.* " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(8) & ";", dbOpenDynaset) ' eTT_AccountType_Alaron
    Do While Not rsFills.EOF
        If rsFills!FillQuantity < 0 Then
            rsFills.Delete
        Else
            rsFills.Edit
            rsFills!BrokerFillID = rsFills!BrokerOrderID & "-" & rsFills!BrokerFillID
            rsFills.Update
        End If
        
        rsFills.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion31"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion32
'' Description: Apply version 32 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion32()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rs As Recordset                 ' Recordset into the database
    Dim i As Index                      ' Index for a table in the database
    Dim r As Relation                   ' Relation object in the database
    Dim q As QueryDef                   ' Query def into the database
    
    ' Add the appropriate fields to the fills table...
    If ItemExists(m.DB.TableDefs, "tblFills") Then
        Set t = m.DB.TableDefs("tblFills")
        With t
            If Not ItemExists(.Fields, "EntryRuleIDCategory") Then
                .Fields.Append .CreateField("EntryRuleIDCategory", dbLong)
                .Fields("EntryRuleIDCategory").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblFills SET tblFills.EntryRuleIDCategory=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "ExitRuleIDCategory") Then
                .Fields.Append .CreateField("ExitRuleIDCategory", dbLong)
                .Fields("ExitRuleIDCategory").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblFills SET tblFills.ExitRuleIDCategory=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "RealFlagCategory") Then
                .Fields.Append .CreateField("RealFlagCategory", dbByte)
                .Fields("RealFlagCategory").DefaultValue = 0
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblFills SET tblFills.RealFlagCategory=0;"
                q.Execute
            End If
        End With
    End If
    
    If ItemExists(m.DB.TableDefs, "tblAccountPositionTrades") Then
        Set t = m.DB.TableDefs("tblAccountPositionTrades")
        With t
            If Not ItemExists(.Fields, "EntryFillID") Then
                .Fields.Append .CreateField("EntryFillID", dbLong)
                .Fields("EntryFillID").DefaultValue = 0&
                
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.EntryFillID=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "ExitFillID") Then
                .Fields.Append .CreateField("ExitFillID", dbLong)
                .Fields("ExitFillID").DefaultValue = 0&
                
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.ExitFillID=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "EntryRuleID") Then
                .Fields.Append .CreateField("EntryRuleID", dbLong)
                .Fields("EntryRuleID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.EntryRuleID=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "ExitRuleID") Then
                .Fields.Append .CreateField("ExitRuleID", dbLong)
                .Fields("ExitRuleID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.ExitRuleID=0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "RealFlag") Then
                .Fields.Append .CreateField("RealFlag", dbByte)
                .Fields("RealFlag").DefaultValue = 0
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.RealFlag=0;"
                q.Execute
            End If
            
            .Fields("AccountPositionID").OrdinalPosition = 0
            .Fields("TradeNumber").OrdinalPosition = 1
            .Fields("Direction").OrdinalPosition = 2
            .Fields("Quantity").OrdinalPosition = 3
            .Fields("EntryFillID").OrdinalPosition = 4
            .Fields("EntryBrokerFillID").OrdinalPosition = 5
            .Fields("EntryPrice").OrdinalPosition = 6
            .Fields("EntryTime").OrdinalPosition = 7
            .Fields("ExitFillID").OrdinalPosition = 8
            .Fields("ExitBrokerFillID").OrdinalPosition = 9
            .Fields("ExitPrice").OrdinalPosition = 10
            .Fields("ExitTime").OrdinalPosition = 11
            .Fields("ClosedProfit").OrdinalPosition = 12
            .Fields("EntryRuleID").OrdinalPosition = 13
            .Fields("ExitRuleID").OrdinalPosition = 14
            .Fields("RealFlag").OrdinalPosition = 15
            .Fields("IsOpen").OrdinalPosition = 16
        End With
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion32"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion33
'' Description: Apply version 33 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion33()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    
    If ItemExists(m.DB.TableDefs, "tblOrders") Then
        Set t = m.DB.TableDefs("tblOrders")
        With t
            If Not ItemExists(.Fields, "Exchange") Then
                .Fields.Append .CreateField("Exchange", dbText, 50)
                .Fields("Exchange").AllowZeroLength = True
                .Fields("Exchange").Required = False
                .Fields("Exchange").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblOrders SET tblOrders.Exchange='';"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion33"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion34
'' Description: Apply version 34 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion34()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblFills] " & _
                "WHERE [EntryRuleIDCategory]<>0 OR [ExitRuleIDCategory]<>0;", dbOpenDynaset)
    Do While Not rs.EOF
        ' If there is an entry rule set for this fill, walk through the account position trades and
        ' set the entry rule for the non-total category, and clear the one for the total category...
        If rs!EntryRuleIdCategory <> 0 Then
            Set rs2 = m.DB.OpenRecordset("SELECT tblAccountPositionTrades.*,tblAccountPositions.AutoTradeItemID " & _
                        "FROM [tblAccountPositionTrades] INNER JOIN [tblAccountPositions] ON tblAccountPositionTrades.AccountPositionID=tblAccountPositions.AccountPositionID " & _
                        "WHERE tblAccountPositionTrades.EntryFillID=" & rs!FillID & ";", dbOpenDynaset)
            Do While Not rs2.EOF
                rs2.Edit
                
                If rs2!AutoTradeItemID = -1& Then
                    rs2!EntryRuleID = 0&
                Else
                    rs2!EntryRuleID = rs!EntryRuleIdCategory
                End If
                
                rs2.Update
                
                rs2.MoveNext
            Loop
        End If
        
        ' If there is an exit rule set for this fill, walk through the account position trades and
        ' set the exit rule for the non-total category, and clear the one for the total category...
        If rs!ExitRuleIdCategory <> 0 Then
            Set rs2 = m.DB.OpenRecordset("SELECT tblAccountPositionTrades.*,tblAccountPositions.AutoTradeItemID " & _
                        "FROM [tblAccountPositionTrades] INNER JOIN [tblAccountPositions] ON tblAccountPositionTrades.AccountPositionID=tblAccountPositions.AccountPositionID " & _
                        "WHERE tblAccountPositionTrades.ExitFillID=" & rs!FillID & ";", dbOpenDynaset)
            Do While Not rs2.EOF
                rs2.Edit
                
                If rs2!AutoTradeItemID = -1& Then
                    rs2!ExitRuleID = 0&
                Else
                    rs2!ExitRuleID = rs!ExitRuleIdCategory
                End If
                
                rs2.Update
                
                rs2.MoveNext
            Loop
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion34"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion35
'' Description: Apply version 35 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion35()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim astrFills As cGdArray           ' Array of unique fill information
    Dim lPos As Long                    ' Position in the array
    Dim strKey As String                ' Key into the array
    Dim lNumber As Long                 ' Number of the copy
    Dim bFound As Boolean               ' Was the key found in the array?
    
    Set astrFills = New cGdArray
    astrFills.Create eGDARRAY_Strings
    
    Set rs = m.DB.OpenRecordset("SELECT tblFills.*, tblAccounts.AccountType " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID;", dbOpenDynaset)
    Do While Not rs.EOF
        strKey = Str(rs!AccountType) & vbTab & rs!BrokerFillID & vbTab
        
        bFound = astrFills.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength)
        Do While bFound
            lNumber = CLng(Val(Parse(astrFills(lPos), vbTab, 3)))
            astrFills(lPos) = strKey & Str(lNumber + 1)
            
            rs.Edit
            rs!BrokerFillID = rs!BrokerFillID & " (" & Str(lNumber) & ")"
            rs.Update
            
            strKey = Str(rs!AccountType) & vbTab & rs!BrokerFillID & vbTab
            bFound = astrFills.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength)
        Loop
        
        astrFills.Add strKey & "1", lPos
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion35"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion37
'' Description: Apply version 37 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion37()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim q As QueryDef                   ' Query def into the database
    Dim bFix As Boolean                 ' Do we need to fix this record?
    Dim bFixed As Boolean               ' Were any records fixed?
    
    bFixed = False
    
    Set rs = m.DB.OpenRecordset("SELECT tblFills.*, tblAccounts.AccountType " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(10) & ";", dbOpenDynaset)
'                "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_PFG) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!IsManual = False Then
            If InStr(rs!BrokerFillID, "-") = 0 Then
                bFix = False
                If Len(rs!BrokerFillID) < 9 Then
                    bFix = True
                Else
                    bFix = (UCase(Left(rs!BrokerFillID, 9)) <> "CARPOSFIX")
                End If
                
                If bFix Then
                    rs.Edit
                    rs!BrokerFillID = rs!BrokerOrderID & "-" & rs!BrokerFillID
                    rs.Update
                    
                    bFixed = True
                End If
            End If
        End If
        
        rs.MoveNext
    Loop

    ' If any records were fixed, set the entry fill ID on all records in tblAccountPositionTrades
    ' to zero so that PostFix32 will rebuild all of the account positions...
    If bFixed Then
        Set q = m.DB.CreateQueryDef("")
        q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.EntryFillID=0;"
        q.Execute
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion37"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion38
'' Description: Apply version 38 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion38()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    
    Set t = m.DB.TableDefs("tblOrderJournal")
    With t
        If Not ItemExists(.Fields, "Action") Then
            .Fields.Append .CreateField("Action", dbText, 10)
            .Fields("Action").AllowZeroLength = True
            .Fields("Action").Required = False
            .Fields("Action").DefaultValue = ""
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrderJournal SET tblOrderJournal.Action='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Feelings") Then
            .Fields.Append .CreateField("Feelings", dbText, 255)
            .Fields("Feelings").AllowZeroLength = True
            .Fields("Feelings").Required = False
            .Fields("Feelings").DefaultValue = ""
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrderJournal SET tblOrderJournal.Feelings='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "WhyTrade") Then
            .Fields.Append .CreateField("WhyTrade", dbText, 255)
            .Fields("WhyTrade").AllowZeroLength = True
            .Fields("WhyTrade").Required = False
            .Fields("WhyTrade").DefaultValue = ""
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrderJournal SET tblOrderJournal.WhyTrade='';"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Thoughts") Then
            .Fields.Append .CreateField("Thoughts", dbText, 255)
            .Fields("Thoughts").AllowZeroLength = True
            .Fields("Thoughts").Required = False
            .Fields("Thoughts").DefaultValue = ""
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrderJournal SET tblOrderJournal.Thoughts='';"
            q.Execute
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion38"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion39
'' Description: Apply version 39 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion39()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim i As Index                      ' Index for a table in the database
    Dim r As Relation                   ' Relation object in the database
    Dim q As QueryDef                   ' Query def into the database
    Dim rsOrders As Recordset           ' Recordset into the database
    Dim rsLegs As Recordset             ' Recordset into the database
    
    ' 1) Create the order legs table if it doesn't already exist...
    If Not ItemExists(m.DB.TableDefs, "tblOrderLegs") Then
        Set t = m.DB.CreateTableDef("tblOrderLegs")
        With t
            If Not ItemExists(.Fields, "OrderID") Then
                .Fields.Append .CreateField("OrderID", dbLong)
                .Fields("OrderID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "LegNumber") Then
                .Fields.Append .CreateField("LegNumber", dbLong)
                .Fields("OrderID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 255)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Exchange") Then
                .Fields.Append .CreateField("Exchange", dbText, 50)
                .Fields("Exchange").AllowZeroLength = True
                .Fields("Exchange").Required = False
                .Fields("Exchange").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "IsBuy") Then
                .Fields.Append .CreateField("IsBuy", dbBoolean)
                .Fields("IsBuy").DefaultValue = 0
            End If
            If Not ItemExists(.Fields, "Quantity") Then
                .Fields.Append .CreateField("Quantity", dbLong)
                .Fields("Quantity").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "LimitPrice") Then
                .Fields.Append .CreateField("LimitPrice", dbDouble)
                .Fields("LimitPrice").DefaultValue = 0
            End If
            If Not ItemExists(.Fields, "StopPrice") Then
                .Fields.Append .CreateField("StopPrice", dbDouble)
                .Fields("StopPrice").DefaultValue = 0
            End If
            If Not ItemExists(.Fields, "OrderType") Then
                .Fields.Append .CreateField("OrderType", dbLong)
                .Fields("OrderType").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "IsEntry") Then
                .Fields.Append .CreateField("IsEntry", dbBoolean)
                .Fields("IsEntry").DefaultValue = 0
            End If
            
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("OrderID")
                    .Fields.Append .CreateField("LegNumber")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    
        If Not ItemExists(m.DB.Relations, "OrderLegs") Then
            Set r = m.DB.CreateRelation("OrderLegs", "tblOrders", "tblOrderLegs", dbRelationDeleteCascade)
            With r
                .Fields.Append .CreateField("OrderID")
                .Fields.Item("OrderID").ForeignName = "OrderID"
            End With
            m.DB.Relations.Append r
        End If
    End If
    
    ' 2) Fill the order legs table with information from the orders table...
    Set rsLegs = m.DB.OpenRecordset("SELECT * FROM [tblOrderLegs];", dbOpenDynaset)
    If (rsLegs.BOF And rsLegs.EOF) Then
        Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
        Do While Not rsOrders.EOF
            Set rsLegs = m.DB.OpenRecordset("SELECT * FROM [tblOrderLegs] " & _
                        "WHERE [OrderID]=" & Str(rsOrders!OrderID) & " AND [LegNumber]=1;", dbOpenDynaset)
            If rsLegs.BOF And rsLegs.EOF Then
                rsLegs.AddNew
                
                rsLegs!OrderID = rsOrders!OrderID
                rsLegs!LegNumber = 1
                rsLegs!SymbolID = rsOrders!SymbolID
                rsLegs!Symbol = rsOrders!Symbol
                rsLegs!Exchange = rsOrders!Exchange
                rsLegs!IsBuy = rsOrders!BuySell
                rsLegs!Quantity = rsOrders!Quantity
                rsLegs!LimitPrice = rsOrders!LimitPrice
                rsLegs!StopPrice = rsOrders!StopPrice
                rsLegs!OrderType = rsOrders!OrderType
                rsLegs!IsEntry = rsOrders!EnterExit
                
                rsLegs.Update
            End If
            
            rsOrders.MoveNext
        Loop
    End If
    
    ' Close the recordset for the orders table so that we can remove the fields...
    If Not rsOrders Is Nothing Then
        rsOrders.Close
    End If
    
    ' 3) Remove leg fields from the orders table...
    Set t = m.DB.TableDefs("tblOrders")
    With t
        If ItemExists(.Fields, "SymbolID") Then
            .Fields.Delete "SymbolID"
        End If
        If ItemExists(.Fields, "Symbol") Then
            .Fields.Delete "Symbol"
        End If
        If ItemExists(.Fields, "BuySell") Then
            .Fields.Delete "BuySell"
        End If
        If ItemExists(.Fields, "Quantity") Then
            .Fields.Delete "Quantity"
        End If
        If ItemExists(.Fields, "LimitPrice") Then
            .Fields.Delete "LimitPrice"
        End If
        If ItemExists(.Fields, "StopPrice") Then
            .Fields.Delete "StopPrice"
        End If
        If ItemExists(.Fields, "OrderType") Then
            .Fields.Delete "OrderType"
        End If
        If ItemExists(.Fields, "EnterExit") Then
            .Fields.Delete "EnterExit"
        End If
        If ItemExists(.Fields, "Exchange") Then
            .Fields.Delete "Exchange"
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion39"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion40
'' Description: Apply version 40 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion40()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rsOrders As Recordset           ' Recordset into the database

    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "GroupID") Then
            .Fields.Append .CreateField("GroupID", dbLong)
            .Fields("GroupID").DefaultValue = 0&
        End If
        If Not ItemExists(.Fields, "GroupName") Then
            .Fields.Append .CreateField("GroupName", dbText, 255)
            .Fields("GroupName").AllowZeroLength = True
            .Fields("GroupName").Required = False
            .Fields("GroupName").DefaultValue = ""
        End If
    End With
    
    Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
    Do While Not rsOrders.EOF
        rsOrders.Edit
        rsOrders!GroupID = rsOrders!OrderID
        rsOrders!GroupName = "Single Order"
        rsOrders.Update
        
        rsOrders.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion40"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion41
'' Description: Apply version 41 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion41()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    
    Set t = m.DB.TableDefs("tblAccountPositionTrades")
    With t
        If Not ItemExists(.Fields, "MaxProfit") Then
            .Fields.Append .CreateField("MaxProfit", dbDouble)
            .Fields("MaxProfit").DefaultValue = 0#
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.MaxProfit = 0.0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "MaxLoss") Then
            .Fields.Append .CreateField("MaxLoss", dbDouble)
            .Fields("MaxLoss").DefaultValue = 0#
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.MaxLoss = 0.0;"
            q.Execute
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion41"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion42
'' Description: Apply version 42 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion42()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    
    Set t = m.DB.TableDefs("tblOrderJournal")
    With t
        If Not ItemExists(.Fields, "EmotionNumber") Then
            .Fields.Append .CreateField("EmotionNumber", dbLong)
            .Fields("EmotionNumber").DefaultValue = -1&
            
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrderJournal SET tblOrderJournal.EmotionNumber = -1;"
            q.Execute
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion42"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion43
'' Description: Apply version 43 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion43()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rsJournals As Recordset         ' Recordset into the database
    Dim rsOrder As Recordset            ' Recordset into the database
    
    Set t = m.DB.TableDefs("tblOrderJournal")
    With t
        If Not ItemExists(.Fields, "AccountID") Then
            .Fields.Append .CreateField("AccountID", dbLong)
            .Fields("AccountID").DefaultValue = 0&
        End If
        
        Set rsJournals = m.DB.OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
        Do While Not rsJournals.EOF
            rsJournals.Edit
        
            Set rsOrder = m.DB.OpenRecordset("SELECT * FROM [tblOrders] " & _
                        "WHERE [OrderID]=" & Str(rsJournals!OrderID) & ";", dbOpenDynaset)
            If rsOrder.BOF And rsOrder.EOF Then
                rsJournals!AccountID = 0&
            Else
                rsJournals!AccountID = rsOrder!AccountID
            End If
            
            If IsNull(rsJournals!Action) Then
                rsJournals!Action = ""
            End If
            If IsNull(rsJournals!Feelings) Then
                rsJournals!Feelings = ""
            End If
            If IsNull(rsJournals!WhyTrade) Then
                rsJournals!WhyTrade = ""
            End If
            If IsNull(rsJournals!Thoughts) Then
                rsJournals!Thoughts = ""
            End If
            
            rsJournals.Update
            
            rsJournals.MoveNext
        Loop
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion43"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion44
'' Description: Apply version 44 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion44()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    Do While Not rs.EOF
        Select Case rs!AccountType
            Case eTT_AccountType_SimTrade, eTT_AccountType_Standard
                If rs!SecTypeMask = 7 Then
                    rs.Edit
                    rs!SecTypeMask = 31
                    rs.Update
                End If
            
            Case 10 ' eTT_AccountType_PFG
                If rs!SecTypeMask = 5 Then
                    rs.Edit
                    rs!SecTypeMask = 13
                    rs.Update
                End If
            
            Case 5, 15 ' eTT_AccountType_LindWaldock, eTT_AccountType_ManExpress
                If rs!SecTypeMask = 1 Then
                    rs.Edit
                    rs!SecTypeMask = 9
                    rs.Update
                End If
            
        End Select
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion44"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion45
'' Description: Apply version 45 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion45()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
    Do While Not rs.EOF
        Select Case rs!AccountType
            Case 10 'eTT_AccountType_PFG
                If rs!SecTypeMask = 13 Then
                    rs.Edit
                    rs!SecTypeMask = 31
                    rs.Update
                End If
            
        End Select
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion45"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion46
'' Description: Apply version 46 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion46()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    
    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "IsAutoExit") Then
            .Fields.Append .CreateField("IsAutoExit", dbBoolean)
            .Fields("IsAutoExit").DefaultValue = False
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrders SET tblOrders.IsAutoExit = 0;"
            q.Execute
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion46"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion47
'' Description: Apply version 47 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion47()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders] " & _
                "WHERE [Status] = " & Str(eTT_OrderStatus_Open) & ";")
    Do While Not rs.EOF
        rs.Edit
        rs!Status = eTT_OrderStatus_Parked
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTTUpdates.ApplyVersion47"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion48
'' Description: Apply version 48 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion48()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    Dim i As Index                      ' Index object in the database
    Dim rs As Recordset                 ' Recordset into the database
    
    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "BrokerCancelOrderID") Then
            .Fields.Append .CreateField("BrokerCancelOrderID", dbLong)
            .Fields("BrokerCancelOrderID").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblOrders SET tblOrders.BrokerCancelOrderID = 0;"
            q.Execute
        End If
    End With
    
    Set t = m.DB.TableDefs("tblAccounts")
    With t
        If Not ItemExists(.Fields, "HoldOcoAtBroker") Then
            .Fields.Append .CreateField("HoldOcoAtBroker", dbBoolean)
            .Fields("HoldOcoAtBroker").DefaultValue = False
            
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts];", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Edit
                rs!HoldOcoAtBroker = (rs!AccountType = 10) 'eTT_AccountType_PFG)
                rs.Update
                
                rs.MoveNext
            Loop
        End If
    End With

    If Not ItemExists(m.DB.TableDefs, "tblOrderLinks") Then
        Set t = m.DB.CreateTableDef("tblOrderLinks")
        With t
            If Not ItemExists(.Fields, "OrderLinkID") Then
                .Fields.Append .CreateField("OrderLinkID", dbLong)
                .Fields("OrderLinkID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "OrderID1") Then
                .Fields.Append .CreateField("OrderID1", dbLong)
                .Fields("OrderID1").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "BrokerID1") Then
                .Fields.Append .CreateField("BrokerID1", dbText, 255)
                .Fields("BrokerID1").AllowZeroLength = True
                .Fields("BrokerID1").Required = False
                .Fields("BrokerID1").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "OrderID2") Then
                .Fields.Append .CreateField("OrderID2", dbLong)
                .Fields("OrderID2").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "BrokerID2") Then
                .Fields.Append .CreateField("BrokerID2", dbText, 255)
                .Fields("BrokerID2").AllowZeroLength = True
                .Fields("BrokerID2").Required = False
                .Fields("BrokerID2").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Status") Then
                .Fields.Append .CreateField("Status", dbLong)
                .Fields("Status").DefaultValue = 0&
            End If
                    
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("OrderLinkID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cTTUpdates.ApplyVersion48"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion49
'' Description: Apply version 49 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion49()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database

    Set t = m.DB.TableDefs("tblAccounts")
    With t
        If Not ItemExists(.Fields, "StockFees") Then
            .Fields.Append .CreateField("StockFees", dbDouble)
            .Fields("StockFees").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccounts SET tblAccounts.StockFees = 0;"
            q.Execute
        End If
    End With

    Set t = m.DB.TableDefs("tblAccountPositionTrades")
    With t
        If Not ItemExists(.Fields, "Fees") Then
            .Fields.Append .CreateField("Fees", dbDouble)
            .Fields("Fees").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblAccountPositionTrades SET tblAccountPositionTrades.Fees = 0;"
            q.Execute
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion49"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion50
'' Description: Apply version 50 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion50()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database

    Set t = m.DB.TableDefs("tblActiveExits")
    With t
        If Not ItemExists(.Fields, "Stop2OrderID") Then
            .Fields.Append .CreateField("Stop2OrderID", dbLong)
            .Fields("Stop2OrderID").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblActiveExits SET tblActiveExits.Stop2OrderID = 0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "Stop3OrderID") Then
            .Fields.Append .CreateField("Stop3OrderID", dbLong)
            .Fields("Stop3OrderID").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblActiveExits SET tblActiveExits.Stop3OrderID = 0;"
            q.Execute
        End If
        If Not ItemExists(.Fields, "StopROrderID") Then
            .Fields.Append .CreateField("StopROrderID", dbLong)
            .Fields("StopROrderID").DefaultValue = 0
        
            Set q = m.DB.CreateQueryDef("")
            q.SQL = "UPDATE tblActiveExits SET tblActiveExits.StopROrderID = 0;"
            q.Execute
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion50"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion51
'' Description: Apply version 51 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion51()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object

    Set t = m.DB.TableDefs("tblFills")
    With t
        If ItemExists(.Fields, "Fees") Then
            .Fields("Fees").DefaultValue = kNullData
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion51"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion52
'' Description: Apply version 52 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion52()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rs As Recordset                 ' Recordset into the database

    Set t = m.DB.TableDefs("tblOrders")
    With t
        If Not ItemExists(.Fields, "IsAutomated") Then
            .Fields.Append .CreateField("IsAutomated", dbBoolean)
            .Fields("IsAutomated").DefaultValue = False
        
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders];", dbOpenDynaset)
            Do While Not rs.EOF
                rs.Edit
                rs!IsAutomated = (rs!AutoTradeItemID <> 0) Or (rs!IsAutoExit <> 0)
                rs.Update
                
                rs.MoveNext
            Loop
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion52"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion53
'' Description: Apply version 53 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion53()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccounts] " & _
            "WHERE [AccountType]=" & Str(eTT_AccountType_SimTrade) & " OR [AccountType]=" & Str(eTT_AccountType_Standard) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        
        If (Len(rs!AccountNumber) = 10) And (UCase(Left(rs!AccountNumber, 5)) = "GENSR") Then
            rs!AccountType = eTT_AccountType_SimReplay
        ElseIf (HasModule("RTG") = False) Or HasWorkingTradeServerOrders(rs!AccountID) Then
            rs!AccountType = eTT_AccountType_SimBroker
        Else
            rs!AccountType = eTT_AccountType_SimStream
        End If
        
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion53"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasWorkingTradeServerOrders
'' Description: Does the given account have working trade server orders?
'' Inputs:      Account ID
'' Returns:     True if has working Trade Server orders, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasWorkingTradeServerOrders(ByVal lAccountID As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim rs As Recordset                 ' Recordset into the database
    
    bReturn = False
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrders] WHERE [AccountID]=" & Str(lAccountID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If (rs!Status <> eTT_OrderStatus_Parked) And (IsOpenOrder(rs!Status, False) = True) Then
            If UCase(Left(rs!BrokerOrderID, 3)) <> "DBO" Then
                bReturn = True
                Exit Do
            End If
        End If
        
        rs.MoveNext
    Loop
    
    HasWorkingTradeServerOrders = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTTUpdates.HasWorkingTradeServerOrders"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion54
'' Description: Apply version 54 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion54()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    Dim i As Index                      ' Index object in the database
    Dim rs As Recordset                 ' Recordset into the database
    
    If Not ItemExists(m.DB.TableDefs, "tblDateJournals") Then
        Set t = m.DB.CreateTableDef("tblDateJournals")
        With t
            If Not ItemExists(.Fields, "DateJournalID") Then
                .Fields.Append .CreateField("DateJournalID", dbLong)
                .Fields("DateJournalID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "JournalDate") Then
                .Fields.Append .CreateField("JournalDate", dbDouble)
                .Fields("JournalDate").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "JournalCategoryID") Then
                .Fields.Append .CreateField("JournalCategoryID", dbLong)
                .Fields("JournalCategoryID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "JournalTime") Then
                .Fields.Append .CreateField("JournalTime", dbDouble)
                .Fields("JournalTime").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "Text") Then
                .Fields.Append .CreateField("Text", dbText, 255)
                .Fields("Text").AllowZeroLength = True
                .Fields("Text").Required = False
                .Fields("Text").DefaultValue = ""
            End If
                    
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("DateJournalID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If
    
    If ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
        Set t = m.DB.TableDefs("tblOrderJournal")
        With t
            If Not ItemExists(.Fields, "JournalDate") Then
                .Fields.Append .CreateField("JournalDate", dbDouble)
                .Fields("JournalDate").DefaultValue = 0#
            End If
        End With
        
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!JournalDate = CDbl(Int(rs!NoteDate))
            rs.Update
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion54"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion55
'' Description: Apply version 55 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion55()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rs As Recordset                 ' Recordset into the database
    
    If ItemExists(m.DB.TableDefs, "tblDateJournals") Then
        Set t = m.DB.TableDefs("tblDateJournals")
        With t
            If Not ItemExists(.Fields, "ImageFilename") Then
                .Fields.Append .CreateField("ImageFilename", dbText, 255)
                .Fields("ImageFilename").AllowZeroLength = True
                .Fields("ImageFilename").Required = False
                .Fields("ImageFilename").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "ImageCaption") Then
                .Fields.Append .CreateField("ImageCaption", dbText, 255)
                .Fields("ImageCaption").AllowZeroLength = True
                .Fields("ImageCaption").Required = False
                .Fields("ImageCaption").DefaultValue = ""
            End If
        End With
        
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblDateJournals];", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!ImageFilename = ""
            rs!ImageCaption = ""
            rs.Update
            
            rs.MoveNext
        Loop
    End If
    
    If ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
        Set t = m.DB.TableDefs("tblOrderJournal")
        With t
            If Not ItemExists(.Fields, "ImageFilename") Then
                .Fields.Append .CreateField("ImageFilename", dbText, 255)
                .Fields("ImageFilename").AllowZeroLength = True
                .Fields("ImageFilename").Required = False
                .Fields("ImageFilename").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "ImageCaption") Then
                .Fields.Append .CreateField("ImageCaption", dbText, 255)
                .Fields("ImageCaption").AllowZeroLength = True
                .Fields("ImageCaption").Required = False
                .Fields("ImageCaption").DefaultValue = ""
            End If
        End With
        
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!ImageFilename = ""
            rs!ImageCaption = ""
            rs.Update
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion55"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion56
'' Description: Apply version 56 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion56()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim i As Index                      ' Index object in the database

    ChangeDataType "tblDateJournals", "Text", dbMemo

    If ItemExists(m.DB.TableDefs, "tblDateJournals") Then
        Set t = m.DB.TableDefs("tblDateJournals")
        With t
            If Not ItemExists(.Indexes, "JournalDate") Then
                Set i = .CreateIndex("JournalDate")
                With i
                    .Fields.Append .CreateField("JournalDate")
                    .Primary = False
                End With
                .Indexes.Append i
            End If
        End With
    End If

    If ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
        Set t = m.DB.TableDefs("tblOrderJournal")
        With t
            If Not ItemExists(.Indexes, "JournalDate") Then
                Set i = .CreateIndex("JournalDate")
                With i
                    .Fields.Append .CreateField("JournalDate")
                    .Primary = False
                End With
                .Indexes.Append i
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion56"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion57
'' Description: Apply version 57 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion57()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim dOrderDate As Double            ' Order date
    Dim alToDelete As cGdArray          ' Array of ID's to delete
    Dim astrIds As cGdArray             ' Array of ID's in the recordset
    Dim lPos As Long                    ' Position to insert into the array
    
    Set alToDelete = New cGdArray
    alToDelete.Create eGDARRAY_Longs
    
    Set astrIds = New cGdArray
    astrIds.Create eGDARRAY_Strings
    
    ' 1) Fix the BrokerOrderID and PreviousBrokerID fields in the orders table...
    Set rs = m.DB.OpenRecordset("SELECT tblOrders.* " & _
                "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_IntBrokers) & " OR tblAccounts.AccountType=" & Str(eTT_AccountType_Ideal) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        If (InStr(rs!BrokerOrderID, "-") > 0) Then
            rs!BrokerOrderID = Parse(rs!BrokerOrderID, "-", 2)
        End If
        If (InStr(rs!PreviousBrokerID, "-") > 0) Then
            rs!PreviousBrokerID = Parse(rs!PreviousBrokerID, "-", 2)
        End If
        rs.Update
        
        rs.MoveNext
    Loop

    ' 2) Fix the duplicates in the orders table...
    astrIds.Clear
    Set rs = m.DB.OpenRecordset("SELECT tblOrders.* " & _
                "FROM tblOrders INNER JOIN tblAccounts ON tblOrders.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_IntBrokers) & " OR tblAccounts.AccountType=" & Str(eTT_AccountType_Ideal) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        dOrderDate = 0
        If Len(rs!BrokerOrderID) > 0 Then
            If (rs!BrokerOrderID = rs!PreviousBrokerID) Then
                Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblOrders] " & _
                            "WHERE [GenesisOrderID]='' AND [BrokerOrderID]='" & rs!BrokerOrderID & "';", dbOpenDynaset)
                Do While Not rs2.EOF
                    If rs2!OrderID <> rs!OrderID Then
                        If (dOrderDate = 0) And (rs2!OrderDate <> 0) Then
                            dOrderDate = rs2!OrderDate
                        End If
                        rs2.Delete
                    End If
                    
                    rs2.MoveNext
                Loop
                
                rs.Edit
                rs!PreviousBrokerID = ""
                rs!OrderDate = dOrderDate
                rs.Update
            Else
                If astrIds.BinarySearch(rs!BrokerOrderID, lPos) = True Then
                    rs.Delete
                Else
                    astrIds.Add rs!BrokerOrderID, lPos
                End If
            End If
        End If
        
        rs.MoveNext
    Loop

    ' 3) Fix the BrokerOrderID field in the fills table...
    alToDelete.Clear
    Set rs = m.DB.OpenRecordset("SELECT tblFills.* " & _
                "FROM tblFills INNER JOIN tblAccounts ON tblFills.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_IntBrokers) & " OR tblAccounts.AccountType=" & Str(eTT_AccountType_Ideal) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If alToDelete.BinarySearch(rs!FillID) = False Then
            If UCase(Left(rs!BrokerFillID, 10)) = "CARPOSFIX " Then
                If alToDelete.BinarySearch(rs!FillID, lPos) = False Then
                    alToDelete.Add rs!FillID, lPos
                End If
            Else
                rs.Edit
                If (InStr(rs!BrokerOrderID, "-") > 0) Then
                    rs!BrokerOrderID = Parse(rs!BrokerOrderID, "-", 2)
                End If
                If Len(rs!BrokerOrderID) > 0 Then
                    Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblOrders] " & _
                                "WHERE [BrokerOrderID]='" & Str(rs!BrokerOrderID) & "';", dbOpenDynaset)
                    If (rs2.BOF And rs2.EOF) Then
                        rs!OrderID = 0
                    Else
                        rs!OrderID = rs2!OrderID
                    End If
                Else
                    rs!OrderID = 0
                End If
                rs.Update
                
                Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblFills] " & _
                            "WHERE [BrokerFillID]='" & rs!BrokerFillID & "';", dbOpenDynaset)
                Do While Not rs2.EOF
                    If rs2!FillID <> rs!FillID Then
                        If alToDelete.BinarySearch(rs2!FillID, lPos) = False Then
                            alToDelete.Add rs2!FillID, lPos
                        End If
                    End If
                    rs2.MoveNext
                Loop
            End If
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblFills];", dbOpenDynaset)
    Do While Not rs.EOF
        If alToDelete.BinarySearch(rs!FillID) = True Then
            rs.Delete
        End If
        
        rs.MoveNext
    Loop
    
    ' 4) Set AccountPosition objects so that they will get recalculated...
    Set rs = m.DB.OpenRecordset("SELECT tblAccountPositions.* " & _
                "FROM tblAccountPositions INNER JOIN tblAccounts ON tblAccountPositions.AccountID=tblAccounts.AccountID " & _
                "WHERE tblAccounts.AccountType=" & Str(eTT_AccountType_IntBrokers) & " OR tblAccounts.AccountType=" & Str(eTT_AccountType_Ideal) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        rs!FillMatchMode = rs!FillMatchMode + 254
        rs.Update
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion57"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion58
'' Description: Apply version 58 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion58()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database

    If ItemExists(m.DB.TableDefs, "tblAccountPositions") Then
        Set t = m.DB.TableDefs("tblAccountPositions")
        With t
            If Not ItemExists(.Fields, "ExpirationDate") Then
                .Fields.Append .CreateField("ExpirationDate", dbLong)
                .Fields("ExpirationDate").DefaultValue = 0&
        
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAccountPositions SET tblAccountPositions.ExpirationDate = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion58"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion59
'' Description: Apply version 59 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion59()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef                   ' Query def into the database
    Dim i As Index                      ' Index object in the database
    Dim r As Relation                   ' Relationship in the database
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    
    If Not ItemExists(m.DB.TableDefs, "tblJournalImages") Then
        Set t = m.DB.CreateTableDef("tblJournalImages")
        With t
            If Not ItemExists(.Fields, "ID") Then
                .Fields.Append .CreateField("ID", dbLong)
                .Fields("ID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "Type") Then
                .Fields.Append .CreateField("Type", dbLong)
                .Fields("Type").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Filename") Then
                .Fields.Append .CreateField("Filename", dbText, 255)
                .Fields("Filename").AllowZeroLength = True
                .Fields("Filename").Required = False
                .Fields("Filename").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Caption") Then
                .Fields.Append .CreateField("Caption", dbText, 255)
                .Fields("Caption").AllowZeroLength = True
                .Fields("Caption").Required = False
                .Fields("Caption").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "OrderJournalID") Then
                .Fields.Append .CreateField("OrderJournalID", dbLong)
                .Fields("OrderJournalID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "DateJournalID") Then
                .Fields.Append .CreateField("DateJournalID", dbLong)
                .Fields("DateJournalID").DefaultValue = 0&
            End If
                    
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("ID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        
        m.DB.TableDefs.Append t
    End If
    
    Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblJournalImages];", dbOpenDynaset)
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
    Do While Not rs.EOF
        If Len(rs!ImageFilename) > 0 Then
            rs2.AddNew
            rs2!Type = 0 'eGDJournalImageType_Chart
            rs2!FileName = rs!ImageFilename
            rs2!Caption = rs!ImageCaption
            rs2!OrderJournalID = rs!JournalID
            rs2!DateJournalID = 0&
            rs2.Update
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblDateJournals];", dbOpenDynaset)
    Do While Not rs.EOF
        If Len(rs!ImageFilename) > 0 Then
            rs2.AddNew
            If UCase(FileExt(rs!ImageFilename)) = "JPG" Then
                rs2!Type = 0 'eGDJournalImageType_Chart
            Else
                rs2!Type = 1 'eGDJournalImageType_SummaryReport
            End If
            rs2!FileName = rs!ImageFilename
            rs2!Caption = rs!ImageCaption
            rs2!OrderJournalID = 0&
            rs2!DateJournalID = rs!DateJournalID
            rs2.Update
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    Set rs2 = Nothing
    
    If ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
        Set t = m.DB.TableDefs("tblOrderJournal")
        With t
            If ItemExists(.Fields, "ImageFilename") Then
                .Fields.Delete "ImageFilename"
            End If
            If ItemExists(.Fields, "ImageCaption") Then
                .Fields.Delete "ImageCaption"
            End If
        End With
    End If
    
    If ItemExists(m.DB.TableDefs, "tblDateJournals") Then
        Set t = m.DB.TableDefs("tblDateJournals")
        With t
            If ItemExists(.Fields, "ImageFilename") Then
                .Fields.Delete "ImageFilename"
            End If
            If ItemExists(.Fields, "ImageCaption") Then
                .Fields.Delete "ImageCaption"
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion59"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion60
'' Description: Apply version 60 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion60()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rsOrders As Recordset           ' Recordset into the orders table
    Dim rsOrderLegs As Recordset        ' Recordset into the order legs table

    ' 1) Create new fields in tblOrders...
    If ItemExists(m.DB.TableDefs, "tblOrders") Then
        Set t = m.DB.TableDefs("tblOrders")
        With t
            If Not ItemExists(.Fields, "OrderType") Then
                .Fields.Append .CreateField("OrderType", dbLong)
                .Fields("OrderType").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "LimitPrice") Then
                .Fields.Append .CreateField("LimitPrice", dbDouble)
                .Fields("LimitPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "StopPrice") Then
                .Fields.Append .CreateField("StopPrice", dbDouble)
                .Fields("StopPrice").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "Quantity") Then
                .Fields.Append .CreateField("Quantity", dbLong)
                .Fields("Quantity").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "UnderlyingSymbol") Then
                .Fields.Append .CreateField("UnderlyingSymbol", dbText, 255)
                .Fields("UnderlyingSymbol").AllowZeroLength = True
                .Fields("UnderlyingSymbol").Required = False
                .Fields("UnderlyingSymbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "UnderlyingSymbolID") Then
                .Fields.Append .CreateField("UnderlyingSymbolID", dbLong)
                .Fields("UnderlyingSymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "DebitCredit") Then
                .Fields.Append .CreateField("DebitCredit", dbLong)
                .Fields("DebitCredit").DefaultValue = 0
            End If
        End With
    End If
    
    ' 2) Move values from fields in tblOrderLegs to new fields in tblOrders...
    Set rsOrderLegs = m.DB.OpenRecordset("SELECT * FROM [tblOrderLegs];", dbOpenDynaset)
    Do While Not rsOrderLegs.EOF
        Set rsOrders = m.DB.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(rsOrderLegs!OrderID) & ";", dbOpenDynaset)
        If Not (rsOrders.BOF And rsOrders.EOF) Then
            rsOrders.Edit
            rsOrders!OrderType = rsOrderLegs!OrderType
            rsOrders!LimitPrice = rsOrderLegs!LimitPrice
            rsOrders!StopPrice = rsOrderLegs!StopPrice
            rsOrders!Quantity = rsOrderLegs!Quantity
            rsOrders!UnderlyingSymbol = rsOrderLegs!Symbol
            rsOrders!UnderlyingSymbolID = rsOrderLegs!SymbolID
            rsOrders!DebitCredit = 0
            rsOrders.Update
        End If
        
        rsOrderLegs.Edit
        If rsOrderLegs!LegNumber = 0 Then
            rsOrderLegs!LegNumber = 1
        End If
        rsOrderLegs!Quantity = 1
        rsOrderLegs.Update
        
        rsOrderLegs.MoveNext
    Loop
    
    ' Need to set these to nothing so that the recordset doesn't have tblOrderLegs locked for the
    ' calls below...
    Set rsOrderLegs = Nothing
    Set rsOrders = Nothing

    ' 3) Remove/rename fields from tblOrderLegs...
    If ItemExists(m.DB.TableDefs, "tblOrderLegs") Then
        Set t = m.DB.TableDefs("tblOrderLegs")
        With t
            If ItemExists(.Fields, "OrderType") Then
                .Fields.Delete "OrderType"
            End If
            If ItemExists(.Fields, "LimitPrice") Then
                .Fields.Delete "LimitPrice"
            End If
            If ItemExists(.Fields, "StopPrice") Then
                .Fields.Delete "StopPrice"
            End If
            If ItemExists(.Fields, "Quantity") Then
                .Fields("Quantity").Name = "Multiplier"
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion60"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion61
'' Description: Apply version 61 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion61()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim rs As Recordset                 ' Recordset into the database
    Dim moneyCodeFields As cGdTree      ' Dictionary of Money Code fields
    Dim rsOrder As Recordset            ' Recordset of an order from the database
    Dim rsOrderLegs As Recordset        ' Recordset of order legs from the database
    Dim lNumLegs As Long                ' Number of legs for the order

    If ItemExists(m.DB.TableDefs, "tblDateJournals") Then
        Set t = m.DB.TableDefs("tblDateJournals")
        With t
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
        End With
    End If
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblDateJournals];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        
        rs!SymbolID = 0
        rs!Symbol = ""
        
        If rs!JournalCategoryID = 4 Then
            If UCase(Left(rs!Text, 8)) = "VERSION=" Then
                Set moneyCodeFields = New cGdTree
                moneyCodeFields.FromKeyValueString rs!Text, "|", "="
                
                If moneyCodeFields.Exists("Symbol") Then
                    rs!SymbolID = kNullData
                    rs!Symbol = moneyCodeFields("Symbol")
                End If
            End If
        End If
        
        rs.Update
        
        rs.MoveNext
    Loop

    If ItemExists(m.DB.TableDefs, "tblOrderJournal") Then
        Set t = m.DB.TableDefs("tblOrderJournal")
        With t
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
        End With
    End If
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblOrderJournal];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        
        rs!SymbolID = 0
        rs!Symbol = ""
        
        Set rsOrder = m.DB.OpenRecordset("SELECT * FROM [tblOrders] WHERE [OrderID]=" & Str(rs!OrderID) & ";", dbOpenDynaset)
        If Not (rsOrder.BOF And rsOrder.EOF) Then
            Set rsOrderLegs = m.DB.OpenRecordset("SELECT * FROM [tblOrderLegs] WHERE [OrderID]=" & Str(rs!OrderID) & ";", dbOpenDynaset)
            If Not (rsOrderLegs.BOF And rsOrderLegs.EOF) Then
                rsOrderLegs.MoveLast
                lNumLegs = rsOrderLegs.RecordCount
                
                If lNumLegs = 1 Then
                    rs!SymbolID = rsOrderLegs!SymbolID
                    rs!Symbol = rsOrderLegs!Symbol
                End If
            End If
        End If
        
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion61"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion62
'' Description: Apply version 62 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion62()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim i As Index                      ' Index object in the database
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database

    If Not ItemExists(m.DB.TableDefs, "tblSymbolInfo") Then
        Set t = m.DB.CreateTableDef("tblSymbolInfo")
        With t
            If Not ItemExists(.Fields, "SymbolInfoID") Then
                .Fields.Append .CreateField("SymbolInfoID", dbLong)
                .Fields("SymbolInfoID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "ExpirationDate") Then
                .Fields.Append .CreateField("ExpirationDate", dbDouble)
                .Fields("ExpirationDate").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "LotSize") Then
                .Fields.Append .CreateField("LotSize", dbDouble)
                .Fields("LotSize").DefaultValue = 0#
            End If
                                
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("SymbolInfoID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            
            m.DB.TableDefs.Append t
        End With
        
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAccountPositions] WHERE [ExpirationDate]>0;", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!SymbolID = 0& Then
                Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [Symbol]='" & rs!Symbol & "';", dbOpenDynaset)
            Else
                Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblSymbolInfo] WHERE [SymbolID]=" & rs!SymbolID & ";", dbOpenDynaset)
            End If
            
            If rs2.BOF And rs2.EOF Then
                rs2.AddNew
            Else
                rs2.Edit
            End If
            
            rs2!Symbol = rs!Symbol
            rs2!SymbolID = rs!SymbolID
            rs2!ExpirationDate = rs!ExpirationDate
            
            If UCase(SecurityType(rs!Symbol, True)) = "SO" Then
                rs2!LotSize = 100
            Else
                rs2!LotSize = 1
            End If
            
            rs2.Update
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion62"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion63
'' Description: Apply version 63 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion63()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object

    If ItemExists(m.DB.TableDefs, "tblAccounts") Then
        Set t = m.DB.TableDefs("tblAccounts")
        With t
            If Not ItemExists(.Fields, "FcmAccountNumber") Then
                .Fields.Append .CreateField("FcmAccountNumber", dbText, 50)
                .Fields("FcmAccountNumber").AllowZeroLength = True
                .Fields("FcmAccountNumber").Required = False
                .Fields("FcmAccountNumber").DefaultValue = ""
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion63"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion64
'' Description: Apply version 64 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion64()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef

    If ItemExists(m.DB.TableDefs, "tblOrders") Then
        Set t = m.DB.TableDefs("tblOrders")
        With t
            If Not ItemExists(.Fields, "ContingencyOptions") Then
                .Fields.Append .CreateField("ContingencyOptions", dbText, 255)
                .Fields("ContingencyOptions").AllowZeroLength = True
                .Fields("ContingencyOptions").Required = False
                .Fields("ContingencyOptions").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblOrders SET tblOrders.ContingencyOptions = '';"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion64"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion65
'' Description: Apply version 65 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion65()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef

    If ItemExists(m.DB.TableDefs, "tblOrders") Then
        Set t = m.DB.TableDefs("tblOrders")
        With t
            If Not ItemExists(.Fields, "OcaGroup") Then
                .Fields.Append .CreateField("OcaGroup", dbText, 255)
                .Fields("OcaGroup").AllowZeroLength = True
                .Fields("OcaGroup").Required = False
                .Fields("OcaGroup").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblOrders SET tblOrders.OcaGroup = '';"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion65"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion66
'' Description: Apply version 66 to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion66()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Table definition object
    Dim q As QueryDef

    If ItemExists(m.DB.TableDefs, "tblFills") Then
        Set t = m.DB.TableDefs("tblFills")
        With t
            If Not ItemExists(.Fields, "IsSpread") Then
                .Fields.Append .CreateField("IsSpread", dbBoolean)
                .Fields("IsSpread").DefaultValue = 0
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblFills SET tblFills.IsSpread = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ApplyVersion66"

End Sub
