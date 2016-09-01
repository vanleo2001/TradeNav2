VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUniverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' This class opens the Symbol Universe DB
' (initializes handles for tables, fields, and tags)
Option Explicit

' SYMBOLS.DBF
Public tblSymbols&
Public fldSymbolID&, fldSymbol&, fldDesc&
Public fldSecType&, fldEodFirst&, fldEodLast&, fldTickFirst&, fldTickLast&
Public fldAccess&
Public tagSymbol&, tagAccessBySymbol&, tagAccessByDesc&, tagSymbolID&

' EXCHANGE.DBF
Public tblExchanges&

' MARKETS.DBF
Public tblMarkets&

Public Function OpenDb() As Boolean
On Error GoTo ErrSection:

    Dim rc&, bSuccess As Boolean
    Dim strPath$
    
    tblSymbols = TblOpen(DataPath & "Symbols.dbf") ', False, True)
    If tblSymbols Then
        bSuccess = True
            
        fldSymbolID = d4field(tblSymbols, "SymbolId")
        fldSymbol = d4field(tblSymbols, "Symbol")
        fldDesc = d4field(tblSymbols, "Desc")
    
        fldSecType = d4field(tblSymbols, "SecType")
        fldAccess = d4field(tblSymbols, "Access")
        
        fldEodFirst = d4field(tblSymbols, "EodFirst")
        fldEodLast = d4field(tblSymbols, "EodLast")
        fldTickFirst = d4field(tblSymbols, "TickFirst")
        fldTickLast = d4field(tblSymbols, "TickLast")
    
        tagSymbol = d4tag(tblSymbols, "Symbol")
        tagSymbolID = d4tag(tblSymbols, "SymID")
If 0 Then ' all
        tagAccessBySymbol = d4tag(tblSymbols, "Symbol")
        tagAccessByDesc = d4tag(tblSymbols, "Desc")
Else
        tagAccessBySymbol = d4tag(tblSymbols, "AccessSym")
        tagAccessByDesc = d4tag(tblSymbols, "AccessDesc")
End If
    End If
    
    Cb4Optimize

    OpenDb = bSuccess

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cUniverse.OpenDB", eGDRaiseError_Raise

End Function

Public Sub CloseDb()
On Error GoTo ErrSection:

    TblClose tblExchanges
    TblClose tblMarkets
    TblClose tblSymbols

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cUniverse.CloseDB", eGDRaiseError_Raise

End Sub
