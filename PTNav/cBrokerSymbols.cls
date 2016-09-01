VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrokerSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cBrokerSymbols.cls
'' Description: Object that handles a collection of broker symbol information
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/13/2014   DAJ         Fixed check for Genesis symbols when exchanges given
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    BrokerSymbols As cGdTree            ' Collection of broker symbol information
End Type
Private m As mPrivate

Public Property Get Item(ByVal vIndexOrKey As Variant) As Object
Attribute Item.VB_UserMemId = 0
    If m.BrokerSymbols.Exists(vIndexOrKey) Then
        Set Item = m.BrokerSymbols(vIndexOrKey)
    Else
        Set Item = Nothing
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add the given broker symbol object to the collection
'' Inputs:      Broker Symbol Object
'' Returns:     Index
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(ByVal BrokerSym As Object) As Long
On Error GoTo ErrSection:

    Add = m.BrokerSymbols.Add(BrokerSym, BrokerSym.GenesisBase)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbols.Add"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Count
'' Description: Return the number of objects that are in the collecion
'' Inputs:      None
'' Returns:     Number of objects in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Count() As Long
On Error GoTo ErrSection:

    Count = m.BrokerSymbols.Count

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbols.Count"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateItem
'' Description: Update the given item in the collection
'' Inputs:      Broker Symbol Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateItem(ByVal BrokerSym As Object)
On Error GoTo ErrSection:

    If m.BrokerSymbols.Exists(BrokerSym.GenesisBase) Then
        Set m.BrokerSymbols(BrokerSym.GenesisBase) = BrokerSym
    Else
        Add BrokerSym
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbols.UdpateItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetByBrokerSymbol
'' Description: Get the broker symbol object for the given broker base symbol
'' Inputs:      Broker Base Symbol, Broker Exchange
'' Returns:     Broker Symbol Object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetByBrokerSymbol(ByVal strBrokerBase As String, Optional ByVal strBrokerExchange As String = "") As Object
On Error GoTo ErrSection:

    Dim Symbol As Object                ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Symbol = Nothing
    For lIndex = 1 To m.BrokerSymbols.Count
        If Item(lIndex).BrokerBase = strBrokerBase Then
            If (Len(strBrokerExchange) = 0) Or (Item(lIndex).BrokerExchange = strBrokerExchange) Then
                Set Symbol = Item(lIndex)
                Exit For
            End If
        End If
    Next lIndex
    
    Set GetByBrokerSymbol = Symbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbols.GetByBrokerSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetByGenesisSymbol
'' Description: Get the broker symbol object for the given Genesis base symbol
'' Inputs:      Genesis Base Symbol
'' Returns:     Broker Symbol Object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetByGenesisSymbol(ByVal strGenesisBase As String) As Object
On Error GoTo ErrSection:

    Dim Symbol As Object                ' Broker symbol object
    
    Set Symbol = Nothing
    If m.BrokerSymbols.Exists(strGenesisBase) Then
        Set Symbol = m.BrokerSymbols(strGenesisBase).MakeCopy
    End If
    
    Set GetByGenesisSymbol = Symbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbols.GetByGenesisSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetByBrokerPositionSymbol
'' Description: Get the broker symbol object for the given broker position base symbol
'' Inputs:      Broker Position Base Symbol, Broker Exchange
'' Returns:     Broker Symbol Object (Nothing if not found)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetByBrokerPositionSymbol(ByVal strBrokerPositionBase As String, Optional ByVal strBrokerExchange As String = "") As Object
On Error GoTo ErrSection:

    Dim Symbol As Object                ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Symbol = Nothing
    For lIndex = 1 To m.BrokerSymbols.Count
        If Item(lIndex).BrokerPositionBase = strBrokerPositionBase Then
            If (Len(strBrokerExchange) = 0) Or (Item(lIndex).BrokerExchange = strBrokerExchange) Then
                Set Symbol = Item(lIndex)
                Exit For
            End If
        End If
    Next lIndex
    
    Set GetByBrokerPositionSymbol = Symbol

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbols.GetByBrokerPositionSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFile
'' Description: Read in the broker symbols from the given file
'' Inputs:      Broker Symbols File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFile(ByVal strFile As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File read into an array
    Dim BrokerSym As cBrokerSymbol      ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strFile) Then
        For lIndex = 0 To astrFile.Size - 1
            Set BrokerSym = New cBrokerSymbol
            BrokerSym.FromString astrFile(lIndex)
            Add BrokerSym
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbols.FromFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromIbFile
'' Description: Read in the Interactive Brokers symbols from the given file
'' Inputs:      Broker Symbols File
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromIbFile(ByVal strFile As String)
On Error GoTo ErrSection:

    Dim astrFile As cGdArray            ' File read into an array
    Dim BrokerSym As cIbSymbol          ' Broker symbol object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrFile = New cGdArray
    If astrFile.FromFile(strFile) Then
        For lIndex = 0 To astrFile.Size - 1
            Set BrokerSym = New cIbSymbol
            BrokerSym.FromString astrFile(lIndex)
            Add BrokerSym
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbols.FromIbFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenesisSymbolList
'' Description: Return a list of the Genesis symbols that are in the file
'' Inputs:      Only if stream?, Exchange List
'' Returns:     List of symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GenesisSymbolList(Optional ByVal bOnlyIfStream As Boolean = True, Optional ByVal strOnlyTheseExchanges As String = "") As cGdArray
On Error GoTo ErrSection:

    Dim astrSymbolList As cGdArray      ' List of symbols to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim bAdd As Boolean                 ' Add the symbol?
    
    Set astrSymbolList = New cGdArray
    astrSymbolList.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.BrokerSymbols.Count
        bAdd = False
        
        If (bOnlyIfStream = False) Then
            bAdd = True
        ElseIf Len(strOnlyTheseExchanges) > 0 Then
            bAdd = (InStr("," & strOnlyTheseExchanges & ",", "," & Item(lIndex).BrokerExchange & ",") > 0)
        ElseIf (Item(lIndex).RtOverride = True) Then
            bAdd = True
        End If
        
        If bAdd = True Then
            astrSymbolList.Add Item(lIndex).GenesisBase
        End If
    Next lIndex
    
    Set GenesisSymbolList = astrSymbolList

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cBrokerSymbol.GenesisSymbolList"
    
End Function

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

    Set m.BrokerSymbols = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbols.Class_Initialize"
    
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

    Set m.BrokerSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cBrokerSymbols.Class_Terminate"
    
End Sub
