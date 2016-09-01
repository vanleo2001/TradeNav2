VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOptionSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOptionSymbol.cls
'' Description: Class to allow conversion of stock option symbols
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/26/2010   DAJ         Added in PFG Future Option capabilities
'' 01/27/2010   DAJ         Removed reference to g.PFG for ETA purposes
'' 01/26/2011   DAJ         Added FromRithmic and stock options in ToRithmic
'' 03/07/2011   DAJ         Added IB calls
'' 04/05/2011   DAJ         Fixed Rithmic Stock Option conversion
'' 05/09/2011   DAJ         Fixes for converting IB Future Options
'' 08/25/2011   DAJ         Added CQG symbol conversion
'' 11/18/2011   DAJ         Fix for converting CQG future option symbols
'' 03/07/2012   DAJ         Fix for converting to PFG stock option symbols
'' 04/04/2012   DAJ         Fix for converting from PFG stock option symbols
'' 10/16/2013   DAJ         Added Open E-Cry conversions
'' 08/22/2014   DAJ         Added E-Trade
'' 09/14/2015   DAJ         Added Tradier
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bIsFutureOption As Boolean          ' Is this a future option symbol?
    strBaseSymbol As String             ' Base symbol: "IBM" or "ES-"
    nYear As Long                       ' Year of the expiration
    nMonth As Long                      ' Month of the expiration
    nDay As Long                        ' Day of the expiration (if applicable)
    bIsPut As Boolean                   ' Is this a Put?
    dStrike As Double                   ' Strike price for the option
    strCode As String                   ' Old two letter code for stock options
    
    strBrokerBase As String             ' Base symbol for the broker symbol
    dBrokerStrike As Double             ' Broker provided strike price
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearMembers
'' Description: Clear the member variables
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearMembers()
On Error GoTo ErrSection:

    m.bIsFutureOption = False
    m.strBaseSymbol = ""
    m.nYear = 0
    m.nMonth = 0
    m.nDay = 0
    m.bIsPut = False
    m.dStrike = 0
    m.strCode = ""
    m.strBrokerBase = ""
    m.dBrokerStrike = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.ClearMembers"

End Sub

Public Property Get IsFutureOption() As Boolean
    IsFutureOption = m.bIsFutureOption
End Property
Public Property Let IsFutureOption(ByVal bIsFutureOption As Boolean)
    m.bIsFutureOption = bIsFutureOption
End Property

Public Property Get BaseSymbol() As String
    BaseSymbol = m.strBaseSymbol
End Property
Public Property Let BaseSymbol(ByVal strBaseSymbol As String)
On Error GoTo ErrSection:
    
    Dim lContract As Long               ' Contract
    
    m.strBaseSymbol = Trim(UCase(strBaseSymbol))
    If IsFuture(m.strBaseSymbol) Then
        m.bIsFutureOption = True
        lContract = Val(Parse(m.strBaseSymbol, "-", 2))
        m.strBaseSymbol = Parse(m.strBaseSymbol, "-", 1)
        If lContract > 200000 Then
            m.nYear = Int(lContract / 100)
            m.nMonth = lContract Mod 100
        End If
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cOptionSymbol.BaseSymbol"

End Property

Public Property Get Year() As Long
    Year = m.nYear
End Property
Public Property Let Year(ByVal nYear As Long)
    If nYear > 0 And nYear < 100 Then nYear = nYear + 2000
    m.nYear = nYear
End Property

Public Property Get Month() As Long
    Month = m.nMonth
End Property
Public Property Let Month(ByVal nMonth As Long)
    m.nMonth = nMonth
End Property

Public Property Get Day() As Long
    Day = m.nDay
End Property
Public Property Let Day(ByVal nDay As Long)
On Error GoTo ErrSection:
    
    If nDay <= 0 Then
        m.nDay = CalcDefaultDay
    Else
        m.nDay = nDay
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cOptionSymbol.Day"

End Property

Public Property Get IsPut() As Boolean
    IsPut = m.bIsPut
End Property
Public Property Let IsPut(ByVal bIsPut As Boolean)
    m.bIsPut = bIsPut
End Property

Public Property Get Strike() As Double
    Strike = m.dStrike
End Property
Public Property Let Strike(ByVal dStrike As Double)
    m.dStrike = dStrike
End Property

Public Property Get BrokerBase() As String
    BrokerBase = m.strBrokerBase
End Property
Public Property Let BrokerBase(ByVal strBrokerBase As String)
    m.strBrokerBase = strBrokerBase
End Property

Public Property Get BrokerStrike() As Double
    BrokerStrike = m.dBrokerStrike
End Property
Public Property Let BrokerStrike(ByVal dBrokerStrike As Double)
    m.dBrokerStrike = dBrokerStrike
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromGenesis
'' Description: Load up the member variables from the Genesis symbol
'' Inputs:      Genesis symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromGenesis(ByVal strGenesisSymbol As String)
On Error GoTo ErrSection:

    Dim nDate&, strStrike$, strDate$, iPos&
    Dim astrSections As cGdArray        ' Sections of the symbol
    
    ClearMembers
    If InStr(strGenesisSymbol, " ") > 0 Then
        Set astrSections = New cGdArray
        ' can be from either "C12.5" or "Call 12.5"
        strGenesisSymbol = UCase(Trim(strGenesisSymbol))
        strGenesisSymbol = Replace(strGenesisSymbol, " CALL ", " C")
        strGenesisSymbol = Replace(strGenesisSymbol, " PUT ", " P")
        astrSections.SplitFields strGenesisSymbol, " "
    ElseIf InStr(strGenesisSymbol, "/") > 0 Then
        Set astrSections = New cGdArray
        astrSections.SplitFields UCase(Trim(strGenesisSymbol)), "/"
    Else
        Exit Sub
    End If
    
    m.strBaseSymbol = astrSections(0)
    If IsFuture(m.strBaseSymbol) Then
        ' Future Option
        m.bIsFutureOption = True
        strStrike = astrSections(1)
        nDate = Val(Parse(m.strBaseSymbol, "-", 2))
        If nDate > 190000 Then
            m.nYear = Int(nDate / 100)
            m.nMonth = nDate Mod 100
        Else
            m.nYear = VBA.Year(Date)
            m.nMonth = VBA.Month(Date)
        End If
        m.strBaseSymbol = Parse(m.strBaseSymbol, "-", 1)
    Else
        ' Stock Option
        m.bIsFutureOption = False
        If astrSections.Size = 2 Then
            ' IBM AD
            m.strCode = astrSections(1)
            If Left(m.strCode, 1) > "L" Then
                m.bIsPut = True
            Else
                m.bIsPut = False
            End If
        ElseIf astrSections.Size = 3 Then
            ' IBM 20100522 C123
            ' or IBM 201005 C123 (if want to just calculate the standard monthly exp day)
            ' or IBM 10May C55 (if had been abbreviated)
            strDate = astrSections(1)
            Do While True ' strip out any dashes
                iPos = InStr(strDate, "-")
                If iPos = 0 Then Exit Do
                strDate = Left(strDate, iPos - 1) & Mid(strDate, iPos + 1)
            Loop
            If Len(strDate) = 6 Or Len(strDate) = 7 Then ' 201005 or 2010May
                strDate = strDate & "00"
            ElseIf Len(strDate) = 5 Then ' 10Aug
                strDate = "20" & strDate & "00"
            End If
            If Len(strDate) = 9 Then ' 2010May22
                nDate = MonthNumber(Mid(strDate, 5, 3))
                strDate = Left(strDate, 4) & Format(nDate, "00") & Right(strDate, 2)
            End If
            If Len(strDate) = 8 Then
                nDate = Val(strDate)
                If nDate >= 20000000 Then
                    m.nYear = Int(nDate / 10000)
                    m.nMonth = Int((nDate Mod 10000) / 100)
                    m.nDay = nDate Mod 100
                    If m.nDay = 0 Then
                        m.nDay = CalcDefaultDay
                    End If
                End If
            End If
            strStrike = astrSections(2)
        End If
    End If

    If Len(strStrike) > 0 Then
        FromStrikeString strStrike
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromGenesis"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToGenesis
'' Description: Put all the pieces together for a Genesis symbol
'' Inputs:      None
'' Returns:     Genesis Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToGenesis() As String
On Error GoTo ErrSection:

    Dim strReturn$, strStrike$
    Dim bAppendStrike As Boolean
    
    If m.bIsFutureOption Then
        ' Future option: ES-201004 P470.5
        strReturn = Parse(m.strBaseSymbol, "-", 1) & "-" & Str(m.nYear * 100 + m.nMonth)
        bAppendStrike = True
    ElseIf Len(m.strCode) > 0 And m.nYear = 0 Then
        ' IBM AD
        strReturn = m.strBaseSymbol & " " & m.strCode
        bAppendStrike = False
    Else
        ' IBM 20100522 C123
        strReturn = m.strBaseSymbol & " " & Str(m.nYear * 10000 + m.nMonth * 100 + m.nDay)
        bAppendStrike = True
    End If
    
    If bAppendStrike Then
        strReturn = strReturn & " " & ToStrikeString(True)
    End If
    
    ToGenesis = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToGenesis"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToDisplay
'' Description: Put all the pieces together for display to the Screen
'' Inputs:      Context
'' Returns:     Genesis Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToDisplay(Optional ByVal nContext& = 0) As String
On Error GoTo ErrSection:

    Dim strReturn$
    
    If nContext = 1 Then
        ' Description (e.g. tooltip)
        If m.bIsFutureOption Then
            ' ES-2010-Aug Call 805
            strReturn = Parse(m.strBaseSymbol, "-", 1) & "-" _
                & Str(m.nYear) & "-" & MonthName(m.nMonth, True, True)
        ElseIf Len(m.strCode) = 0 Then
            ' MSFT 2010-Aug-16 Call 85
            strReturn = m.strBaseSymbol & " " & Str(m.nYear) & "-" _
                & MonthName(m.nMonth, True, True) & "-" & Str(m.nDay)
        End If
        If Len(strReturn) > 0 Then
            If m.bIsPut Then
                strReturn = strReturn & " Put "
            Else
                strReturn = strReturn & " Call "
            End If
            strReturn = strReturn & ToStrikeString(False)
        End If
    ElseIf nContext = 2 Then
        ' for quote board cell display
        If Len(m.strCode) = 0 And m.nYear > 0 And Not m.bIsFutureOption Then
            If m.nDay = CalcDefaultDay Then
                ' MSFT 10Aug C55
                strReturn = m.strBaseSymbol & " " & Right(Str(m.nYear), 2) & _
                    MonthName(m.nMonth, True, True) & " " & ToStrikeString(True)
            End If
        End If
    End If
    
    If Len(strReturn) > 0 Then
        ToDisplay = strReturn
    Else ' if nothing else, just return the Genesis symbol
        ToDisplay = ToGenesis
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToDisplay"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromPfg
'' Description: Extract the pieces out of a PFG symbol
'' Inputs:      Pfg Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromPfg(ByVal strPfgSymbol As String)
On Error GoTo ErrSection:

    Dim strDate As String               ' Date portion from the symbol
    Dim astrSections As cGdArray        ' Sections of the symbol
    Dim strSymbol As String             ' Symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the string
    Dim strFutureSymbol As String       ' Future symbol
    Dim strGenesisBase As String        ' Genesis base symbol
    Dim dStrikeMult As Double           ' Strike price multiplier
    
    ClearMembers
    strPfgSymbol = UCase(Trim(strPfgSymbol))
    If Len(strPfgSymbol) > 4 Then
        If Left(strPfgSymbol, 4) = "EQ-+" Then
            m.bIsFutureOption = False
            
            Set astrSections = New cGdArray
            astrSections.SplitFields Trim(Mid(strPfgSymbol, 5)), "\"
            If astrSections.Size = 1 Then
                ' MSQAH
                m.strBaseSymbol = Left(astrSections(0), Len(astrSections(0)) - 2)
                m.strCode = Right(astrSections(0), 2)
            ElseIf astrSections.Size = 3 Then
                ' MSFT3\16A0\40, BP\5D2\43
                m.strBaseSymbol = astrSections(0)
                FromStrikeString astrSections(2)
                strDate = astrSections(1)
                m.nDay = Val(Left(strDate, Len(strDate) - 2))
                FromMonthCode Mid(strDate, Len(strDate) - 1, 1)
                m.nYear = YearFromDigit(CLng(Val(Right(strDate, 1))))
            End If
        ElseIf UCase(Right(strPfgSymbol, 1)) = "O" Then
            m.bIsFutureOption = True
            
            ' ESF9C990O
            strSymbol = strPfgSymbol
            
            ' Strip off the O on the end...
            strSymbol = Left(strSymbol, Len(strSymbol) - 1)
            
            ' Walk back from end to first non-numeric character...
            For lIndex = Len(strSymbol) - 1 To 0 Step -1
                If IsNumeric(Mid(strSymbol, lIndex, 1)) = False Then
                    lPos = lIndex
                    Exit For
                End If
            Next lIndex
            
            ' Grab the strike price...
            m.dBrokerStrike = Val(Mid(strSymbol, lPos + 1))
            
            ' Determine Put/Call...
            m.bIsPut = (UCase(Mid(strSymbol, lPos, 1)) = "P")
            
            ' Now grab the base symbol...
            strFutureSymbol = Left(strSymbol, lPos - 1)
            m.strBrokerBase = Left(strFutureSymbol, Len(strFutureSymbol) - 2)
            m.nMonth = CodeToMonth(Mid(strFutureSymbol, Len(strFutureSymbol) - 1, 1))
            m.nYear = YearFromDigit(CLng(Val(Right(strFutureSymbol, 1))), m.nMonth)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromPfg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ConvertBrokerInfo
'' Description: Convert the broker info using the given information
'' Inputs:      Genesis Base Symbol, Strike Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ConvertBrokerInfo(ByVal strGenesisBase As String, ByVal dStrikeMult As Double)
On Error GoTo ErrSection:
    
    If m.bIsFutureOption Then
        m.strBaseSymbol = Replace(strGenesisBase, "O:", "") & "-"
    Else
        m.strBaseSymbol = strGenesisBase
    End If
    m.dStrike = m.dBrokerStrike * dStrikeMult

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.ConvertBrokerInfo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToPfg
'' Description: Put all the pieces together for a PFG symbol
'' Inputs:      None
'' Returns:     Pfg Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToPfg(Optional ByVal strPfgBaseSymbol As String = "", Optional ByVal dStrikeMult As Double = 1#) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strPutCall As String            ' Put/Call letter

    If m.bIsFutureOption Then
        ' ESF9C990O
        If Len(strPfgBaseSymbol) > 0 Then
            If m.bIsPut Then
                strPutCall = "P"
            Else
                strPutCall = "C"
            End If
            
            strReturn = strPfgBaseSymbol & MonthToCode(m.nMonth) & Right(Str(m.nYear), 1) & strPutCall & _
                        Format(Int(RoundToSigDigits(m.dStrike / dStrikeMult)), "#") & "O"
        End If
    ElseIf Len(m.strCode) > 0 Then
        strReturn = "EQ-+" & m.strBaseSymbol & m.strCode
    Else ' EQ-+MSFT3\16A0\40
        ' 03/07/2012 DAJ: As per e-mail from Dave Rettger at PFG, we don't want to be formatting
        ' the expiration day with two digits...
        strReturn = "EQ-+" & m.strBaseSymbol & "\" & Str(m.nDay) & ToMonthCode _
                    & Right(Str(m.nYear), 1) & "\" & ToStrikeString(False)
        'strReturn = "EQ-+" & m.strBaseSymbol & "\" & Format(m.nDay, "00") & ToMonthCode _
                    & Right(Str(m.nYear), 1) & "\" & ToStrikeString(False)
    End If
    
    ToPfg = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToPfg"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromRithmic
'' Description: Extract the pieces out of a Rithmic symbol
'' Inputs:      Rithmic Symbol, Rithmic Exchange, Strike Price Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromRithmic(ByVal strRithmicSymbol As String, ByVal strRithmicExchange As String, Optional ByVal dStrikeMult As Double = 1#)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the string
    Dim strContract As String           ' Contract in the string
    Dim strOptPart As String            ' Option part of the string

    Select Case UCase(strRithmicExchange)
        Case "CME", "CBOT", "NYMEX", "COMEX"
            ' ESF1 C0840
            strContract = Parse(strRithmicSymbol, " ", 1)
            strOptPart = Parse(strRithmicSymbol, " ", 2)
            
            m.bIsFutureOption = True
            m.strBrokerBase = Left(strContract, Len(strContract) - 2)
            m.nMonth = CodeToMonth(Mid(strContract, Len(strContract) - 1, 1))
            m.nYear = YearFromDigit(CLng(Val(Right(strContract, 1))), m.nMonth)
            m.bIsPut = (UCase(Left(strOptPart, 1)) = "P")
            m.dBrokerStrike = Val(Mid(strOptPart, 2))
            
        Case "NYBOT"
            ' KC FMG0011 OMCA0000305002011311
            m.bIsFutureOption = True
            m.strBrokerBase = Parse(strRithmicSymbol, " ", 1)
            
            strContract = Parse(strRithmicSymbol, " ", 2)
            m.nMonth = CodeToMonth(Mid(strContract, 3, 1))
            m.nYear = 2000 + CLng(Val(Right(strContract, 2)))
            
            strOptPart = Parse(strRithmicSymbol, " ", 3)
            m.bIsPut = (UCase(Mid(strOptPart, 3, 1)) = "P")
            m.dBrokerStrike = Val(Mid(strOptPart, 5, 8))
            
        
        Case "CBOE"
            ' F16Apr11C11
            m.bIsFutureOption = False
            
            lPos = -1&
            For lIndex = 1 To Len(strRithmicSymbol)
                If IsDigit(strRithmicSymbol, lIndex) = True Then
                    lPos = lIndex
                    Exit For
                End If
            Next lIndex
            
            If (lPos > 0) And (Len(strRithmicSymbol) > lPos + 7) Then
                m.strBaseSymbol = Left(strRithmicSymbol, lPos - 1)
                m.nDay = CLng(Val(Mid(strRithmicSymbol, lPos, 2)))
                m.nMonth = MonthNumber(Mid(strRithmicSymbol, lPos + 2, 3))
                m.nYear = 2000 + CLng(Mid(strRithmicSymbol, lPos + 5, 2))
                m.bIsPut = (UCase(Mid(strRithmicSymbol, lPos + 7, 1)) = "P")
                m.dStrike = Val(Mid(strRithmicSymbol, lPos + 8)) * dStrikeMult
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromRithmic"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToRithmic
'' Description: Put all the pieces together for a Rithmic symbol
'' Inputs:      None
'' Returns:     Rithmic Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToRithmic(ByVal strRithmicBaseSymbol As String, ByVal strRithmicExchange As String, Optional ByVal dStrikeMult As Double = 1#, Optional ByVal strStrikeFormat As String = "#") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strPutCall As String            ' Put/Call letter

    If m.bIsFutureOption Then
        If (Len(strRithmicBaseSymbol) > 0) And (Len(strRithmicExchange) > 0) Then
            Select Case UCase(strRithmicExchange)
                Case "CME", "CBOT", "NYMEX"
                    ' ESH1 P1234
                    If m.bIsPut Then
                        strPutCall = "P"
                    Else
                        strPutCall = "C"
                    End If
                    
                    strReturn = strRithmicBaseSymbol & MonthToCode(m.nMonth) & Right(Str(m.nYear), 1) & " " & _
                                strPutCall & Format(Int(RoundToSigDigits(m.dStrike / dStrikeMult)), strStrikeFormat)
                        
            End Select
        End If
    Else
        ' F16Apr11C11
        If (Len(strRithmicBaseSymbol) > 0) And (Len(strRithmicExchange) > 0) Then
            If m.bIsPut Then
                strPutCall = "P"
            Else
                strPutCall = "C"
            End If
            
            ' DAJ 04/05/2011: Make sure that the month abbreviation is NOT in all caps...
            strReturn = UCase(strRithmicBaseSymbol) & Format(m.nDay, "00") & MonthName(m.nMonth, , True) & _
                        Right(Str(m.nYear), 2) & strPutCall & Format(Int(RoundToSigDigits(m.dStrike / dStrikeMult)), strStrikeFormat)
        End If
    End If
    
    ToRithmic = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToRithmic"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromComstock
'' Description: Extract the pieces out of a Comstock symbol
'' Inputs:      Comstock Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromComstock(ByVal strComstockSymbol As String)
On Error GoTo ErrSection:

    Dim iPos&, strDate$
    Dim astrSections As cGdArray        ' Sections of the symbol
    
    ClearMembers
    strComstockSymbol = UCase(Trim(strComstockSymbol))
    If Len(strComstockSymbol) >= 3 Then
        Set astrSections = New cGdArray
        astrSections.SplitFields strComstockSymbol, "\"
        If astrSections.Size = 1 Then
            ' MSQAH
            m.strBaseSymbol = Left(strComstockSymbol, Len(strComstockSymbol) - 2)
            m.strCode = Right(strComstockSymbol, 2)
        ElseIf astrSections.Size = 3 Then
            ' O:MSFT3\10A16\40.0
            m.strBaseSymbol = astrSections(0)
            iPos = InStr(m.strBaseSymbol, ":")
            If iPos > 0 Then
                m.strBaseSymbol = Trim(Mid(m.strBaseSymbol, iPos + 1))
            End If
            FromStrikeString astrSections(2)
            strDate = astrSections(1)
            m.nYear = Val(Left(strDate, 2)) + 2000
            FromMonthCode Mid(strDate, 3, 1)
            m.nDay = Val(Right(strDate, 2))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromComstock"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToComstock
'' Description: Put all the pieces together for a Comstock symbol
'' Inputs:      None
'' Returns:     Comstock Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToComstock() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If Len(m.strCode) > 0 Then
        strReturn = m.strBaseSymbol & m.strCode
    Else ' O:MSFT3\10A16\40.0
        strReturn = "O:" & m.strBaseSymbol & "\" & Right(Str(m.nYear), 1) & ToMonthCode _
                    & Format(m.nDay, "00") & "\" & ToStrikeString(False, True)
    End If
    
    ToComstock = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToComstock"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromIb
'' Description: Fill in the members from the given Interactive Brokers info
'' Inputs:      IB Symbol, IB Exchange, Future?, Strike Mult
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromIb(ByVal strIbSymbol As String, ByVal strIbExchange As String, ByVal bFuture As Boolean)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Fields in the symbol

    m.bIsFutureOption = bFuture
    Set astrFields = New cGdArray
    astrFields.SplitFields strIbSymbol, " "
    
    If bFuture Then
        Select Case UCase(strIbExchange)
            Case "CME", "GLOBEX", "NYMEX"
                ' ESG1 C1330
                If (astrFields.Size = 2) And (Len(astrFields(0) > 2)) And (Len(astrFields(1) > 1)) Then
                    m.strBrokerBase = Left(astrFields(0), Len(astrFields(0)) - 2)
                    Month = CodeToMonth(Mid(astrFields(0), Len(astrFields(0)) - 1, 1))
                    Year = YearFromDigit(CLng(Val(Right(astrFields(0), 1))), Month)
                    m.bIsPut = (UCase(Left(astrFields(1), 1)) = "P")
                    m.dBrokerStrike = Val(Mid(astrFields(1), 2))
                End If
                
            Case "ECBOT"
                ' C OYM  MAR 11  12250
                If astrFields.Size = 5 Then
                    m.strBrokerBase = astrFields(1)
                    Year = CLng(Val(astrFields(3)))
                    Month = MonthNumber(astrFields(2))
                    m.bIsPut = (UCase(astrFields(0)) = "P")
                    m.dBrokerStrike = Val(astrFields(4))
                End If
                
        End Select
    Else
        ' F     110319P00017000
        If (astrFields.Size = 2) And (Len(astrFields(1)) = 15) Then
            m.strBaseSymbol = astrFields(0)
            m.strBrokerBase = astrFields(0)
            Year = CLng(Val(Left(astrFields(1), 2)))
            Month = CLng(Val(Mid(astrFields(1), 3, 2)))
            Day = CLng(Val(Mid(astrFields(1), 5, 2)))
            m.bIsPut = (UCase(Mid(astrFields(1), 7, 1)) = "P")
            m.dBrokerStrike = Val(Mid(astrFields(1), 8))
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromIb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToIb
'' Description: Put all the pieces together for an Interactive Brokers symbol
'' Inputs:      None
'' Returns:     Interactive Brokers Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToIb(ByVal strIbBaseSymbol As String, ByVal strIbExchange As String, Optional ByVal dStrikeMult As Double = 1#, Optional ByVal strStrikeFormat As String = "#") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If m.bIsFutureOption Then
        Select Case UCase(strIbExchange)
            Case "CME", "GLOBEX", "NYMEX"
                ' ESG1 C1330
                strReturn = strIbBaseSymbol
                strReturn = strReturn & MonthToCode(Month)
                strReturn = strReturn & Right(Format(Year, "0000"), 1)
                If m.bIsPut Then
                    strReturn = strReturn & " P"
                Else
                    strReturn = strReturn & " C"
                End If
                strReturn = strReturn & Str(m.dStrike / dStrikeMult)
                
            Case "ECBOT"
                ' C OYM  MAR 11  12250
                If m.bIsPut Then
                    strReturn = "P "
                Else
                    strReturn = "C "
                End If
                strReturn = strReturn & Pad(strIbBaseSymbol, 5, "L")
                strReturn = strReturn & UCase(MonthName(Month, True, True)) & " "
                strReturn = strReturn & Right(Format(Year, "0000"), 2)
                strReturn = strReturn & Pad(Str(m.dStrike / dStrikeMult), 7, "R")
                
        End Select
    Else
        ' F     110319P00017000
        strReturn = Pad(strIbBaseSymbol, 6, "L")
        strReturn = strReturn & Right(Format(Year, "0000"), 2)
        strReturn = strReturn & Format(Month, "00")
        strReturn = strReturn & Format(Day, "00")
        If m.bIsPut Then
            strReturn = strReturn & "P"
        Else
            strReturn = strReturn & "C"
        End If
        strReturn = strReturn & Format(m.dStrike / dStrikeMult, "00000000")
    End If
    
    ToIb = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToIb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromCqg
'' Description: Fill in the members from the given CQG symbol
'' Inputs:      CQG Symbol, Strike Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromCqg(ByVal strCqgSymbol As String, Optional ByVal dStrikeMult As Double = 1#)
On Error GoTo ErrSection:

    Dim astrSymbol As cGdArray          ' Symbol broken out into fields
    Dim lFirstDigit As Long             ' First digit in the symbol string
    
    m.bIsFutureOption = True
    
    ' {C|P}.US.<BaseSym><M><YY><Strike> (e.g. C.US.QOJ091355)
    Set astrSymbol = New cGdArray
    astrSymbol.SplitFields strCqgSymbol, "."
    
    If astrSymbol.Size = 3 Then
        If (UCase(astrSymbol(0)) = "C") Or (UCase(astrSymbol(0)) = "P") Then
            lFirstDigit = IndexOfLastNonDigit(astrSymbol(2)) + 1
            If lFirstDigit > 0 Then
                m.strBrokerBase = Left(astrSymbol(2), lFirstDigit - 2)
                Month = CodeToMonth(Mid(astrSymbol(2), lFirstDigit - 1, 1))
                Year = YearFromDigit(CLng(Val(Mid(astrSymbol(2), lFirstDigit, 2))))
                m.bIsPut = (UCase(astrSymbol(0)) = "P")
                m.dBrokerStrike = Val(Mid(astrSymbol(2), lFirstDigit + 2))
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromCqg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToCqg
'' Description: Put all the pieces together for a CQG symbol
'' Inputs:      None
'' Returns:     CQG Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToCqg(ByVal strCqgBaseSymbol As String, Optional ByVal dStrikeMult As Double = 1#) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    ' {C|P}.US.<BaseSym><M><YY><Strike> (e.g. C.US.QOJ091355)
    If m.bIsPut Then
        strReturn = "P.US."
    Else
        strReturn = "C.US."
    End If
    strReturn = strReturn & strCqgBaseSymbol
    strReturn = strReturn & MonthToCode(Month)
    strReturn = strReturn & Right(Str(Year), 2)
    strReturn = strReturn & Format(m.dStrike / dStrikeMult, "#")
    
    ToCqg = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToCqg"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromOec
'' Description: Fill in the members from the given Open E-Cry symbol
'' Inputs:      Open E-Cry Symbol, Strike Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromOec(ByVal strOecSymbol As String, Optional ByVal dStrikeMult As Double = 1#)
On Error GoTo ErrSection:

    Dim astrSymbol As cGdArray          ' Symbol broken out into fields
    Dim lFirstDigit As Long             ' First digit in the symbol string
    
    m.bIsFutureOption = True
    
    ' <BaseSym><M><Y> {C|P}<Strike> ( e.g. OESV3 C1710 )
    Set astrSymbol = New cGdArray
    astrSymbol.SplitFields strOecSymbol, " "
    
    If astrSymbol.Size = 2 Then
        m.strBrokerBase = Left(astrSymbol(0), Len(astrSymbol(0)) - 2)
        Month = CodeToMonth(Mid(astrSymbol(0), Len(astrSymbol(0)) - 1, 1))
        Year = YearFromDigit(CLng(Val(Right(astrSymbol(0), 1))))
        m.bIsPut = (UCase(Left(astrSymbol(1), 1)) = "P")
        m.dBrokerStrike = Val(Mid(astrSymbol(1), 2))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromOec"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToOec
'' Description: Put all the pieces together for an Open E-Cry symbol
'' Inputs:      None
'' Returns:     Open E-Cry Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToOec(ByVal strOecBaseSymbol As String, Optional ByVal dStrikeMult As Double = 1#, Optional ByVal strStrikeFormat As String = "#") As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim dNewStrike As Double            ' New strike price

    ' <BaseSym><M><Y> {C|P}<Strike> ( e.g. OESV3 C1710 )
    strReturn = strOecBaseSymbol
    strReturn = strReturn & MonthToCode(Month)
    strReturn = strReturn & Right(Str(Year), 1)
    
    If m.bIsPut Then
        strReturn = strReturn & " P"
    Else
        strReturn = strReturn & " C"
    End If
    
    'strReturn = strReturn & Format(m.dStrike / dStrikeMult, "#")
    
    dNewStrike = m.dStrike / dStrikeMult
    If strStrikeFormat = "#" Then
        strReturn = strReturn & Str(m.dStrike / dStrikeMult)
    Else
        strReturn = strReturn & Format(m.dStrike / dStrikeMult, strStrikeFormat)
    End If
    
    ToOec = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToOec"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromEtrade
'' Description: Fill in the members from the given E-Trade symbol
'' Inputs:      E-Trade Symbol, Strike Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromEtrade(ByVal strEtradeSymbol As String, Optional ByVal dStrikeMult As Double = 1#)
On Error GoTo ErrSection:

    Dim astrSymbol As cGdArray          ' Symbol broken out into fields
    Dim lFirstDigit As Long             ' First digit in the symbol string
    
    m.bIsFutureOption = True
    
    ' <BaseSymbol>|<Put/Call>|<Year>|<Month>|<Day>|<Strike>
    Set astrSymbol = New cGdArray
    astrSymbol.SplitFields strEtradeSymbol, "|"
    
    If astrSymbol.Size = 6 Then
        m.strBrokerBase = astrSymbol(0)
        m.bIsPut = (UCase(astrSymbol(1)) = "PUT")
        Year = CLng(Val(astrSymbol(2)))
        Month = CLng(Val(astrSymbol(3)))
        Day = CLng(Val(astrSymbol(4)))
        m.dBrokerStrike = Val(astrSymbol(5))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromEtrade"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToEtrade
'' Description: Put all the pieces together for an E-Trade symbol
'' Inputs:      None
'' Returns:     E-Trade Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToEtrade(ByVal strEtradeBaseSymbol As String, Optional ByVal dStrikeMult As Double = 1#, Optional ByVal strStrikeFormat As String = "#") As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Return value for the function
    Dim dNewStrike As Double            ' New strike price

    ' <BaseSymbol>|<Put/Call>|<Year>|<Month>|<Day>|<Strike>
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings, 6
    
    astrReturn(0) = strEtradeBaseSymbol
    If m.bIsPut Then
        astrReturn(1) = "PUT"
    Else
        astrReturn(1) = "CALL"
    End If
    astrReturn(2) = Str(Year)
    astrReturn(3) = Str(Month)
    astrReturn(4) = Str(Day)
    
    dNewStrike = m.dStrike / dStrikeMult
    If strStrikeFormat = "#" Then
        astrReturn(5) = Str(m.dStrike / dStrikeMult)
    Else
        astrReturn(5) = Format(m.dStrike / dStrikeMult, strStrikeFormat)
    End If
    
    ToEtrade = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToEtrade"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromTradier
'' Description: Fill in the members from the given Tradier symbol
'' Inputs:      Tradier Symbol, Strike Multiplier
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromTradier(ByVal strTradierSymbol As String, Optional ByVal dStrikeMult As Double = 1#)
On Error GoTo ErrSection:

    Dim lFirstDigit As Long             ' First digit in the symbol string
    
    m.bIsFutureOption = False
    
    ' CSCO150117C00035000
    lFirstDigit = IndexOfFirstDigit(strTradierSymbol)
    If lFirstDigit > -1 Then
        m.strBrokerBase = Left(strTradierSymbol, lFirstDigit - 1)
        Year = CLng(Val(Mid(strTradierSymbol, lFirstDigit, 2)))
        Month = CLng(Val(Mid(strTradierSymbol, lFirstDigit + 2, 2)))
        Day = CLng(Val(Mid(strTradierSymbol, lFirstDigit + 4, 2)))
        m.bIsPut = (UCase(Mid(strTradierSymbol, lFirstDigit + 6, 1)) = "P")
        m.dBrokerStrike = Val(Mid(strTradierSymbol, lFirstDigit + 7))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromTradier"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToTradier
'' Description: Put all the pieces together for a Tradier symbol
'' Inputs:      None
'' Returns:     Tradier Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToTradier(ByVal strTradierBaseSymbol As String, Optional ByVal dStrikeMult As Double = 1#, Optional ByVal strStrikeFormat As String = "00000000") As String
On Error GoTo ErrSection:
    
    Dim strReturn As String             ' Return value for the function
    Dim dNewStrike As Double            ' New strike price

    ' CSCO150117C00035000
    strReturn = strTradierBaseSymbol & Format(Year Mod 100, "00") & Format(Month, "00") & Format(Day, "00")
    If m.bIsPut Then
        strReturn = strReturn & "P"
    Else
        strReturn = strReturn & "C"
    End If
    strReturn = strReturn & Format(m.dStrike / dStrikeMult, strStrikeFormat)
    
    ToTradier = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToTradier"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToMonthCode
'' Description: Put the month and put/call information into a code
'' Inputs:      None
'' Returns:     Pfg Month Code
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToMonthCode() As String
On Error GoTo ErrSection:

    If m.nMonth = 0 Then
        ToMonthCode = Left(m.strCode, 1)
    ElseIf m.bIsPut Then
        ToMonthCode = Chr(m.nMonth + 76)
    Else
        ToMonthCode = Chr(m.nMonth + 64)
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToMonthCode"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromMonthCode
'' Description: Get the month and put/call information from a PFG month code
'' Inputs:      Pfg Month Code
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FromMonthCode(ByVal strMonthCode As String)
On Error GoTo ErrSection:

    Dim nAsc As Long                  ' ASCII code for the month code
    
    nAsc = Asc(UCase(Trim(strMonthCode)))
    If nAsc > 76 Then
        m.nMonth = nAsc - 76
        m.bIsPut = True
    ElseIf nAsc > 64 Then
        m.nMonth = nAsc - 64
        m.bIsPut = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromMonthCode"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromStrikeString
'' Description: Get the strike and put/call information
'' Inputs:      Strike string
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FromStrikeString(ByVal strStrike As String)
On Error GoTo ErrSection:

    Dim iPos&
    
    strStrike = UCase(Trim(strStrike))
    If Len(strStrike) > 0 Then
        iPos = InStr(strStrike, ",")
        If iPos > 0 Then
            Mid(strStrike, iPos, 1) = "."
        End If
        Do While True
            iPos = InStr(strStrike, " ")
            If iPos = 0 Then Exit Do
            strStrike = Left(strStrike, iPos - 1) & Mid(strStrike, iPos + 1)
        Loop
        
        Select Case UCase(Left(strStrike, 1))
        Case "C"
            m.bIsPut = False
            If Left(strStrike, 4) = "CALL" Then
                m.dStrike = Val(Mid(strStrike, 5))
            Else
                m.dStrike = Val(Mid(strStrike, 2))
            End If
        Case "P"
            m.bIsPut = True
            If Left(strStrike, 3) = "PUT" Then
                m.dStrike = Val(Mid(strStrike, 4))
            Else
                m.dStrike = Val(Mid(strStrike, 2))
            End If
        Case Else
            m.dStrike = Val(strStrike)
        End Select
    End If
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOptionSymbol.FromStrikeString"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToStrikeString
'' Description: Format the strike price to put in the symbol
'' Inputs:      Prepend Put/Call?, Force Decimal?
'' Returns:     Strike Price String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToStrikeString(ByVal bPrependCallPut As Boolean, Optional ByVal bForceDecimal As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strStrike$, iPos&

    ' format the strike (but must override regional settings)
    strStrike = Format(m.dStrike, "#0.0#####")
    iPos = InStr(strStrike, ",")
    If iPos > 0 Then
        Mid(strStrike, iPos, 1) = "."
    End If
    If Not bForceDecimal Then
        If Right(strStrike, 2) = ".0" Then
            strStrike = Left(strStrike, Len(strStrike) - 2)
        End If
    End If
        
    If bPrependCallPut Then
        If m.bIsPut Then
            strStrike = "P" & strStrike
        Else
            strStrike = "C" & strStrike
        End If
    End If
    
    ToStrikeString = strStrike
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.ToStrikeString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcDefaultDay
'' Description: Calculate the default day from the month and year
'' Inputs:      None
'' Returns:     Day of month (Sat after 3rd Friday)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalcDefaultDay() As Long
On Error GoTo ErrSection:

    Dim lDate As Long                   ' Date from rule
    Dim lReturn As Long                 ' Return value for the function
    
    If m.nYear > 0 And m.nMonth > 0 Then
        lDate = GetDateFromRule(m.nYear, m.nMonth, "3F") + 1
        lReturn = VBA.Day(lDate)
    Else
        lReturn = 0
    End If
    
    CalcDefaultDay = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.CalcDefaultDay"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    YearFromDigit
'' Description: Get the full year from a single digit
'' Inputs:      Single Digit, Month
'' Returns:     Full Year
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function YearFromDigit(ByVal lDigit As Long, Optional ByVal lMonth As Long = 0&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lLastYear As Long               ' Last year
    Dim dDate As Double                 ' Current date
    Dim lCurrentMonth As Long           ' Current month
    Dim lCurrentYear As Long            ' Current year
    
    dDate = Date
    lCurrentYear = VBA.Year(dDate)
    lLastYear = VBA.Year(dDate) - 1
    lCurrentMonth = VBA.Month(dDate)
    
    lReturn = lDigit + 2000&
    Do While lReturn < lLastYear
        lReturn = lReturn + 10&
    Loop
    
    If (lMonth > 0&) Then
        If lReturn = lLastYear Then
            lReturn = lReturn + 10&
        ElseIf (lReturn = lCurrentYear) And (lMonth < lCurrentMonth) Then
            lReturn = lReturn + 10&
        End If
    End If
    
    YearFromDigit = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.YearFromDigit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsFuture
'' Description: Is the symbol a future symbol?
'' Inputs:      Symbol
'' Returns:     True if Future symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsFuture(ByVal strSymbol As String) As Boolean
On Error GoTo ErrSection:
            
    Dim bReturn As Boolean              ' Return value for the function
            
    bReturn = False
    If (Left(strSymbol, 1) <> "$") And (Left(strSymbol, 1) <> "#") Then
        If InStr(strSymbol, "-") > 0 Then
            bReturn = True
        End If
    End If
    
    IsFuture = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.IsFuture"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexOfFirstDigit
'' Description: Find the index of the first digit
'' Inputs:      String
'' Returns:     Index of the first digit (-1 if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IndexOfFirstDigit(ByVal strString As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    For lIndex = 1 To Len(strString)
        If IsDigit(Mid(strString, lIndex, 1)) = True Then
            lReturn = lIndex
            Exit For
        End If
    Next lIndex
    
    IndexOfFirstDigit = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.IndexOfFirstDigit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IndexOfLastNonDigit
'' Description: Find the index of the last non-digit
'' Inputs:      String
'' Returns:     Index of the last non-digit (-1 if none)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IndexOfLastNonDigit(ByVal strString As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    For lIndex = Len(strString) To 1 Step -1
        If IsDigit(Mid(strString, lIndex, 1)) = False Then
            lReturn = lIndex
            Exit For
        End If
    Next lIndex
    
    IndexOfLastNonDigit = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOptionSymbol.IndexOfLastNonDigit"
    
End Function

