VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEquityFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cEquityFilter.cls
'' Description: Wrapper class for equity filter options
''
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lMovingAveragePeriod As Long        ' Period for the equity moving average
    strMovingAverageType As String      ' Type of moving average (Simple, Exponential)
    bEquityFilterOn As Boolean          ' Is the equity filter turned on?
    nMode As eGDEquityFilterMode        ' Mode of the equity filter
End Type
Private m As mPrivate

Public Property Get MovingAveragePeriod() As Long
    MovingAveragePeriod = m.lMovingAveragePeriod
End Property
Public Property Let MovingAveragePeriod(ByVal lMovingAveragePeriod As Long)
    m.lMovingAveragePeriod = lMovingAveragePeriod
End Property

Public Property Get MovingAverageType() As String
    MovingAverageType = m.strMovingAverageType
End Property
Public Property Let MovingAverageType(ByVal strMovingAverageType As String)
    m.strMovingAverageType = strMovingAverageType
End Property

Public Property Get EquityFilterOn() As Boolean
    EquityFilterOn = m.bEquityFilterOn
End Property
Public Property Let EquityFilterOn(ByVal bEquityFilterOn As Boolean)
    m.bEquityFilterOn = bEquityFilterOn
End Property

Public Property Get EquityFilterMode() As eGDEquityFilterMode
    EquityFilterMode = m.nMode
End Property
Public Property Let EquityFilterMode(ByVal nEquityFilterMode As eGDEquityFilterMode)
    m.nMode = nEquityFilterMode
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnglishString
'' Description: Generate an english string based on the values
'' Inputs:      Take Next Trade?, Description Text or Action Text?
'' Returns:     English String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnglishString(ByVal nTakeNextTrade As eGDTakeNextTradeValue, Optional ByVal bDescription As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If bDescription Then
        If m.bEquityFilterOn = False Then
            strReturn = "Equity Filter OFF: All entries taken"
        Else
            If m.nMode = eGDEquityFilterMode_BelowMa Then
                strReturn = "Equity Filter ON: Entries ignored if equity is below " & Str(m.lMovingAveragePeriod) & _
                            " trade " & LCase(m.strMovingAverageType) & " moving average"
            Else
                strReturn = "Equity Filter ON: Entries ignored if " & Str(m.lMovingAveragePeriod) & " trade " & _
                            LCase(m.strMovingAverageType) & " moving average is falling"
            End If
        End If
    Else
        If m.bEquityFilterOn = False Then nTakeNextTrade = eGDTakeNextTrade_NoEquityFilter
            
        Select Case nTakeNextTrade
            Case eGDTakeNextTrade_NoEquityFilter
                strReturn = "Take the next trade: No equity filter applied"
                
            Case eGDTakeNextTrade_NotEnoughData
                strReturn = "Not enough data to calculate equity filter"
                
            Case eGDTakeNextTrade_Yes
                If m.nMode = eGDEquityFilterMode_BelowMa Then
                    strReturn = "Take the next trade: Equity is not below " & Str(m.lMovingAveragePeriod) & _
                                " trade " & LCase(m.strMovingAverageType) & " moving average"
                Else
                    strReturn = "Take the next trade: " & Str(m.lMovingAveragePeriod) & " trade " & _
                                LCase(m.strMovingAverageType) & " moving average is not falling"
                End If
                
            Case eGDTakeNextTrade_No
                If m.nMode = eGDEquityFilterMode_BelowMa Then
                    strReturn = "Do not the take next trade: Equity is below " & Str(m.lMovingAveragePeriod) & _
                                " trade " & LCase(m.strMovingAverageType) & " moving average"
                Else
                    strReturn = "Do not the take next trade: " & Str(m.lMovingAveragePeriod) & " trade " & _
                                LCase(m.strMovingAverageType) & " moving average is falling"
                End If
            
        End Select
    End If
    
    EnglishString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cEquityFilter.EnglishString", , g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RptPrintString
'' Description: Generate an english print string for performance report print
'' Inputs:      None
'' Returns:     English String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RptPrintString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If m.bEquityFilterOn = False Then
        strReturn = "Equity Filter:|OFF - All entries taken"
    Else
        If m.nMode = eGDEquityFilterMode_BelowMa Then
            strReturn = "Equity Filter:|ON - Entries ignored if equity is below " & Str(m.lMovingAveragePeriod) & _
                        " trade " & LCase(m.strMovingAverageType) & " moving average"
        Else
            strReturn = "Equity Filter:|ON - Entries ignored if " & Str(m.lMovingAveragePeriod) & " trade " & _
                        LCase(m.strMovingAverageType) & " moving average is falling"
        End If
    End If
    
    RptPrintString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cEquityFilter.RptPrintString", , g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Set the class variables from the delimited string passed in
'' Inputs:      Delimited String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strSettings As String)
On Error GoTo ErrSection:

    Dim astrSettings As cGdArray        ' Array of values split out from delimited string

    Set astrSettings = New cGdArray
    astrSettings.Create eGDARRAY_Strings
    astrSettings.SplitFields strSettings, ";"
    
    m.strMovingAverageType = astrSettings(0)
    m.lMovingAveragePeriod = CLng(Val(astrSettings(1)))
    m.nMode = CLng(Val(astrSettings(2)))
    If astrSettings(3) = "1" Then
        m.bEquityFilterOn = True
    Else
        m.bEquityFilterOn = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEquityFilter.FromString", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create a delimited string from the class variables
'' Inputs:      None
'' Returns:     Delimited String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrReturn As cGdArray          ' Array of values to join together to return
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    astrReturn(0) = m.strMovingAverageType
    astrReturn(1) = Str(m.lMovingAveragePeriod)
    astrReturn(2) = Str(m.nMode)
    If m.bEquityFilterOn Then
        astrReturn(3) = "1"
    Else
        astrReturn(3) = "0"
    End If
    
    ToString = astrReturn.JoinFields(";")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cEquityFilter.ToString", , g.strAppPath
    
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

    m.strMovingAverageType = "None"
    m.lMovingAveragePeriod = 0&
    m.nMode = eGDEquityFilterMode_BelowMa
    m.bEquityFilterOn = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEquityFilter.Class_Initialize", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cEquityFilter.Class_Terminate", , g.strAppPath
    
End Sub
