VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPriceEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPriceEditor.cls
'' Description: Wrapper class for controls to change a price or quantity
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/08/2013   DAJ         Allow for minimum order quantity, minimum lot size on orders
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents gdUpDown As gdScrollBar
Attribute gdUpDown.VB_VarHelpID = -1
Private WithEvents txtPrice As ctlUniTextBoxXP 'RH was TextBox
Attribute txtPrice.VB_VarHelpID = -1

Private Type mPrivate
    Bars As cGdBars
    bAsTradingUnits As Boolean
    dPrice As Double
    dMin As Double
    dMinMove As Double
    bChanged As Boolean
    bShowIfZero As Boolean
End Type
Private m As mPrivate

Public Sub Init(gdSB As gdScrollBar, txt As ctlUniTextBoxXP, Bars As cGdBars, Optional ByVal dPrice# = 0, Optional ByVal dMin# = 0, Optional ByVal dMax# = 999999, Optional ByVal bAutoPlaceScrollbar As Boolean = True, Optional ByVal bShowIfZero As Boolean = False, Optional ByVal dMinMove As Double = 0#)

    Set gdUpDown = gdSB
    Set txtPrice = txt
    Set m.Bars = Bars
    m.bShowIfZero = bShowIfZero
    
    If (Bars Is Nothing) And (dMinMove <> 0#) And (dMinMove <> 1#) Then
        Set m.Bars = New cGdBars
        m.Bars.Prop(eBARS_TickMove) = dMinMove
        m.Bars.Prop(eBARS_MinMoveInTicks) = 1
    End If
    
    m.dMin = dMin
    If Not m.Bars Is Nothing Then
        m.dMinMove = m.Bars.MinMove(Date)
    ElseIf dMinMove = 0# Then
        m.dMinMove = 1
    Else
        m.dMinMove = dMinMove
    End If
    If m.dMinMove = 0 Then m.dMinMove = 0.01
    
    With gdUpDown
        .TabStop = False
        .Move txt.Left + txt.Width, Int((txt.Top + (txt.Height - .Height) / 2) / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        .Min = -Round(dMax / m.dMinMove)
        .Max = -Round(dMin / m.dMinMove)
        If .Min >= .Max Then
            txtPrice.Locked = True
            gdSB.Visible = False
        Else
            txtPrice.Locked = False
            gdSB.Visible = True
        End If
        If bAutoPlaceScrollbar Then
            .Move txt.Left + txt.Width, Int((txt.Top + (txt.Height - .Height) / 2) / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        End If
    End With
    
    FixPrice dPrice

End Sub

Public Sub ChangeMinMove(ByVal dNewMinMove As Double)

    m.dMinMove = dNewMinMove
    If m.dMinMove = 0# Then m.dMinMove = 0.01
    
    If m.dMinMove = 1 Then
        Set m.Bars = Nothing
        m.dMinMove = 1#
    Else
        Set m.Bars = New cGdBars
        m.Bars.Prop(eBARS_TickMove) = m.dMinMove
        m.Bars.Prop(eBARS_MinMoveInTicks) = 1
    End If
    
    FixPrice txtPrice
    
End Sub

Private Sub Class_Initialize()

    m.bAsTradingUnits = True

End Sub

Private Sub Class_Terminate()
    
    Set gdUpDown = Nothing
    Set txtPrice = Nothing
    Set m.Bars = Nothing
    
End Sub

Private Sub gdUpDown_Change()
    
    FixPrice -gdUpDown.Value * m.dMinMove
    
    If Screen.ActiveControl Is gdUpDown Then
        MoveFocus txtPrice
    End If

End Sub

Private Sub txtPrice_Change()

    If m.Bars Is Nothing Then
        FixPrice txtPrice
    Else
        m.bChanged = True
    End If

End Sub

Private Sub txtPrice_GotFocus()
    
    m.bChanged = False

End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 Then
        If KeyCode = vbKeyUp Then
            KeyCode = 0
            If gdUpDown.Value > gdUpDown.Min And Not txtPrice.Locked Then
                gdUpDown.Value = gdUpDown.Value - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            KeyCode = 0
            If gdUpDown.Value < gdUpDown.Max And Not txtPrice.Locked Then
                gdUpDown.Value = gdUpDown.Value + 1
            End If
        End If
    End If

End Sub

Private Sub txtPrice_LostFocus()

    If m.bChanged Then
        FixPrice txtPrice
        m.bChanged = False
    End If

End Sub

Private Sub FixPrice(Optional vNewPrice)

    Dim dValue#, strText$
    Static bInProgress As Boolean

    If m.dMinMove = 0 Then
        txtPrice = ""
        Exit Sub
    End If
    
    ' exit if called recursively from changing the gdUpDown value
    If bInProgress Then Exit Sub
    bInProgress = True

    ' get new price if passed
    If Not IsMissing(vNewPrice) Then
        If VarType(vNewPrice) = vbString Then
            If Len(Trim(vNewPrice)) = 0 Then
                'txtPrice = ""
                'bInProgress = False
                'Exit Sub
            End If
        End If
        If VarType(vNewPrice) = vbString And Not m.Bars Is Nothing Then
            m.dPrice = m.Bars.PriceFromString(vNewPrice)
        Else
            m.dPrice = Val(vNewPrice)
        End If
    End If
    
    ' set value of scrollbar, which automatically takes
    ' Min and Max into account as well as rounding to min move
    dValue = -Round(m.dPrice / m.dMinMove)
    If gdUpDown.Value <> dValue Then
        gdUpDown.Value = dValue
    End If
    m.dPrice = -gdUpDown.Value * m.dMinMove
    
    ' display price
    If m.dPrice = -999999 Or (m.dPrice = 0 And gdUpDown.Max = 0 And Not m.bShowIfZero) Then
        strText = ""
    ElseIf m.Bars Is Nothing Then
        If m.dMinMove = 0.25 Then
            strText = Format(m.dPrice, "0.00")
        ElseIf m.dMinMove = 0.5 Then
            strText = Format(m.dPrice, "0.0")
        Else
            strText = Str(m.dPrice)
        End If
    Else
        strText = m.Bars.PriceDisplay(m.dPrice, m.bAsTradingUnits)
    End If
    If strText <> txtPrice Then
        txtPrice = strText
        txtPrice.SelStart = Len(txtPrice)
    End If
    
    bInProgress = False

End Sub

' returns price as a double
Public Property Get Price() As Variant
Attribute Price.VB_UserMemId = 0
    Price = m.dPrice
End Property

' sets price from a number or string
Public Property Let Price(ByVal vNewPrice As Variant)
    FixPrice vNewPrice
End Property

Public Property Get AsTradingUnits() As Boolean
    AsTradingUnits = m.bAsTradingUnits
End Property

Public Property Let AsTradingUnits(ByVal vNewValue As Boolean)
    m.bAsTradingUnits = vNewValue
    FixPrice
End Property

Public Property Get Min() As Double
    Min = m.dMin
End Property
