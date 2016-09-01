VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStyler.cls
'' Description: Applies styling to the form using Hexagora controls
'' Author:      Pilgrim Consulting
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
 
Private Const COLOR_CAPTIONTEXT = 9

Option Explicit

Private mColors(0 To 50) As OLE_COLOR
Private mButtonStyle As eCtlButtonStyle

Private mHexagoraWrap As New clsCommonWrapper

'Public Property Let buttonStyle(buttonStyle As eCtlButtonStyle)
'    mButtonStyle = buttonStyle
'End Property
'Public Property Get buttonStyle() As eCtlButtonStyle
'    buttonStyle = mButtonStyle
'End Property

Public Sub SetColor(eStyleType As eStyleColorTypes, Color As OLE_COLOR)
    mColors(eStyleType) = Color
End Sub

Public Function GetColor(eStyleType As eStyleColorTypes) As OLE_COLOR
    GetColor = mColors(eStyleType)
End Function

Public Function StyleForm(frm As Form)
On Error GoTo Trap
    
    'DEPRECATE  - this function is too slow and is no longer needed
    Exit Function
    
    Dim ctl As Control
    Dim ret As Long

    ret = SetSysColors(1, COLOR_CAPTIONTEXT, vbBlack)
    
    'recolor
    frm.BackColor = Me.GetColor(eForm_Background)
    
    For Each ctl In frm
    
        
        'Frame
        If TypeOf ctl Is ctlUniFrameWL Then
            ctl.BackColor = GetColor(eFrame_Background)
            ctl.BorderColor = GetColor(eFrame_Border)
            'Hexagora added a default caption for frames (the frame's name) so it needs to be removed
        
            If ctl.Caption = ctl.Name Or Left$(ctl.Caption, 5) = "Frame" Then
                ctl.Caption = ""
            End If
        End If
        
        'background colors
        If TypeOf ctl Is ctlUniRadioXP Or TypeOf ctl Is ctlUniLabelXP Or TypeOf ctl Is vsElastic Then
            ctl.BackColor = GetColor(eForm_Background)
        End If
        
        'Button
        'button styles are already defaulted
        If TypeOf ctl Is ctlUniButtonImageXP Then
            ctl.BackColor = GetColor(eButton_Background)
            ctl.BorderColor = GetColor(eButton_Border)
        End If
        
        If TypeOf ctl Is VSFlexGrid Then
            ctl.BackColorBkg = GetColor(eGrid_Background)
        End If
        
'        If TypeOf ctl Is ctlUniLabelWL Or TypeOf ctl Is ctlUniFrameWL Or TypeOf ctl Is ctlUniCheckWL Or TypeOf ctl Is ctlUniRadioWL Then
'            eType = tmsTranslateLabel
'        ElseIf TypeOf ctl Is ctlUniButtonImageWL Then
'            eType = tmsTranslateCommandButton
'        End If
'
'        sTemp = Me.GetTranslation(sTemp, "")
'        If Len(sTemp) Then
'            ctl.Caption = sTemp
'            On Error Resume Next
'            ctl.ToolTipText = sTemp
'            On Error GoTo Trap
'        End If
'
'        'recolor
'        If TypeOf ctl Is ctlUniLabelWL Or TypeOf ctl Is ctlUniFrameWL Or TypeOf ctl Is ctlUniCheckWL Or TypeOf ctl Is ctlUniRadioWL Or TypeOf ctl Is ctlUniTabbedXP Or TypeOf ctl Is PictureBox Then
'            ctl.BackColor = vbWindowBackground  ''vbWhite
'        End If
    
    Next ctl
   
    
Quit:
    Exit Function
Trap:
    'LogError Err, Err.Description, Erl, "EnumFormControls"
    Resume Quit
    Resume
End Function

Public Sub SetButtonStyleDefault(buttonStyle As eCtlButtonStyle)
    mHexagoraWrap.DefaultButtonStyle = buttonStyle
End Sub

Private Sub Class_Initialize()
    
    mHexagoraWrap.ActivateLibrary ("BCP78J75X84N82S7AM86P74Q83B85O75H76W72W83X74ZBEQn")

    '''.SetButtonStyleDefault iCtlBtnStyle_Flat

    'default colors
    
    'Form
    SetColor eForm_Background, vbWhite
    
    'Frames
    SetColor eFrame_Background, vbWhite
    SetColor eFrame_Border, 13421772  '#cccccc
    
    'Buttons
    SetColor eButton_Background, 13469470  '#1e87cd blue
    SetColor eButton_Border, 11761177 '#1976b3 darker blue
    SetColor eButton_Text, vbWhite '#ffffff
    
    'Checkboxes (checkmarks)
    SetColor eCheck_Border, vbBlue
        
    'Flexgrid
    SetColor eGrid_Background, 15921906  '#f2f2f2 light grey
    

End Sub

