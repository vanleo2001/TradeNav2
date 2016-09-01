VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.16#0"; "gdOCX.ocx"
Begin VB.Form frmSingleSymHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Download"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
Begin HexUniControls.ctlUniFrameWL fraExistingData
VistaStyle      =   0   'False
      Caption         =   "Data you currently have"
      Height          =   1335
      Left            =   105
      TabIndex        =   8
      Top             =   780
      Width           =   2805
      Begin gdOCX.gdSelectDate gdDateCurrFrom 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Enabled         =   0   'False
         AllowWeekends   =   0   'False
         MaxDate         =   40522
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin gdOCX.gdSelectDate gdDateCurrTo 
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Enabled         =   0   'False
         AllowWeekends   =   0   'False
         MaxDate         =   40522
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
Begin HexUniControls.ctlUniLabelXP Label1
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   405
         Width           =   495
      End
Begin HexUniControls.ctlUniLabelXP Label2
         Caption         =   "To"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   495
      End
   End
Begin HexUniControls.ctlUniFrameWL fraDataToGet
VistaStyle      =   0   'False
      Caption         =   "Data to get"
      Height          =   1335
      Left            =   3135
      TabIndex        =   3
      Top             =   780
      Width           =   2805
      Begin gdOCX.gdSelectDate gdDateGetFrom 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   40522
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
      Begin gdOCX.gdSelectDate gdDateGetTo 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         AllowWeekends   =   0   'False
         MaxDate         =   40522
         MaxDateIsToday  =   -1  'True
         Value           =   37015
      End
Begin HexUniControls.ctlUniLabelXP Label3
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   495
      End
Begin HexUniControls.ctlUniLabelXP Label4
         Caption         =   "To"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   495
      End
   End
Begin HexUniControls.ctlUniButtonImageXP cmdCancel
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2348
      Width           =   1095
   End
Begin HexUniControls.ctlUniButtonImageXP cmdDownload
      Caption         =   "Download"
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   2348
      Width           =   1095
   End
Begin HexUniControls.ctlUniLabelXP lblHistoryInfo
      Caption         =   "To get additional history data ..."
      Height          =   495
      Left            =   105
      TabIndex        =   0
      Top             =   128
      Width           =   5805
   End
End
Attribute VB_Name = "frmSingleSymHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kDownloadStr = "@From-To;Period;I;SecurityType;"
Private Const kMaxRangeIntraday = 60
Private Const kInfoCaption = "To get additional data adjust the dates for 'Data to get', if desired, then click download when ready. You can only download up to 60 days of data at a time."

Private Type mPrivate
    Bars As cGdBars
    aDownloadStrings As New cGdArray
    
    'start & end data user already have
    dDateDataStart As Double
    dDateDataEnd As Double
    'start & end data to get
    dDateGetStart As Double
    dDateGetEnd As Double
    
End Type

Private m As mPrivate

Public Sub ShowMe(Chart As cChart)
On Error GoTo ErrSection:
        
    If Chart Is Nothing Then
        Unload Me
        Exit Sub
    End If
            
    Set m.Bars = Chart.Bars
    If m.Bars.IsIntraday Then
        If InitIntradayDate Then
            lblHistoryInfo.Caption = kInfoCaption
            CenterTheForm Me
            ShowForm Me, eForm_Modal
        Else
            Unload Me
            Exit Sub
        End If
    Else
        If InitDailyDate Then
            BuildDownloadStrings
            DownloadData
        End If
        Unload Me
        Exit Sub
    End If
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSingleSymHistory.ShowMe", eGDRaiseError_Raise

End Sub

Private Function InitIntradayDate() As Boolean
On Error GoTo ErrSection:
'Design note:
'----------------------------------------------------------------------------
'    history  FD                    LD  daily update EOD       now
'  xxxxxxxxxxx|----------------------|eeeeeeeeeeeeeeeeeeexxxxxxx|
' xxxxx = data user does not have
' eeeee = data user has for end of day, but not tick
' |---| = data user does have
'    FD = first day of data
'    LD = last day of data
' data prior to FD can be gotten using this form
' data after daily EOD update should be gotten via daily download
'----------------------------------------------------------------------------
'If last day of tick < last day of EOD then (catch up)
'   Get From: (default) = last day of tick + 1
'             (min) = first day of tick
'             (max) = last day of tick + 1
'   Get To: (default) = last daily download or max range
'           (min) = "Get From:" date
'           (max) = last daily download or max range
'else (get history)
'   Get From: (default) = first day of tick - max range
'             (min) = same as default
'             (max) = "Get To:" date
'   Get To: (default) = first day of tick - 1
'           (min) = same as default
'           (max) = last day of tick or max range
'endif

    Dim nSymId&, dDateEndEOD#
    Dim dLastDailyDownload As Double
    Dim bOkay As Boolean

    nSymId = m.Bars.Prop(eBARS_SymbolID)
    dDateEndEOD = g.SymbolPool.EodLastDate(nSymId)
    m.dDateDataStart = g.SymbolPool.TickFirstDate(nSymId)
    m.dDateDataEnd = g.SymbolPool.TickLastDate(nSymId)
    bOkay = True
    
    dLastDailyDownload = LastDailyDownload
    
    If m.dDateDataStart = 0 Or m.dDateDataEnd = 0 Then
        MsgBox "There is no intraday data for " & m.Bars.Prop(eBARS_Symbol) & "."
        Exit Function
    ElseIf dDateEndEOD <> dLastDailyDownload Then       'LastDailyDownload() Then
        MsgBox "Your data is out of sync. Do a daily update then try again."
        Exit Function
    End If

    If m.dDateDataEnd < dDateEndEOD Then
        'user have been doing daily download for EOD but not ticks
        'catch the tick data up to the last daily EOD download
        m.dDateGetStart = m.dDateDataEnd + 1
        m.dDateGetEnd = dDateEndEOD
        If Abs(m.dDateGetEnd - m.dDateGetStart) > kMaxRangeIntraday Then
            m.dDateGetEnd = m.dDateGetStart + kMaxRangeIntraday
        End If
    ElseIf m.dDateDataEnd = dDateEndEOD Or (m.dDateDataEnd - dDateEndEOD) = 1 Then
        'user's tick data is up-to-date with EOD data - get history
        'allow end date for tick data to be 1 day ahead of EOD data
        'to account for tick data coming in from real time
        m.dDateGetStart = m.dDateDataStart - kMaxRangeIntraday
        m.dDateGetEnd = m.dDateDataStart - 1
    ElseIf Weekday(m.dDateDataEnd) = vbMonday And (m.dDateDataEnd - dDateEndEOD) = 3 Then
        'If latest tick data is gotten via realtime streaming or quoteboard refresh
        'on a MONDAY then latest tick data would be THREE days ahead of EOD data.
        m.dDateGetStart = m.dDateDataStart - kMaxRangeIntraday
        m.dDateGetEnd = m.dDateDataStart - 1
    Else
        MsgBox "Your data is out of sync. Do a daily update then try again."
        bOkay = False
    End If
    
    If bOkay Then
        'controls for data user currently have
        gdDateCurrFrom = m.dDateDataStart
        gdDateCurrTo = m.dDateDataEnd
        'controls for data to get
        gdDateGetFrom = m.dDateGetStart
        gdDateGetTo = m.dDateGetEnd
    End If
    
    InitIntradayDate = bOkay
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSingleSymHistory.InitIntradayDate", eGDRaiseError_Raise
    
End Function

Private Function InitDailyDate() As Boolean
On Error GoTo ErrSection:
'Design note:
'----------------------------------------------------------------------------
'   history  FD                    LD      now
' xxxxxxxxxxx|----------------------|xxxxxxx|
' xxxxx = data user does not have
' |---| = data user does have
'    FD = first day of data
'    LD = last day of data
' data prior to FD and up to LD can be gotten using this form
' data after LD should be gotten via daily download
'----------------------------------------------------------------------------
'This form will only get history data for EOD.

    If InStr(m.Bars.Prop(eBARS_Symbol), "$-") Then
        MsgBox "This is a sector or subsector symbol. Additional data not available."
        Exit Function
    End If
    
    m.dDateDataStart = g.SymbolPool.EodFirstDate(m.Bars.Prop(eBARS_SymbolID))
    m.dDateDataEnd = g.SymbolPool.EodLastDate(m.Bars.Prop(eBARS_SymbolID))
        
    m.dDateGetStart = gdDateGetFrom.MinDate     'this is 01/01/1900
    m.dDateGetEnd = m.dDateDataEnd              'in case current data is corrupted

    'controls showing data user currently have
    gdDateCurrFrom = m.dDateDataStart
    gdDateCurrTo = m.dDateDataEnd
    'controls showing data to get
    gdDateGetFrom = m.dDateGetStart
    gdDateGetTo = m.dDateGetEnd
    gdDateGetFrom.Enabled = False
    
    InitDailyDate = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSingleSymHistory.InitDailyDate", eGDRaiseError_Raise
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDownload_Click()
On Error Resume Next

    If BuildDownloadStrings Then
        Me.Hide
        DownloadData
        Unload Me
    End If
    
End Sub

Private Function BuildDownloadStrings() As Boolean
On Error GoTo ErrSection:

    Dim strDownload$, strSecurityType$, strSym$
    Dim strFrom$, strTo$
    Dim bOkay As Boolean

    m.aDownloadStrings.Size = 0
    bOkay = True
        
    m.dDateGetStart = gdDateGetFrom
    m.dDateGetEnd = gdDateGetTo
    
    strFrom = Format(m.dDateGetStart, "YYYYMMDD")
    strTo = Format(m.dDateGetEnd, "YYYYMMDD")
    strSecurityType = SecurityType(m.Bars)
    strSym = m.Bars.Prop(eBARS_Symbol)
    
    'downlad EOD history no matter what
    strDownload = kDownloadStr
    strDownload = Replace(strDownload, "From", strFrom)
    strDownload = Replace(strDownload, "To", strTo)
    strDownload = Replace(strDownload, "Period", "E")
    strDownload = Replace(strDownload, "SecurityType", strSecurityType)
    strDownload = strDownload & strSym
    
    m.aDownloadStrings.Add strDownload
    
    If m.Bars.IsIntraday Then
        bOkay = IsValidIntradayGet()
        If bOkay Then
            strDownload = kDownloadStr
            strDownload = Replace(strDownload, "From", strFrom)
            strDownload = Replace(strDownload, "To", strTo)
            strDownload = Replace(strDownload, "Period", "T")
            strDownload = Replace(strDownload, "SecurityType", strSecurityType)
            strDownload = strDownload & strSym
            m.aDownloadStrings.Add strDownload
        End If
    End If
    
    BuildDownloadStrings = bOkay
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSingleSymHistory.BuildDownloadStrings", eGDRaiseError_Raise

End Function

Private Function DownloadData() As Boolean
On Error GoTo ErrSection:
           
           
    If ProcessIsBusy(True) Or m.aDownloadStrings.Size = 0 Then Exit Function
    
    ' initialize the status form
    frmStatus.IsBusy = True
    frmStatus.Status = eStatus_Initialized
    frmStatus.ShowDetails False
    frmStatus.SetTitle "Retrieving History"
    frmStatus.AddDetail "Downloading History"

    ' Make the ftp request
    Set MsgForm = frmStatus
    If FtpRequest(m.aDownloadStrings) = False Then
        MsgBox "FtpRequest failed."
        If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
        GoTo ErrExit
    End If
    
    ' Distribute data from REX
    If frmStatus.Status = eStatus_Completed Then
                        
        ' Distribute the data (if data to distribute)
        If Not DistributeData("Distributing Data", False, ",4") Then
            frmStatus.Status = eStatus_Error
            frmStatus.AddDetail "ERROR downloading data"
        End If
    
        If frmStatus.Status = eStatus_Completed Then
            frmStatus.AddDetail "Final Updating"
            DM_DistribData ""
            frmStatus.Status = eStatus_Running
        
            ' refresh data on all forms
            frmStatus.AddDetail "Reloading Data"
            frmStatus.UpdateProgress "Reloading Data"
            g.RealTime.RefreshAllFormData True                  '6054
            frmStatus.AddDetail "Finished"
            
            frmStatus.Status = eStatus_Completed
            DownloadData = True
        End If
    End If
    
    
ErrExit:
    Set MsgForm = Nothing
    frmStatus.IsBusy = False
    Exit Function
    
ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    frmStatus.IsBusy = False
    RaiseError "frmSingleSymHistory.DownloadData", eGDRaiseError_Raise
    Resume ErrExit

End Function

Private Function IsValidIntradayGet() As Boolean
On Error GoTo ErrSection

    Dim dNewGetStart#, dNewGetEnd#
    Dim dDateEndEOD#
    Dim dDateMin#, dDateMax#
    Dim bValid As Boolean
    
    dNewGetStart = gdDateGetFrom
    dNewGetEnd = gdDateGetTo
    
    If dNewGetEnd < dNewGetStart Then
        MsgBox "The 'To: date' cannot be earlier than the 'From: date'."
        Exit Function
    End If
    
    dDateEndEOD = g.SymbolPool.EodLastDate(m.Bars.Prop(eBARS_SymbolID))
    
    If dDateEndEOD <> LastDailyDownload() Then
        MsgBox "Your data is out of sync. Do a daily update then try again."
        Exit Function
    End If
    
    'see design note for min and max date info
    If m.dDateDataEnd < dDateEndEOD Then
        dDateMin = m.dDateDataStart
        dDateMax = m.dDateDataEnd + 1
        If dNewGetStart >= dDateMin Then
            If dNewGetStart <= dDateMax Then
                dDateMin = dNewGetStart
                dDateMax = dDateEndEOD
                If dNewGetEnd >= dDateMin Then
                    If dNewGetEnd <= dDateMax Then
                        If Abs(dNewGetEnd - dNewGetStart) <= kMaxRangeIntraday Then
                            bValid = True
                        Else
                            MsgBox "You can only download " & Str(kMaxRangeIntraday) & " days of data at a time."
                        End If
                    Else
                        MsgBox "The 'To: date' cannot be later than " & DateFormat(dDateMax) & "."
                    End If
                Else
                    MsgBox "The 'To: date' cannot be earlier than " & DateFormat(dNewGetStart) & "."
                End If
            Else
                MsgBox "The 'From: date' cannot be later than " & DateFormat(dDateMax) & "."
            End If
        Else
            MsgBox "The 'From: date' cannot be earlier than " & DateFormat(dDateMin) & "."
        End If
    ElseIf m.dDateDataEnd >= dDateEndEOD Then
        dDateMin = m.dDateDataStart - 1
        dDateMax = dDateEndEOD
        If dNewGetEnd >= dDateMin Then
            If dNewGetEnd <= dDateMax Then
                dDateMin = m.dDateDataStart - kMaxRangeIntraday
                dDateMax = dNewGetEnd
                If dNewGetStart >= dDateMin Then
                    If dNewGetStart <= dNewGetEnd Then
                        If Abs(dNewGetEnd - dNewGetStart) <= kMaxRangeIntraday Then
                            bValid = True
                        Else
                            MsgBox "You can only download " & Str(kMaxRangeIntraday) & " days of data at a time."
                        End If
                    Else
                        MsgBox "The 'From: date' cannot be later than " & DateFormat(dNewGetEnd) & "."
                    End If
                Else
                    MsgBox "The 'From: date' cannot be earlier than " & DateFormat(dDateMin) & "."
                End If
            Else
                MsgBox "The 'To: date' cannot be later than " & DateFormat(dDateMax) & "."
            End If
        Else
            MsgBox "The 'To: date' cannot be earlier than " & DateFormat(dDateMin) & "."
        End If
    End If
    
    IsValidIntradayGet = bValid

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSingleSymHistory.IsValidIntradayGet", eGDRaiseError_Raise
    Resume ErrExit

End Function

