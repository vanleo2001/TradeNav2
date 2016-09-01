VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTnCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTnCore.cls
'' Description: Common application side functions for the DLL's
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 05/20/2015   DAJ         Added the GetPrintHeader function
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDialogMessage
'' Description: Add a dialog message to be shown with a timer
'' Inputs:      Message, Caption, Icon, Buttons
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddDialogMessage(ByVal strMessage As String, Optional ByVal strIcon As String = "", Optional ByVal strButtons As String = "", Optional ByVal strCaption As String = "")
    frmOnlineBroker.AddDialogMessage strMessage, strIcon, strButtons, strCaption
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeGridFont
'' Description: Allow the user to change the font for the given grid
'' Inputs:      Grid, Resize Columns?
'' Returns:     True if done, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ChangeGridFont(fg As VSFlexGrid, Optional ByVal bResizeColumns As Boolean = True) As Boolean
    ChangeGridFont = mMain.ChangeGridFont(fg, bResizeColumns)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CurrentTime
'' Description: If streaming get the feed time, otherwise return now
'' Inputs:      To Time Zone, Symbol, Allow Replay Time?
'' Returns:     Feed Time or Now
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CurrentTime(Optional ByVal strToTimeZone As String = "", Optional ByVal strSymbol As String = "", Optional ByVal bAllowReplayTime As Boolean = False) As Double
    CurrentTime = mDataNav.CurrentTime(strToTimeZone, strSymbol, bAllowReplayTime)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DoPrintHeader
'' Description: Do the print header for Trade Navigator
'' Inputs:      Font Size, Printer Object
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoPrintHeader(Optional ByVal lFontSize As Long = 12, Optional ByVal vp As VSPrinter = Nothing)
    mMain.DoPrintHeader lFontSize, vp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatVersion
'' Description: Formatted version of the Trade Navigator version
'' Inputs:      Include Revision?, Include Date?
'' Returns:     Formatted version
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatVersion(Optional ByVal bIncludeRevision As Boolean = False, Optional ByVal bIncludeFileDate As Boolean = False) As String
    FormatVersion = mMain.FormatVersion(bIncludeRevision, bIncludeFileDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetMachineID
'' Description: Get the machine ID
'' Inputs:      None
'' Returns:     Machine ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetMachineID() As String
    GetMachineID = RI_GetMachineID
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPassword
'' Description: Get the pasword
'' Inputs:      None
'' Returns:     Password
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPassword() As String
    GetPassword = Trim(RI_GetUserPassword)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbol
'' Description: Get a symbol for the given Symbol or Symbol ID
'' Inputs:      Symbol or Symbol ID
'' Returns:     Symbol
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbol(ByVal vSymbolOrSymbolID As Variant) As String
    GetSymbol = mDmDll.GetSymbol(vSymbolOrSymbolID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolID
'' Description: Get a symbol ID for the given Symbol or Symbol ID
'' Inputs:      Symbol or Symbol ID
'' Returns:     Symbol ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSymbolID(ByVal vSymbolOrSymbolID As Variant) As Long
    GetSymbolID = mDmDll.GetSymbolID(vSymbolOrSymbolID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridScrollCheck
'' Description: Help fix the inadvertant scrolling issues with the FlexGrids
''              while streaming is on
'' Inputs:      Grid, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GridScrollCheck(fg As VSFlexGrid, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    mMain.GridScrollCheck fg, OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasModule
'' Description: Determine if the user has the given module(s)
'' Inputs:      Module(s), Include Source Code?
'' Returns:     True if have module(s), False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasModule(ByVal strModules As String, Optional ByVal bIncludeSourceCode As Boolean = False) As Boolean
    HasModule = mMain.HasModule(strModules, bIncludeSourceCode)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastKnownPrice
'' Description: Get the last known price for the given symbol
'' Inputs:      Symbol, Local Time?, Date/Time of last price(return)
'' Returns:     Last Known Price (kNullData if not found)
''
'' iBidTradeOrAsk:  -1 = LastBid,  0 = LastTrade,  1 = LastAsk
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastKnownPrice(ByVal vSymbolOrSymbolID As Variant, Optional ByVal iBidTradeOrAsk As Integer = 0, Optional ByVal bLocalTime As Boolean = False, Optional dDateTime As Double) As Double
    LastKnownPrice = g.RealTime.LastKnownPrice(vSymbolOrSymbolID, iBidTradeOrAsk, bLocalTime, dDateTime)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RaiseError
'' Description: Raise an error appropriately
'' Inputs:      Source, Mode, Path
'' Returns:     True if has had errors, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RaiseError(Optional ByVal strErrSource As String = "", Optional ByVal Mode As eGDRaiseErrorMode = eGDRaiseError_Default, Optional ByVal strPath As String = "") As Boolean
    RaiseError = mGenesis.RaiseError(strErrSource, Mode, strPath)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowForm
'' Description: Show the given form in the given mode
'' Inputs:      Form to Show, Mode, Owner Form, Allow Off Screen?
'' Returns:     None
''
'' (FYI: ActModal is needed for things like a form with the Tradesense editor control
''  which has a non-modal popup window being displayed overtop the control)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowForm(frm As Form, Optional ByVal eModal As eShowFormMode = eForm_Nonmodal, Optional ByVal frmOwner As Form = Nothing, Optional ByVal bAllowOffScreen As Boolean = False)
    If frmOwner Is Nothing Then
        Set frmOwner = frmMain
    End If
    mGenesis.ShowForm frm, eModal, frmOwner, bAllowOffScreen
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimerStart
'' Description: Notify that a timer is starting
'' Inputs:      Name of the timer
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TimerStart(ByVal strName As String)
    mMain.TimerStart strName
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TimerEnd
'' Description: Notify that a timer is ending
'' Inputs:      Name of the timer, Interval of the timer
'' Returns:     True if logged the time, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TimerEnd(ByVal strName As String, ByVal lInterval As Long) As Boolean
    TimerEnd = mMain.TimerEnd(strName, lInterval)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolbarIcon
'' Description: Return the toolbar icon name for the given ID
'' Inputs:      ID
'' Returns:     Toolbar Icon name
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToolbarIcon(ByVal strID As String) As String
    ToolbarIcon = mMain.ToolbarIcon(strID)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Unloading
'' Description: Is Trade Navigator currently unloading?
'' Inputs:      None
'' Returns:     True if unloading, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Unloading() As Boolean
    Unloading = g.bUnloading
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowPrint
'' Description: Show the print form
'' Inputs:      Name, Form, Arguments, Top Margin, Bottom Margin, Left Margin,
''              Right Margin, Landscape?, Override?, To File Options,
''              Call after header?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowPrint(ByVal strName As String, ByVal objReport As Object, Optional ByVal vArgs As Variant = 0, Optional dTopMargin# = 1, Optional ByVal dBottomMargin# = 1, Optional ByVal dLeftMargin# = 1, Optional ByVal dRightMargin# = 1, Optional ByVal bLandscape As Boolean = False, Optional ByVal bOverride As Boolean = False, Optional ByVal ToFileOptions As ePrintToFile = ePrintToFile_Both, Optional ByVal bCallAfterHeaderEvent As Boolean = False) As Boolean
    ShowPrint = frmPrintPreview.ShowMe(strName, objReport, vArgs, dTopMargin, dBottomMargin, dLeftMargin, dRightMargin, bLandscape, bOverride, ToFileOptions, bCallAfterHeaderEvent)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPrintHeader
'' Description: Get the header for the pages during a print
'' Inputs:      None
'' Returns:     Header
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPrintHeader() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    If ExtremeCharts >= 1 Then
        strReturn = "|Extreme Charts -- "
    Else
        strReturn = "|Trade Navigator" & vbCrLf & "Genesis Financial Technologies - "
    End If
    strReturn = strReturn & GetProvidedProperty("Website", , True)
    
    GetPrintHeader = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTnCore.GetPrintHeader"
    
End Function
