Attribute VB_Name = "mExGridDll"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'enum & struct for ExGrid base class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type ExGrid_Text_Specs
    Color As Long
    FontBold As Long
    FontItalic As Long
    FontUnderline As Long
    FontSize As Long
    gshFontName As Long
End Type

Type ExGrid_Extended_Properties
    gdProfiler As Long      '1=turn on gdProfiler for exgrid dll
    gridBkColor As Long     'background color for grid (default is white)
    GridLines As Long       '0=no lines, 1=lines
    reserved3 As Long       'usage TBD
    reserved4 As Long       'usage TBD
    reserved5 As Long
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'enum & struct specific to MktProfile grid (sub class of ExGrid)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum MktProfile_Char_Sequence
    MktProf_CHR_SeqOne = 0      'A-Z, a-z, 0-9      (this is default)
    MktProf_CHR_SeqTwo          'A-Z, 0-9, a-z
    MktProf_CHR_SeqThree        'a-z, A-Z, 0-9
    MktProf_CHR_SeqFour         'a-z, 0-9, A-Z
    MktProf_CHR_SeqFive         '0-9, A-Z, a-z
    MktProf_CHR_SeqSix          '0-9, a-z, A-Z
    MktProf_CHR_Blocks          'use blocks, no characters
    MktProf_CHR_Classic
End Enum

Public Enum MktProfile_Vol_Text
    MktProf_Text_None = 0
    MktProf_Text_Percent           'volume percent only (no TPO)
    MktProf_Text_Actual            'volume actual only (no TPO)
    MktProf_Text_Both              'volume percent AND actual (no TPO)
    MktProf_Text_Tpo               'TPO only (no volume)
    MktProf_Text_TpoVolPercent     'TPO and volume percent
    MktProf_Text_TpoVolActual      'TPO and volume actual
    MktProf_Text_All               'TPO and volume percent and volume actual
End Enum

Public Enum MktProfile_Char_ColorScheme
    MktProf_Color_Gradient = 0
    MktProf_Color_Rainbow
    MktProf_Color_OpenClose
    MktProf_Color_BidAsk
    MktProf_Color_VolIterator
    MktProf_Color_NoScheme
End Enum

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Color Schemes:
'   1. rainbow:     color1, color2 ignored
'   2. gradient:    color1 = from color, color2 = to color
'   3. up/down:     color1 = up color, color2 = down color
'   4. bid/ask:     color1 = bid color, color2 = ask color
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type MktProfile_Display
    CharSequence As Long    'use char_sequence enum
    ExtraRows As Long       'build <= 46: mode/POC Vol color, build >= 47: # of extra rows above/below
    ColorScheme As Long     'build <= 46: mean color, build >= 47: enumerated color scheme
    VolumeColor As Long     'single color for now
    VolumeText As Long      'MktProfile_Vol_Text      'enum for whether to show volume values
    Color1 As Long          'Color1 / Color2 are color schemes context dependendent (see note above)
    Color2 As Long
    BoxFirst As Long        '1=start sequence with rectangular boxes
End Type

Type MktProfile_Stats_Prop
    POC_Vol_Color As Long
    POC_Vol_PenSize As Long
    POC_TPO_Color As Long
    POC_TPO_PenSize As Long
    VA_Vol_Color As Long
    VA_Vol_PenSize As Long
    VA_TPO_Color As Long
    VA_TPO_PenSize As Long
    IB_Color As Long
    IB_PenSize As Long
    StdDev_Color As Long
    StdDev_PenSize As Long
    Mean_Color As Long
    Mean_PenSize As Long
    reserved1 As Long
    reserved2 As Long
    reserved3 As Long
    reserved4 As Long
    reserved5 As Long
    reserved6 As Long
    VA_Vol_Percent As Double
    VA_TPO_Percent As Double
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'functions specific to market profile GridControl (some are overrides of ExGrid functions)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'constructor/destructor type functions
Declare Function gxMktProfileNew& Lib "Exgrid.dll" ()
Declare Function gxMktProfileInit& Lib "Exgrid.dll" (ByVal hWndParent&, ByVal hWndChild&, ByVal pGridCtrl&)
Declare Function gxMktProfileDestroy& Lib "Exgrid.dll" (ByVal pGridCtrl&)

'data-related functions
Declare Function gxMktInitPrices& Lib "Exgrid.dll" (ByVal pGridCtrl&, ByVal pProfileBars&, ByVal pMinuteBars&, ByVal pSmpBars&, ByVal dtBegin#, ByVal dtEnd#, ByVal iProfileCount&, ByVal iExtraRows&, ByVal iTicksPerRow&)
Declare Function gxMktProfileData& Lib "Exgrid.dll" (ByVal pGridCtrl&, ByVal aBarsHandles&, ByVal nProfileInterval&, ByVal nTotalProfiles&, ByVal dtBegin#, ByVal dtEnd#, ByVal nSMP&)
Declare Function gxMktUpdateRT& Lib "Exgrid.dll" (ByVal pGridCtrl&, ByVal pBarsRT&, ByVal nTBD&)

'display-related functions
Declare Sub gxMktGridRefresh Lib "Exgrid.dll" (ByVal pGridCtrl&)
Declare Function gxMktProfileProperties& Lib "Exgrid.dll" (ByVal pGridCtrl&, pProfileProp As MktProfile_Display)
Declare Function gxMktGridTextSpecs& Lib "Exgrid.dll" (ByVal pGridCtrl&, pFontSpecs As ExGrid_Text_Specs)
Declare Function gxMktGridExtendedProp& Lib "Exgrid.dll" (ByVal pGridCtrl&, pExtendeProp As ExGrid_Extended_Properties)
Declare Function gxMktStatisticsProp& Lib "Exgrid.dll" (ByVal pGridCtrl&, pVAProp As MktProfile_Stats_Prop)
Declare Function gxMktCenterPrice& Lib "Exgrid.dll" (ByVal pGridCtrl&)
Declare Function gxMktVolIteratorBars Lib "Exgrid.dll" (ByVal pGridCtrl&, ByVal pBars&) As Long
Declare Function gxMktAutoTicks Lib "Exgrid.dll" (ByVal pGridCtrl&) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'callback function to give status on DLL's data processing progress
'   nStatus is currently not used, but may be used in future for error feedback
'   hString has text that can be used to display in status area of window owning hWnd
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Declare Sub gxMktProfileCallback Lib "Exgrid.dll" (ByVal pGridCtrl&, ByVal fnCallback&)

Function fnDLLStatus(ByVal nStatus&, ByVal hString&, ByVal hWnd&) As Long
On Error Resume Next:

    Dim i&
    
    If hString = 0 Or hWnd = 0 Then Exit Function
    If gdGetSize(hString) <= 0 Then Exit Function

    'find the market profile form that goes with this hwnd
    For i = 0 To Forms.Count - 1
        If TypeOf Forms(i) Is frmMarketProfile Then
            If Forms(i).hWnd = hWnd Then
                Forms(i).DLLStatusInfo nStatus, hString&
                Exit For
            End If
        End If
    Next

End Function

