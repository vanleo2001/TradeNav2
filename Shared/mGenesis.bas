Attribute VB_Name = "mGenesis"
Option Explicit

Public Enum eShowFormMode
    eForm_Nonmodal = 0
    eForm_Modal = 1
    eForm_ActModal = 2
End Enum

' used by "DateFormat"
Public Enum eDateFormat
    NO_DATE = -1
    MM_DD_YYYY = 0
    MM_DD_YY = 1
    M_D_YY = 2
    M_D = 3
    MMM_YY = 4
End Enum
Public Enum eTimeFormat
    NO_TIME = -1
    HH_MM = 0
    H_MM = 1
    HH_MM_SS = 2
    H_MM_SS = 3
End Enum
Public Enum eAmPmFormat
    NO_AMPM = -1
    AMPM_UPPER = 0
    AMPM_LOWER = 1
    AP_UPPER = 2
    AP_LOWER = 3
End Enum

Public Enum vbWindowState
    wsNormal = 0
    wsMinimized = 1
    wsMaximized = 2
End Enum

Public Enum eGridMode
    eGridMode_Grid = 0
    eGridMode_List = 1
    eGridMode_Tree = 2
End Enum

Public Enum eGDAlignment
    eGDAlign_Left = 0
    eGDAlign_Right = 1
    eGDAlign_Center = 2
End Enum

Public Enum eGDRaiseErrorMode
    eGDRaiseError_Default = -1
    eGDRaiseError_Init = 0
    eGDRaiseError_Show = 1
    eGDRaiseError_Raise = 2
    eGDRaiseError_HasHadErrors = 3
End Enum

' Folder ID's which can be used by the SpecialFolderPath function -- please NOTE:
' - none of the "virtual" folders get returned by the SpecialFolderPath function
' - not all of the following work for all operating systems
' - Win98 has a very small subset of these (e.g. none of the "_ALLUSERS_" folders)
' - XP has most of the following, Vista has a few more
' - but the "x86" folders are only returned by 64-bit operating systems
' - the "Quick Launch" is simply built from the USER_APPDATA folder (works for XP and Vista)
' - FYI: the {All Users} folders in Vista are usually under either C:\Users\ or C:\ProgramData\
Public Enum CSIDL_FOLDERS
    'CSIDL_DESKTOP = &H0            '// The Desktop - virtual folder
    'CSIDL_INTERNET = &H1           '//Internet virtual folder
    CSIDL_USER_PROGRAMS = &H2       ' {user}\Start Menu\Programs\
    'CSIDL_CONTROLS = &H3           '// Control Panel - virtual folder
    'CSIDL_PRINTERS = &H4           '// Printers - virtual folder
    CSIDL_USER_MYDOCUMENTS = &H5    ' {user}\My Documents\
    CSIDL_USER_FAVORITES = &H6      ' {user}\Favorites\
    CSIDL_USER_STARTUP = &H7        ' {user}\Start Menu\Programs\Startup\
    CSIDL_USER_RECENT = &H8         ' {user}\Recent\
    CSIDL_USER_SENDTO = &H9         ' {user}\SendTo\
    'CSIDL_BITBUCKET = &HA          '// Recycle Bin - virtual folder
    CSIDL_USER_STARTMENU = &HB      ' {user}\Start Menu\
    CSIDL_USER_MYMUSIC = &HD        ' {user}\My Documents\My Music\
    CSIDL_USER_MYVIDEOS = &HE       ' {user}\My Documents\My Videos\
    CSIDL_USER_DESKTOP = &H10       ' {user}\Desktop\
    'CSIDL_DRIVES = &H11            '// My Computer - virtual folder
    'CSIDL_NETWORK = &H12           '// Network Neighbourhood - virtual folder
    CSIDL_USER_NETHOOD = &H13       ' {user}\NetHood\
    CSIDL_FONTS = &H14              ' C:\Windows\Fonts\
    CSIDL_USER_TEMPLATES = &H15     ' {user}\Templates\  (SHELLNEW)
    CSIDL_ALLUSERS_STARTMENU = &H16 ' {All Users}\Start Menu\
    CSIDL_ALLUSERS_PROGRAMS = &H17  ' {All Users}\Start Menu\Programs\
    CSIDL_ALLUSERS_STARTUP = &H18   ' {All Users}\Start Menu\Programs\Startup\
    CSIDL_ALLUSERS_DESKTOP = &H19   ' {All Users}\Desktop\
    CSIDL_USER_APPDATA = &H1A       ' {user}\Application Data\
    CSIDL_USER_PRINTHOOD = &H1B     ' {user}\PrintHood\
    CSIDL_USER_LOCAL_APPDATA = &H1C ' {user}\Local Settings\Application Data\
    'CSIDL_USER_ALTSTARTUP = &H1D   ' non localized startup (not supported by XP)
    'CSIDL_ALLUSERS_ALTSTARTUP = &H1E ' non localized common startup (not supported by XP)
    CSIDL_ALLUSERS_FAVORITES = &H1F ' {All Users}\Favorites\
    CSIDL_USER_INTERNET_CACHE = &H20 ' {user}\Local Settings\Temporary Internet Files\
    CSIDL_USER_COOKIES = &H21       ' {user}\Cookies\
    CSIDL_USER_HISTORY = &H22       ' {user}\Local Settings\History\
    CSIDL_ALLUSERS_APPDATA = &H23   ' {All Users}\Application Data\
    CSIDL_WINDOWS = &H24            ' C:\Windows\
    CSIDL_SYSTEM = &H25             ' C:\Windows\System32\
    CSIDL_PROGRAMFILES = &H26       ' C:\Program Files\
    CSIDL_USER_MYPICTURES = &H27    ' {user}\My Documents\My Pictures\
    CSIDL_USER = &H28               ' {user}\
    CSIDL_SYSTEMx86 = &H29          ' system folder for x86 apps (Alpha)
    CSIDL_PROGRAMFILESx86 = &H2A    ' Program Files folder for x86 apps (Alpha)
    CSIDL_PROGRAMFILES_COMMONFILES = &H2B ' C:\Program Files\Common Files\
    CSIDL_PROGRAMFILES_COMMONx86 = &H2C 'x86 \Program Files\Common on RISC
    CSIDL_ALLUSERS_TEMPLATES = &H2D ' {All Users}\Templates\
    CSIDL_ALLUSERS_DOCUMENTS = &H2E ' {All Users}\Documents\
    CSIDL_ALLUSERS_ADMINTOOLS = &H2F ' {All Users}\Start Menu\Programs\Administrative Tools\
    CSIDL_USER_ADMINTOOLS = &H30    ' {user}\Start Menu\Programs\Administrative Tools\
    CSIDL_ALLUSERS_MYMUSIC = &H35   ' {All Users}\Documents\My Music\
    CSIDL_ALLUSERS_MYPICTURES = &H36 ' {All Users}\Documents\My Pictures\
    CSIDL_ALLUSERS_MYVIDEOS = &H37  ' {All Users}\Documents\My Videos\
    CSIDL_RESOURCES = &H38          ' C:\WINDOWS\resources\
    CSIDL_USER_CDBURN = &H3B        ' {user}\Local Settings\Application Data\Microsoft\CD Burning\
    CSIDL_QUICKLAUNCH = &HFF        ' Quick Launch folder: USER_APPDATA\Microsoft\Internet Explorer\Quick Launch\
End Enum

' General constants
Public Const PI = 3.141593
Public Const MIN_INTEGER = -32768#
Public Const MAX_INTEGER = 32767#
Public Const NULL_DATA = -999999
Public Const kDarkThemeColor = 2105376 'RGB(32, 32, 32)

' API constants
Public Const WM_USER = &H400
Public Const ABOVE_PRIORITY_CLASS = &H8000
Public Const ATTR_DIRECTORY% = 16
Public Const BELOW_PRIORITY_CLASS = &H4000
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DFACE = 15
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_3DHILIGHT = 20
Public Const COLOR_3DSHADOW = 16
Public Const CREATE_NEW_CONSOLE = &H10
Public Const DFC_BUTTON = 4
Public Const DFC_CAPTION = 1
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_CAPTIONCLOSE = 0&
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CHECKED = &H400
Public Const DFCS_PUSHED As Long = &H200&
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
Public Const DST_BITMAP = 4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINESCROLL = &HB6
Public Const EM_SETTABSTOPS = 203
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_WNDPROC As Long = -4
Public Const HC_ACTION = 0
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const HTCAPTION = 2
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const LB_FINDSTRINGEXACT = WM_USER + 35
Public Const LB_SETTABSTOPS = WM_USER + 19
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0
Public Const MF_UNCHECKED = &H0
Public Const MF_CHECKED = &H8
Public Const MF_GRAYED = &H1&
Public Const MF_HILITE = &H80
Public Const MF_BYCOMMAND = &H0
Public Const MF_BYPOSITION = &H400
Public Const MF_POPUP = &H10
Public Const MF_STRING = &H0
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const SC_CLOSE = &HF060&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_MINIMIZE = &HF020&
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXSMSIZE = 52
Public Const SM_CYSMSIZE = 53
Public Const SM_CYCAPTION = 4
Public Const SM_CYSMCAPTION = 51
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_HIDE = 0
Public Const SW_OTHERUNZOOM = 4
Public Const SW_RESTORE = 9
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNA = 8
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const WH_MOUSE = 7
Public Const WM_ACTIVATE = &H6
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COPYDATA = &H4A
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETTEXT = &HD
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_PAINT = &HF
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SIZE = &H5
Public Const WM_STYLECHANGED = &H7D
Public Const WM_STYLECHANGING = &H7C
Public Const WM_SYSCOMMAND = &H112
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_THEMECHANGED = &H31A
Public Const MIIM_ID = &H2&
Public Const TPM_RETURNCMD = &H100
Public Const CB_SHOWDROPDOWN = &H14F
' Window Styles (used with GWL_STYLE)
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_CAPTION = WS_BORDER Or WS_DLGFRAME
Public Const WS_THICKFRAME = &H40000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_CHILD = &H40000000 ' cannot be used with WS_POPUP
Public Const WS_POPUP = &H80000000 ' cannot be used with WS_CHILD
Public Const WS_VISIBLE = &H10000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_CLIPCHILDREN As Long = &H2000000
' Extended Window Styles (used with GWL_EXSTYLE)
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4
Public Const WS_EX_TOPMOST As Long = &H8
Public Const WS_EX_ACCEPTFILES As Long = &H10
Public Const WS_EX_TRANSPARENT As Long = &H20
Public Const WS_EX_MDICHILD As Long = &H40
Public Const WS_EX_TOOLWINDOW As Long = &H80
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_CLIENTEDGE As Long = &H200
Public Const WS_EX_CONTEXTHELP As Long = &H400
Public Const WS_EX_RIGHT As Long = &H1000
Public Const WS_EX_LEFT As Long = &H0
Public Const WS_EX_RTLREADING As Long = &H2000
Public Const WS_EX_LTRREADING As Long = &H0
Public Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Public Const WS_EX_RIGHTSCROLLBAR As Long = &H0
Public Const WS_EX_CONTROLPARENT As Long = &H10000
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_LAYERED = &H80000
'Dialog Styles (also present in the GWL_STYLE area)
Public Const DS_ABSALIGN As Long = &H1
Public Const DS_SYSMODAL As Long = &H2
Public Const DS_3DLOOK As Long = &H4
Public Const DS_FIXEDSYS As Long = &H8
Public Const DS_NOFAILCREATE As Long = &H10
Public Const DS_LOCALEDIT As Long = &H20 'Edit items get Local storage.
Public Const DS_SETFONT As Long = &H40 'User specified font for Dlg controls
Public Const DS_MODALFRAME As Long = &H80 'Can be combined with WS_CAPTION
Public Const DS_NOIDLEMSG As Long = &H100 'WM_ENTERIDLE message will not be sent
Public Const DS_SETFOREGROUND As Long = &H200 'not in win3.1
Public Const DS_CONTROL As Long = &H400
Public Const DS_CENTER As Long = &H800
Public Const DS_CENTERMOUSE As Long = &H1000
Public Const DS_CONTEXTHELP As Long = &H2000
' for Transparency
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4


' API structures
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        ShowCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As Rect
End Type

Public Type STARTUPINFO
        cb As Long
        lpReserved As Long 'String
        lpDesktop As Long 'String
        lpTitle As Long 'String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type
' The same type with HWORD and LWORD seperated
Type VS_FIXEDFILEINFO2
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As Rect
    rcWork As Rect
    dwFlags As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    ftype As Long           ' used if MIIM_TYPE (4.0) or MIIM_FTYPE (>4.0)
    fState As Long          ' used if MIIM_STATE
    wID As Long             ' used if MIIM_ID
    hSubMenu As Long        ' used if MIIM_SUBMENU
    hbmpChecked As Long     ' used if MIIM_CHECKMARKS
    hbmpUnchecked As Long   ' used if MIIM_CHECKMARKS
    dwItemData As Long      ' used if MIIM_DATA
    dwTypeData As String    ' used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
    cch As Long             ' used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
    hbmpItem As Long        ' used if MIIM_BITMAP
End Type

Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type MEMORY_STATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
' This is needed for the larger memory variables in Windows NT and above
Private Type MEMORY_STATUS_EX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As ULARGE_INTEGER
    ullAvailPhys As ULARGE_INTEGER
    ullTotalPageFile As ULARGE_INTEGER
    ullAvailPageFile As ULARGE_INTEGER
    ullTotalVirtual As ULARGE_INTEGER
    ullAvailVirtual As ULARGE_INTEGER
    ullAvailExtendedVirtual As ULARGE_INTEGER
End Type

' stuff used by BrowseForFolder ...
Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfnCallback     As Long
   lParam           As Long
   iImage           As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal PV As Long)


'API: Windows
Private Declare Function AllowSetForegroundWindowAPI Lib "user32" Alias "AllowSetForegroundWindow" (ByVal hProcessID As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetLastActivePopup Lib "user32" (ByVal hWndOwner As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, lpMonInfo As MONITORINFO) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpWndPl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MonitorFromRect Lib "user32" (lpRect As Rect, ByVal dwFlags As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpWndPl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
' transparency
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

'API: menus
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPos As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal nPos As Long, ByVal nFlag As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal nItem As Long, ByVal nEnable As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal nFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Rect) As Long

'API: memory
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDestination As Any, lpSource As Any, ByVal dwBytes As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDestination As Any, ByVal dwBytes As Long, ByVal cFill As Byte)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDestination As Any, lpSource As Any, ByVal dwBytes As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDestination As Any, ByVal dwBytes As Long)
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapCompact Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Public Declare Function HeapSetInformation Lib "kernel32" (ByVal hHeap As Long, ByVal dwHeapInformationClass As Long, ByRef pHeapInformation As Long, ByVal dwHeapInformationLength As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORY_STATUS)
Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORY_STATUS_EX) As Long
'Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, lpProcessAffinityMask As Long, lpSystemAffinityMask As Long) As Long

'API: Ini files
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'API: processes
'Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Any, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As Any, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
' (the following are not supported in 95/95/ME)
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

'API: Version information
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

' API: custom properties for a window
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpsz As String, ByVal hData As Long) As Long 'Bool
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpsz As String) As Long 'HANDLE
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpsz As String) As Long 'HANDLE

'API: Misc.
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function SetThreadExecutionState Lib "kernel32" (ByVal iFlags As Long) As Long
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
' call "PlaySoundFile" to use this function:
Private Declare Function PlaySound Lib "winmm" (ByVal strFile$, ByVal hModule&, ByVal nFlag&) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal strDrive As String) As Long
Public Declare Sub SetThemeAppProperties Lib "UxTheme" (ByVal uFlags As Long)
Public Declare Function SetWindowTheme Lib "UxTheme" (ByVal hWnd As Long, lpSubAppName As String, lpSubIdList As String) As Long

' hooks
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Sub MouseEvent Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

'API: Icons
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function LoadIconNum Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

'API: Drawing
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function SetRect Lib "user32" (rRect As Rect, ByVal nLeft&, ByVal nTop&, ByVal nRight&, ByVal nBottom&) As Long

'API: current Mouse/Key state
Public Const VK_LBUTTON = &H1   ' Left mouse button
Public Const VK_RBUTTON = &H2   ' Right mouse button
Public Const VK_MBUTTON = &H4   ' Middle mouse button
Public Const VK_CONTROL = &H11  ' Ctl key
Public Const VK_SHIFT = &H10    ' Shift key
Public Const VK_MENU = &H12     ' Alt key
Public Const VK_CANCEL = &H3    ' Ctrl-Break
Public Const VK_PAUSE = &H13    ' Pause key
Public Const VK_ESCAPE = &H1B   ' Escape key
Public Const VK_CAPITAL = &H14  ' Caps Lock
Public Const VK_NUMLOCK = &H90  ' Num Lock
Public Const VK_SCROLL = &H91   ' Scroll Lock
'(use "KeyIsPressed" and "MouseIsPressed" for easier use of this function)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' Vista/Win7 specific
Private Declare Function DwmEnableComposition Lib "dwmapi" (ByVal uiEnable As Long) As Long
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi" (bEnabled As Long) As Long
Private Declare Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hWnd As Long, ByVal dwAttrib As Long, pvAttrib As Long, ByVal cbAttrib As Long) As Long

' ZIP functions --------------------------------------
Type file_info
    Size As Long
    JDate As Long
    Millisec As Long
    Year As String * 1 ' add 1900
    Month As String * 1
    Day As String * 1
    Hour As String * 1
    Min As String * 1
    Sec As String * 1
    Attrib As Integer 'String * 1
    FileName As String * 14
    FullName As String * 150
End Type

'Zip actions: 'C'=Create, 'A'=Append, 'D'=Delete, 'U'=Unzip,
'        'N'=Number, 'S'=Size(uncompressed), 'F'=Find matching files
Type zip_args
    Action As String * 1
    Version As String * 1
    zip_file As String * 150
    file_masks As String * 150
    unzip_path As String * 150
    err_msg As String * 150
    file_window As Long 'Integer
    progress_window As Long 'Integer
    start_date As Long
    end_date As Long
    overwrite_newer As Integer
    strip_paths As Integer
    recursive As Integer
    compress_level As Integer
End Type
Public Declare Function GenZip& Lib "G32_ZIP.DLL" (Args As zip_args)
Public Declare Sub GenZipAbort Lib "G32_ZIP.DLL" (ByVal bAbort%)
Public Declare Function ZipMemory& Lib "G32_ZIP.DLL" (ByVal nSourcePtr&, ByVal nSourceSize&, ByVal nDestPtr&, ByVal nDestSize&, nCrc&)
' NOTE: discovered that you should always allocate a little more for the destination
' than what will actually end up being used -- don't know for sure how much.
Public Declare Function UnzipMemory& Lib "G32_ZIP.DLL" (ByVal nSourcePtr&, ByVal nSourceSize&, ByVal nDestPtr&, ByVal nDestSize&, nCrc&)

''Public Declare Function ArcType% Lib "G32_ZIP.DLL" (ByVal strFilename$)
''Public Declare Function BreakupMultiFile% Lib "G32_ZIP.DLL" (ByVal strMultiFile$, ByVal strPath$, ByVal nUnzipMode%)


'--- MISC other functions in GEN_ZIP ---------------

' Dates

' fixes DateTime (rounds to nearest millisecond) in order to avoid precision errors
Public Declare Function gdFixDateTime Lib "G32_ZIP.DLL" (ByVal dDateTime As Double) As Double
' newer version allows rounding arg: False = round to Millisecond, True = round to Second
Public Declare Function gdFixDateTime2 Lib "G32_ZIP.DLL" (ByVal dDateTime As Double, ByVal bRoundToSecond&) As Double

' much faster than the original VB routines, and avoids the rounding issues (at least as much as possible)
Public Declare Function gdRoundNum Lib "G32_ZIP.DLL" (ByVal dValue As Double, ByVal iDecimalDigits&) As Double
Public Declare Function SigDigits Lib "G32_ZIP.DLL" Alias "gdRoundSigDigits" (ByVal dValue As Double, ByVal iSigDigits&) As Double
    
' gdIsHoliday returns the following byte if date is one of these holidays:
'   'N'ewYears, 'K'ings, 'P'residents, 'G'oodFriday, 'E'aster, 'M'emorial,
'   'J'uly4, 'L'aborDay, 'C'olumbus, 'V'eterans, 'T'hanksgiving, 'X'mas
' - strHolidays is string of holidays to check for (empty string defaults to all)
' - returns 0 if the date is not one of the holidays in strHolidays
' - the holiday rules are as follows:
'   'N' = New Years Day (Jan. 1, but if Sunday then observed Monday)
'   'K' = M.L.King's Birthday (3rd Mon. of Jan.)
'   'P' = President's Day (3rd Mon. of Feb.)
'   'G' = Good Friday (Fri. before Easter)
'   'E' = Easter (uses special astronomical algorithm)
'   'M' = Memorial Day (last Mon. of May)
'   'J' = July 4th (nearest weekday to Jul. 4)
'   'L' = Labor Day (1st Mon. of Sep.)
'   'C' = Columbus Day (2nd Mon. of Oct.)
'   'V' = Veterans Day (nearest weekday to Nov. 11)
'   'T' = Thanksgiving (4th Thu. of Nov.)
'   'X' = Christmas (nearest weekday to Dec. 25)
Public Declare Function gdIsHoliday Lib "G32_ZIP.DLL" (ByVal JDate&, ByVal strHolidays$) As Byte

' To get date of one of the following holidays for the specified year:
'   'N'ewYears, 'K'ings, 'P'residents, 'G'oodFriday, 'E'aster, 'M'emorial,
'   'J'uly4, 'L'aborDay, 'C'olumbus, 'V'eterans, 'T'hanksgiving, 'X'mas
' or for Daylight Saving Time (started last Sun. of April from 1966-1986,
' except for 1/6/74 and 2/23/75, then first Sun. of April since 1987):
'   'S' = Spring/Start daylight saving time (actual dates since 1966)
'   'F' = Fall/Finish daylight saving time (last Sun. of Oct. since 1966)
Public Declare Function gdGetHoliday Lib "G32_ZIP.DLL" (ByVal strHoliday$, ByVal nYear&) As Long

' format for "rule": [Day][B][m][+/-#][B], where ...
'   Day  - base day to start from: just a number for day of month, or
'   F/L for first/last day of month, or "nW" for Nth weekday
'       (Weekday: "S"un, "M"on, "T"ue, "W"ed, t"H"u, "F"ri, s"A"t)
'   B    - optional, to force the base day to a business day
'   m    - optional month adjustment: "P"revious or "N"ext month
'   +/-# - plus or minus # of days (with "B" means # of business days)
' e.g. rule: "3MP-2B"  = 2 bus. days before 3rd Mon. of Prev. month
' e.g. rule: "LB-4B"   = 4 bus. days before last business day of month
' e.g. rule: "FB+3"    = 3 days after first business day of month
' e.g. rule: "25N-4B"  = 4 bus. days before 25th of next month
' e.g. rule: "25BN"    = business day on or before 25th of next month
Public Declare Function GetDateFromRule Lib "G32_ZIP.DLL" (ByVal nBaseYear%, ByVal nBaseMonth%, ByVal strRuleSpec$) As Long

' Returns true if it is daylight saving time for this DateTime in the specified time zone.
' (see "ConvertTimeZone" function for description of "TimeZone" formatting)
Public Declare Function gdIsDaylightSavingTime Lib "G32_ZIP.DLL" (ByVal dDateTime As Double, ByVal strTimeZone$) As Byte

' used by the "ConvertTimeZone" function
Private Declare Function gdConvertTimeZone Lib "G32_ZIP.DLL" (ByVal dDateTime As Double, ByVal strFromTimeZone$, ByVal strToTimeZone$) As Double


' File operations
Public Declare Function FileOpen& Lib "G32_ZIP.DLL" (ByVal strFileName$, ByVal strMode$)
Public Declare Sub FileClose Lib "G32_ZIP.DLL" (hFile&)
Public Declare Function FileBinaryIO& Lib "G32_ZIP.DLL" (ByVal hFile&, vPtr As Any, ByVal nBytes&, ByVal bPut&)
Public Declare Function FileStringIO& Lib "G32_ZIP.DLL" (ByVal hFile&, ByVal strString$, ByVal nBytes&, ByVal bPut&)

Public Declare Function DeleteFolder& Lib "G32_ZIP.DLL" (ByVal strFolder$, ByVal bDeleteContents&)
Public Declare Function DeleteFiles& Lib "G32_ZIP.DLL" (ByVal strFile$, ByVal bEvenIfReadOnlyOrHidden&)
'extern "C" long EXPORTDLL CopyFiles (char *strFromSpec, char *strTo, char *strMode, BOOL bOverwriteReadOnly, BOOL bBypassNormalFileLocking)
Private Declare Function gdCopyFiles& Lib "G32_ZIP.DLL" (ByVal strFromSpec$, ByVal strTo$, ByVal strMode$, ByVal bEvenIfReadOnlyOrHidden&, ByVal bBypassNormalFileLocking&)
Private Declare Function gdMoveFiles& Lib "G32_ZIP.DLL" (ByVal strFromSpec$, ByVal strTo$, ByVal strMode$, ByVal bEvenIfReadOnlyOrHidden&)
Private Declare Function gdGetAllDrives& Lib "G32_ZIP.DLL" (ByVal strDrivesBuffer$)
Public Declare Function GetDiskSize Lib "G32_ZIP.DLL" Alias "gdGetDiskSize" (ByVal strPath$) As Double
Public Declare Function GetDiskFreeSpace Lib "G32_ZIP.DLL" Alias "gdGetDiskFreeSpace" (ByVal strPath$) As Double

Public Declare Function CalcFileCrc& Lib "G32_ZIP.DLL" (ByVal strFile$)
Public Declare Sub Ieee2Msb Lib "G32_ZIP.DLL" (fNum As Single)
Public Declare Sub Msb2Ieee Lib "G32_ZIP.DLL" (fNum As Single)
Public Declare Function JulFromLong& Lib "G32_ZIP.DLL" (ByVal nLongDate&)
Public Declare Function JulToLong& Lib "G32_ZIP.DLL" (ByVal nJulian&, ByVal bWithCentury%)
Public Declare Function FileExist% Lib "G32_ZIP.DLL" (ByVal strFileName$)
Public Declare Function DirExist% Lib "G32_ZIP.DLL" (ByVal strPathName$)
Public Declare Function GetAddress& Lib "G32_ZIP.DLL" (ptr As Any)
Public Declare Sub IdleSleep Lib "G32_ZIP.DLL" (ByVal nMilliseconds&, ByVal bSkipFullIdle&)
Public Declare Function gdTickCountVB# Lib "G32_ZIP.DLL" (ByVal bSkipHighRes&)

' Call "GetNextFile" first time with id=0; passes back id to use thereafter.
' (to abort prematurely, pass empty "root")
''Public Declare Function GetNextFile% Lib "G32_ZIP.DLL" (ID&, ByVal strRoot$, file As file_info)
''Public Declare Function WildCardMatch% Lib "G32_ZIP.DLL" (ByVal strCheck$, ByVal strPatterns$, ByVal nMode%)

' used by VB EnumWindowTitles function (now obsolete)
''Public Declare Function EnumWindowTitlesDLL& Lib "G32_ZIP.DLL" Alias "EnumWindowTitles" (ByVal hWnd&, ByVal strTitle$, ByVal nMaxLen&)

'NOTE #1: the DLL routine is faster than the VB version,
'   but is less secure (since we're passing the key)
'NOTE #2: this method is now OBSOLETE -- we should now
'   be using "gdEncrypt" using G32_GD.DLL
''Public Declare Sub OldEncryptMemory Lib "G32_ZIP.DLL" Alias "EncryptMemory" (ByVal strPassword$, ByVal strToggleString$, ByVal nStrLen&)
''Public Declare Function FileCrypt% Lib "G32_ZIP.DLL" (ByVal strPassword$, ByVal strFilename$, ByVal bDecrypt%)

' Parse numbers from mjk-type string (comma or space delimited)
'  -- puts numbers into array (max items: num_array), returns # found
Public Declare Function ParseDoubles% Lib "G32_ZIP.DLL" (ByVal strToParse$, ByVal nNumSkip%, ByVal nNumArray%, daArray#)

Public Declare Sub Parse_OHLCV Lib "G32_ZIP.DLL" (ByVal strToParse$, ByVal nNumSkip%, dOpen#, dHigh#, dLow#, dClose#, nVol&, nOI&, nTotVol&, nTotOI&)

' 11/28/97 Convert comma-quote delimited string to tab-delimited string
' (changes field delimiter to tab, strips quotes, returns number of fields)
Public Declare Function CommaQuoteToTabDelimited% Lib "G32_ZIP.DLL" (ByVal strString$)

''Public Declare Function ParseHLSTR& Lib "G32_ZIP.DLL" (ptr As Any, ByVal nOffset%, nStrLen%)
''Public Declare Sub QuickSort Lib "G32_ZIP.DLL" (ptr As Any, ByVal nDataType%, ByVal nNum%, ByVal bDescending%)

' With the following two functions, can treat an integer
'   or long or array of "n" integers as an array of 16,
'   32, or "n" * 16  Boolean flags (true/false).  For
'   both functions, pass the integer or zero element of
'   the array as the "ptr" (points to bitposition 1).
' SetBitFlag sets bit at "bitposition" to OFF (if "value"
'   is 0) or ON (if "value" is non-zero).
' GetBitFlag returns status of bit at bitposition (True/False).
Public Declare Sub SetBitFlag Lib "G32_ZIP.DLL" (ptr As Any, ByVal nBitPosition&, ByVal bValue%)
Public Declare Function GetBitFlag% Lib "G32_ZIP.DLL" (ptr As Any, ByVal nBitPosition&)


'private array for use by GetWindowHandles and EnumWindowsCallback
Private EnumWindowHandles() As Long

' these allow for setting a custom "BackColor" for the entire application
' (used by the "SetAppBackColor" and "FixFormControls" routines)
Private m_AppBackColor As Long
Private m_PrevAppBackColor As Long
Private m_bAppWhiteForeColor As Boolean

' used by callback of BrowseForFolder
Private m_InitialBrowseForFolder As String

' (this function here just to keep some common
'  variable names consistently capitalized -- why?:
'  cause otherwise it really screws up doing SourceSafe "Diffs"!)
Private Sub aDummy()
    Dim hWnd, hDC, dX, dY, dVal, eMode, hWndParent, INet, _
        Text, Tag, Caption, Value, Icon, Index, Item, Str, _
        Height, Width, Top, Bottom, Left, Right, Mid, X, Y, Parm, _
        Name, FileName, txtFileName, strFileName, ID, DataPath, _
        Show, Parse, RowData, LoadData, Version, _
        Year, Month, Day, Hour, Minute, Second, Sec, _
        Min, Max, Filter, Field, Fields, Bars, Expression, Expr, _
        List, CodedText, Mode, ColFormat, Hide, Size, _
        Length, MemPtr, Cos, Sin, Category, Source, Count, _
        FontSize, FontName, FontBold, FontItalic, FontUnderline, _
        chk, cmdOK, bInProgress, b, i, Interval, eBARS_DateTime, eBARS_Eod, _
        GridLines, tmrRealTime, UserName, Password, ToolTipText, strUrl
    'Dim hWnd, hDC, dX, dY, dVal, eMode, hWndParent, INet, _
        Text, Tag, Caption, Value, Icon, Index, Item, Str, _
        Height, Width, Top, Bottom, Left, Right, Mid, X, Y, Parm, _
        Name, FileName, txtFileName, strFileName, ID, DataPath, _
        Show, Parse, RowData, LoadData, Version, _
        Year, Month, Day, Hour, Minute, Second, Sec, _
        Min, Max, Filter, Field, Fields, Bars, Expression, Expr, _
        List, CodedText, Mode, ColFormat, Hide, Size, _
        Length, MemPtr, Cos, Sin, Category, Source, Count, _
        FontSize, FontName, FontBold, FontItalic, FontUnderline, _
        chk, cmdOK, bInProgress, b, i, Interval, eBARS_DateTime, eBARS_Eod, _
        GridLines , tmrRealTime, UserName, Password, ToolTipText, strUrl
End Sub

' Sets ENABLED prop if not already set to new setting
Public Sub Enable(ctl As Control, _
        Optional ByVal bEnabled As Long = True)
    On Error Resume Next
    If bEnabled = 0 Then
        If ctl.Enabled Then ctl.Enabled = False
    Else
        If Not ctl.Enabled Then ctl.Enabled = True
    End If
End Sub
Public Sub Disable(ctl As Control)
    On Error Resume Next
    If ctl.Enabled Then ctl.Enabled = False
End Sub
' Sets VISIBLE prop if not already set to new setting
Public Sub SetVisible(ctl As Control, _
        Optional ByVal bVisible As Long = True)
    On Error Resume Next
    If bVisible = 0 Then
        ctl.Visible = False
    ElseIf Not ctl.Visible Then
        ctl.Visible = True
    End If
End Sub

' Obsolete (left for backwards-compatibility) -- should use "InfBox"
Public Function AskBox(ByVal strArgs$) As String
    AskBox = InfBox(strArgs)
End Function

' Provides extended functionality for MsgBox and InputBox
' - pass one or more args delimited by " ; "
' - arg format: arg=value  (only first letter of arg is significant)
' - function returns first letter of button selected by user,
'       or returns string if "get=" is used
'  General args (like MsgBox) ...
'    Message=Text message to show in the box (can use pipe "|" delimiters to force new line)
'    Buttons=+Default|-Cancel|Etc  (up to three buttons, delimited by "|", can use "+" prefix for default button and "-" prefix for cancel button)
'    Icon=?  -- options: ?, !, Error(stop sign), Info, Timer(stopwatch), HappyFace, SadFace (only first letter is significant)
'    Header=Title of window
'    Color=White  (background color, defaults to Gray)
'    Timeout=nn  (where nn=number of seconds)
'    Wait=NoWait  (to show non-modally, e.g. while process is going on)
'    Size=nn  (where nn=font size of message,  e.g. 8, 10, etc.)
'    FontBold=MBITD  (where M=Message, B=Buttons, I=InputBox, T=Timeout, D=Day, A=All)
'  Args to get input (like InputBox) ...
'    Get=string   -- options: string, date, number, password (shows asterisks)
'    Default=Default string  (optional)
' Examples ...
'   To just display a message (default is "OK" button):
'       InfBox "icon=inf ; msg=A message for you to see"
'   To ask whether to rename or make a copy:
'       rtrn = InfBox("icon=? ; buttons=+Copy|-Rename ; header=Copy or Rename ; msg=Do you wish to rename this system, or make a new copy?"
'   To show non-modal message while processing:
'       InfBox "i=t ; msg=Initializing ..."
'   To clear non-modal message when done:
'       InfBox ""
'   To get a string from user:
'       rtrn = InfBox("i=? ; get=str ; default=Default string ; msg=Please enter a string ..."
Public Function InfBox(Optional ByVal strMessage$ = "", Optional ByVal strIcon$ = "", _
        Optional ByVal strButtons$ = "", Optional ByVal strTitle$ = "", _
        Optional ByVal bNoWait As Boolean = False, Optional ByVal nTimeout& = 0, _
        Optional ByVal nBackColor& = -1, Optional ByVal nFontSize& = 0, Optional ByVal strFontBold$ = "", _
        Optional ByVal strGetInput$ = "", Optional ByVal strDefaultInput$ = "", _
        Optional ByVal nAlignment As eGDAlignment = eGDAlign_Center, _
        Optional ByVal bShowDontAsk As Boolean = False, _
        Optional ByVal lProgress As Long = 0&, _
        Optional ByVal nLeft& = -1, Optional ByVal nTop& = -1) As String
On Error Resume Next  ' in case modal form active
    
    Dim strArgs$, i&, frm As frmAsk
    Dim bDoEvents As Boolean

    ' Build args string from parms (to be backward compatible, the args could
    ' either be in the passed parms or all in strMessage)
    If Len(strIcon) > 0 Then strArgs = strArgs & " ; I=" & strIcon
    If Len(strButtons) > 0 Then strArgs = strArgs & " ; B=" & strButtons
    If Len(strTitle) > 0 Then strArgs = strArgs & " ; H=" & strTitle
    If bNoWait Then strArgs = strArgs & " ; W=NOWAIT"
    If nTimeout <> 0 Then strArgs = strArgs & " ; T=" & Str(nTimeout)
    If nBackColor >= 0 Then strArgs = strArgs & " ; C=" & Str(nBackColor)
    If nFontSize > 0 Then strArgs = strArgs & " ; S=" & Str(nFontSize)
    If Len(strFontBold) > 0 Then strArgs = strArgs & " ; F=" & strFontBold
    If Len(strGetInput) > 0 Then strArgs = strArgs & " ; G=" & strGetInput
    If Len(strDefaultInput) > 0 Then strArgs = strArgs & " ; D=" & strDefaultInput
    If Len(strArgs) > 0 Then strArgs = strArgs & " ; A=" & Str(nAlignment)
    If bShowDontAsk Then strArgs = strArgs & " ; Z=SHOW"
    If lProgress > 0 Then strArgs = strArgs & " ; P=" & Str(lProgress)
    If Len(strMessage) > 0 Then
        If Len(strArgs) = 0 Then
            ' for backwards-compatibility with old style
            ' (all args passed in as a single string)
            strArgs = strMessage
        Else
            strArgs = strArgs & " ; M=" & strMessage
        End If
    End If

    ' Pass parms to form
    If Left(strArgs, 3) = " ; " Then strArgs = Mid(strArgs, 4)
    strArgs = Trim(strArgs)

    ' Clear existing non-modal form
    If FormIsLoaded("frmAsk") Then
        frmAsk.strArgs = ""
        Unload frmAsk '(ok if not loaded since doing an "on error resume next")
        
        ' TLB 9/6/2013: not even sure we really need a DoEvents when just clearing it?
        'bDoEvents = True
    End If

    ' Show box.
    If InStr(UCase(strArgs), "=NOWAIT") Then
        ' for non-modal, use the original form directly
        frmAsk.strArgs = Trim(strArgs)
        i = frm.chkDontAsk '(this just forces the form to load now if not already loaded)
        If nLeft <> -1 Then frmAsk.Left = nLeft
        If nTop <> -1 Then frmAsk.Top = nTop
        ShowForm frmAsk, eForm_Nonmodal
        bDoEvents = True
    ElseIf Len(Trim(strArgs)) > 0 Then
        ' for modal, use a new instance of the form (so another call
        ' to InfBox while this is showing won't blow this one away)
        If Screen.MousePointer = 11 Then Screen.MousePointer = 0
        Set frm = New frmAsk
        frm.strArgs = Trim(strArgs)
        i = frm.chkDontAsk '(this just forces the form to load now if not already loaded)
        If nLeft <> -1 Then frm.Left = nLeft
        If nTop <> -1 Then frm.Top = nTop
        ShowForm frm, eForm_Modal
        bDoEvents = True
        
        ' Pass output back.
        InfBox = frm.strArgs
        Set frm = Nothing
    End If
    ' TLB 9/6/2013: to eliminate a DoEvents when calling InfBox to clear it (esp. if wasn't even loaded)
    If bDoEvents Then
        DoEvents '(to refresh quickly)
    End If

End Function

' Use this routine to set focus so won't bomb out
' at run-time if control or form not available for focus.
'   10/16/98: only if active window (so won't
'   grab focus from another app, etc.)
Public Sub MoveFocus(ctl As Object)

    Dim whActive%, whParent%
    On Error Resume Next
    If Not ctl Is Nothing Then
        'if GetActiveWindow returns 0, our app is not
        'active in which case we don't want to do anything
        If GetActiveWindow() <> 0 Then
            ctl.SetFocus
        End If
    End If

End Sub

' Selects entire text in a control
' (for auto: call from control's "GotFocus" routine)
Public Sub SelectAll(ctl As Control)

    On Error Resume Next
    Dim i%

    If TypeOf ctl Is ListBox Then
        If ctl.ListCount > 0 Then
            MoveFocus ctl
            SendKeys "{End}+{Home}"
        End If
    Else
        ctl.SelStart = 0
        i = 999 ' in case no ".Text" property
        i = Len(ctl.Text)
        ctl.SelLength = i
    End If

End Sub

' Will return the QBColor code of the color (as a string) passed.
Public Function QbClr(ByVal strColor$) As Long

    Dim nColor&, bLite As Boolean
    strColor = UCase(Trim(strColor))

    ' Check for "Light" as a prefix.
    If Left(strColor, 1) = "+" Then
        bLite = True
        strColor = Mid(strColor, 2)
    Else
        bLite = False
    End If

    ' Determine color.
    Select Case Left(strColor, 3)
        Case "BLA"
            nColor = 0 'black
        Case "BLU"
            nColor = 1 'blue
        Case "GRE"
            nColor = 2 'green
        Case "CYA"
            nColor = 3 'cyan
        Case "RED"
            nColor = 4 'red
        Case "MAG", "PUR"
            nColor = 5 'magenta/purple
        Case "YEL"
            nColor = 6 'yellow
        Case "GRA"
            nColor = 7 'gray
        Case "WHI"
            nColor = 15 'white
            bLite = False
    End Select

    If bLite Then nColor = nColor + 8
    QbClr = QBColor(nColor)
End Function

Public Function AddSlash(ByVal strPath$) As String

    strPath = Trim(strPath)
    If Right(strPath, 1) <> "\" And Len(strPath) > 0 Then
        AddSlash = strPath + "\"
    Else
        AddSlash = strPath
    End If

End Function

Public Function AskPassword(ByVal strGoodPassword$, ByVal strPrompt$) As Boolean
    ' can delimit multiple good passwords with "|"
    Dim strPassword$, bRtrn As Boolean

    If Len(Trim(strGoodPassword)) = 0 Then
        AskPassword = True
        Exit Function
    End If

    bRtrn = False
    If Len(Trim(strPrompt)) = 0 Then strPrompt = "Enter password ..."
    strPassword = AskBox("icon=? ; get=password ; h=Password ; m=" + strPrompt)
    If Len(strPassword) > 0 Then
        strPassword = "|" + Trim(UCase(strPassword)) + "|"
        strGoodPassword = "|" + Trim(UCase(strGoodPassword)) + "|"
        If InStr(strGoodPassword, strPassword) > 0 Then
            bRtrn = True
        Else
            InfBox "i=Err ; Incorrect Password!"
        End If
    End If

    AskPassword = bRtrn
End Function

' Return right-justified number string of specified width.
' If zero width, return number string with one space.
' If dec < 0 then let it default to significant digits
' If negative width and Numb = 0, return spaces.
Public Function NumStr(ByVal dNum#, ByVal iWidth%, Optional ByVal iAfterDec% = 0) As String
    
    Dim iNumSpaces%, strFormat$, strNum$, bNoLeadingZero As Boolean

    If iAfterDec > 100 Then
        bNoLeadingZero = True
        iAfterDec = iAfterDec - 100
    End If
    Select Case iAfterDec
        Case 0
            strFormat = "0"
        Case 1
            strFormat = "0.0"
        Case 2
            strFormat = "0.00"
        Case 3
            strFormat = "0.000"
        Case 4
            strFormat = "0.0000"
        Case 5
            strFormat = "0.00000"
        Case 10 'currency
            strFormat = "$#,##0.00"
        Case Else
            strNum = Trim(Str(dNum))
    End Select
    If iAfterDec >= 0 Then
        If bNoLeadingZero Then
            If Mid(strFormat, 2, 1) = "." Then strFormat = Mid(strFormat, 2)
        End If
        strNum = Format(dNum, strFormat)
    End If

    If iWidth = 0 Then
        NumStr = " " + strNum
    ElseIf iWidth < 0 And dNum = 0 Then
        NumStr = Space(Abs(iWidth))
    Else
        iNumSpaces = Abs(iWidth) - Len(strNum)
        If iNumSpaces > 0 Then
            NumStr = Space(iNumSpaces) & strNum
        Else
            NumStr = strNum
        End If
        'NumStr = Right(Space(Abs(iWid)) + Str(dNumb), Abs(iWid))
    End If

End Function

' calculate number of weekdays between dates
Public Function NumWeekDays(ByVal start_date As Variant, ByVal end_date As Variant) As Long
    
    Dim num_days&

    start_date = Int(start_date)
    end_date = Int(end_date)

    ' move end to Friday if on weekend
    If Weekday(end_date) = 1 Then
        end_date = end_date - 2
    ElseIf Weekday(end_date) = 7 Then
        end_date = end_date - 1
    End If
    If end_date <= start_date Or start_date = 0 Then
        NumWeekDays = 0
        Exit Function
    End If

    ' move start to same day as end
    Do While Weekday(start_date) <> Weekday(end_date)
        If Weekday(start_date) <> 1 And Weekday(start_date) <> 7 Then
            num_days = num_days + 1
        End If
        start_date = start_date + 1
    Loop

    ' now add 5 for each week
    num_days = num_days + (end_date - start_date) / 7 * 5

    NumWeekDays = num_days
End Function

Public Function Pad(strText$, ByVal nWidth&, strJustif$) As String
    Select Case UCase(Left(strJustif, 1))
        Case "R"
            ' Right-justified:
            Pad = Right(Space(nWidth) + strText, nWidth)
        Case "C", "B"
            ' Centered:
            Pad = Left(Space((nWidth - Len(strText)) \ 2) + strText + Space(nWidth / 2 + 2), nWidth)
        Case Else
            ' Left-justified:
            Pad = Left(strText + Space(nWidth), nWidth)
    End Select
End Function

' This function will parse "fields" out of a string.
' e.g. to get the 3rd field of a comma-delimited
'    string:  fld3 = Parse(whole_string, ",", 3)
Public Function Parse(ByVal strToParse$, ByVal strDelim$, ByVal iFldNum%, _
        Optional ByVal bTrimField As Boolean = True) As String

    Dim i&, nStartPos&, nEndPos&, strRtrn$
    Static nSavePos&
    strRtrn$ = ""

    ' If FldNum = 0, just get the next field.
    ' (should first have been called with FldNum > 0)
    If iFldNum = 0 Then
        nStartPos = nSavePos
    Else
        nSavePos = 0
        nStartPos = 1

        ' Find start of FldNum ...
        For i = 1 To iFldNum - 1
            nStartPos = InStr(nStartPos, strToParse, strDelim)
            If nStartPos = 0 Then
                Exit For    ' Not that many fields!
            End If
            nStartPos = nStartPos + Len(strDelim)
            If strDelim = " " Then
                Do While Mid(strToParse, nStartPos, 1) = " "
                    nStartPos = nStartPos + 1
                Loop
            End If
        Next 'i
    End If

    ' Find end of FldNum ...
    If nStartPos > 0 Then
        nEndPos = InStr(nStartPos, strToParse, strDelim)
        If nEndPos = 0 Then
            strRtrn = Mid(strToParse, nStartPos)
            nSavePos = 0
        Else
            strRtrn = Mid(strToParse, nStartPos, nEndPos - nStartPos)
            nSavePos = nEndPos + Len(strDelim)
            If strDelim = " " Then
                Do While Mid(strToParse, nSavePos, 1) = " "
                    nSavePos = nSavePos + 1
                Loop
            End If
        End If
    End If

    If bTrimField Then
        Parse = Trim(strRtrn)
    Else
        Parse = strRtrn
    End If
End Function

' compute coordinates (polar-rect conversion)
Public Function PolarToRect(ByVal dDistance#, ByVal dDegrees#, ByVal bDoY As Boolean) As Double

    If bDoY Then
        PolarToRect = dDistance * Sin(dDegrees * PI / 180#)
    Else
        PolarToRect = dDistance * Cos(dDegrees * PI / 180#)
    End If
    
End Function

' Performs "QuickSort" on a String Array
'   ... commonly accepted as the most efficient sorting algorithm
' (from "Visual Basic Programmer's Journal", May 1995)
Public Sub QuickSortStr(strArray() As String, _
        Optional ByVal bIgnoreCase As Boolean = False, _
        Optional ByVal Bottom As Long = -2000000000, _
        Optional ByVal Top As Long = 2000000000, _
        Optional ByVal bAscending As Boolean = True)

    Dim iLo&, iHi&, iMid&, pivot_val As String, Swap As String

    If Bottom < LBound(strArray) Then Bottom = LBound(strArray)
    If Top > UBound(strArray) Then Top = UBound(strArray)
    iLo = Bottom    ' lower boundary of partition
    iHi = Top       ' upper boundary of partition
    ' arbitrarily select value in middle as "pivot"
    iMid = Bottom + ((Top - Bottom) \ 2)
    If bIgnoreCase Then
        pivot_val = UCase(strArray(iMid))
    Else
        pivot_val = strArray(iMid)
    End If
    ' look until lower and upper cross ...
    Do While iLo <= iHi
        If bIgnoreCase Then
            If bAscending Then
                While UCase(strArray(iLo)) < pivot_val And iLo < Top
                    iLo = iLo + 1   ' look from bottom for next lower > pivot
                Wend
                While UCase(strArray(iHi)) > pivot_val And iHi > Bottom
                    iHi = iHi - 1   ' look from top for next upper < pivot
                Wend
            Else ' if descending:
                While UCase(strArray(iLo)) > pivot_val And iLo < Top
                    iLo = iLo + 1
                Wend
                While UCase(strArray(iHi)) < pivot_val And iHi > Bottom
                    iHi = iHi - 1
                Wend
            End If
        Else
            If bAscending Then
                While strArray(iLo) < pivot_val And iLo < Top
                    iLo = iLo + 1   ' look from bottom for next lower > pivot
                Wend
                While strArray(iHi) > pivot_val And iHi > Bottom
                    iHi = iHi - 1   ' look from top for next upper < pivot
                Wend
            Else ' if descending:
                While strArray(iLo) > pivot_val And iLo < Top
                    iLo = iLo + 1
                Wend
                While strArray(iHi) < pivot_val And iHi > Bottom
                    iHi = iHi - 1
                Wend
            End If
        End If
        If iLo <= iHi Then
            ' swap values for lower and upper
            Swap = strArray(iLo)
            strArray(iLo) = strArray(iHi)
            strArray(iHi) = Swap
            iLo = iLo + 1
            iHi = iHi - 1
        End If
    Loop

    ' recursive calls (for unsorted subpartitions):
    If Bottom < iHi Then Call QuickSortStr(strArray(), bIgnoreCase, Bottom, iHi, bAscending)
    If iLo < Top Then Call QuickSortStr(strArray(), bIgnoreCase, iLo, Top, bAscending)
End Sub

' Performs "QuickSort" on a Variant Array
'   ... commonly accepted as the most efficient sorting algorithm
' (from "Visual Basic Programmer's Journal", May 1995)
Public Sub QuickSortV(vArray() As Variant, _
        Optional ByVal Bottom As Long = -2000000000, _
        Optional ByVal Top As Long = 2000000000, _
        Optional ByVal bAscending As Boolean)

    Dim iLo&, iHi&, iMid&, pivot_val As Variant, Swap As Variant

    If Bottom < LBound(vArray) Then Bottom = LBound(vArray)
    If Top > UBound(vArray) Then Top = UBound(vArray)
    iLo = Bottom    ' lower boundary of partition
    iHi = Top       ' upper boundary of partition
    ' arbitrarily select value in middle as "pivot"
    iMid = Bottom + ((Top - Bottom) \ 2)
    pivot_val = vArray(iMid)
    ' look until lower and upper cross ...
    Do While iLo <= iHi
        If bAscending Then
            While vArray(iLo) < pivot_val And iLo < Top
                iLo = iLo + 1   ' look from bottom for next lower > pivot
            Wend
            While vArray(iHi) > pivot_val And iHi > Bottom
                iHi = iHi - 1   ' look from top for next upper < pivot
            Wend
        Else ' if descending:
            While vArray(iLo) > pivot_val And iLo < Top
                iLo = iLo + 1
            Wend
            While vArray(iHi) < pivot_val And iHi > Bottom
                iHi = iHi - 1
            Wend
        End If
        If iLo <= iHi Then
            ' swap values for lower and upper
            Swap = vArray(iLo)
            vArray(iLo) = vArray(iHi)
            vArray(iHi) = Swap
            iLo = iLo + 1
            iHi = iHi - 1
        End If
    Loop

    ' recursive calls (for unsorted subpartitions):
    If Bottom < iHi Then Call QuickSortV(vArray(), Bottom, iHi, bAscending)
    If iLo < Top Then Call QuickSortV(vArray(), iLo, Top, bAscending)
End Sub

' Performs Binary Search on a SORTED string array
'   ... e.g. needs to only look at 32 items to search list of 4 billion!
' (modified from algorithm found in "Understanding and Using Visual Basic", p. 284)
' - if FOUND: returns TRUE, iPos is position of match
' - if NOT FOUND: returns FALSE, iPos is position to insert
Public Function BinarySearchStr(search_for As String, strArray() As String, Bottom&, Top&, _
        Optional iPos As Long) As Boolean

    Dim iLo&, iHi&, iMid&, Found As Boolean
    iLo = Bottom    ' usually 0
    iHi = Top       ' largest elem# in array
    Found = False
    Do While iLo <= iHi And Not Found
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        If search_for < strArray(iMid) Then
            iHi = iMid - 1  ' in lower half
        ElseIf search_for > strArray(iMid) Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            Found = True
            ' now back up to the very first match
            Do While iMid > Bottom
                If strArray(iMid - 1) <> search_for Then Exit Do
                iMid = iMid - 1
            Loop
        End If
    Loop
    If Not IsMissing(iPos) Then
        iPos = iMid ' position of match or where to insert
    End If

    BinarySearchStr = Found
End Function

' Performs Binary Search on a SORTED variant array
'   ... e.g. needs to only look at 32 items to search list of 4 billion!
' (modified from algorithm found in "Understanding and Using Visual Basic", p. 284)
' - if FOUND: returns TRUE, iPos is position of match
' - if NOT FOUND: returns FALSE, iPos is position to insert
Public Function BinarySearchV(search_for As Variant, vArray() As Variant, Bottom&, Top&, _
        Optional iPos As Long) As Boolean

    Dim iLo&, iHi&, iMid&, Found As Boolean
    iLo = Bottom    ' usually 0
    iHi = Top       ' largest elem# in array
    Found = False
    Do While iLo <= iHi And Not Found
        iMid = iLo + (iHi - iLo) \ 2  ' so sum does not overflow
        If search_for < vArray(iMid) Then
            iHi = iMid - 1  ' in lower half
        ElseIf search_for > vArray(iMid) Then
            iLo = iMid + 1  ' in upper half
            iMid = iLo  ' position to insert
        Else
            Found = True
            ' now back up to the very first match
            Do While iMid > Bottom
                If vArray(iMid - 1) <> search_for Then Exit Do
                iMid = iMid - 1
            Loop
        End If
    Loop
    If Not IsMissing(iPos) Then
        iPos = iMid ' position of match or where to insert
    End If

    BinarySearchV = Found
End Function


' Returns a random number within a given range
Function RandomNum(ByVal nLowerBound&, ByVal nUpperBound&) As Long

    Static bRandomized As Boolean
    
    ' initialize random number generator once
    If Not bRandomized Then
        Randomize
        bRandomized = True
    End If
    
    ' get a random number
    If nUpperBound <= nLowerBound Then
        RandomNum = nLowerBound
    Else
        RandomNum = Int((nUpperBound - nLowerBound + 1) * Rnd) + nLowerBound
    End If
    
End Function

' To properly round a number to "dec" digits after decimal point
Function RoundNum#(ByVal dNum#, Optional ByVal iDecimalDigits% = 0)
    
#If 1 Then
    ' C++ version is much faster (and now avoids any rounding issues)
    RoundNum = gdRoundNum(dNum, iDecimalDigits)
#Else
    ' MUST include CDbl -- otherwise slight rounding errors, don't know why!!
    ' (e.g. passing "7.1235, 3" without using the CDbl causes it to return 7.123)
    Dim dMult#
    If iDecimalDigits = 0 Then
        RoundNum = Int(CDbl(dNum + 0.5))
    ElseIf iDecimalDigits < 0 Then
        dMult = 10# ^ Abs(iDecimalDigits)
        RoundNum = Int(CDbl(dNum / dMult + 0.5)) * dMult
    Else
        dMult = 10# ^ iDecimalDigits
        RoundNum = Int(CDbl(dNum * dMult + 0.5)) / dMult
    End If
#End If

End Function

' Rounds value to specified number of significant digits
Public Function RoundToSigDigits(ByVal dValue#, Optional ByVal iSigDigits% = 9) As Double

#If 1 Then
    ' C++ version is much faster (and now avoids any rounding issues)
    RoundToSigDigits = SigDigits(dValue, iSigDigits)
#Else
    Dim dAbs#, iDigits%
    dAbs = Abs(dValue)
    If dAbs > 1 Then
        For iDigits = 1 To 308
            If dAbs < 10# ^ iDigits Then
                Exit For
            End If
        Next
    ElseIf dAbs > 0 And dAbs < 1 Then
        For iDigits = -1 To -324 Step -1
            If dAbs > 10# ^ iDigits Then
                iDigits = iDigits + 1
                Exit For
            End If
        Next
    End If
    RoundToSigDigits = RoundNum(dValue, iSigDigits - iDigits)
#End If
    
End Function

' Execute a process, and can wait till it finishes
' - to display a web page, pass InternetBrowser as the
'       program and the web page address as the args
Public Function RunProcess(ByVal strProgram As String, _
    Optional ByVal strArgs As String = "", _
    Optional ByVal bWaitTillDone As Boolean = False, _
    Optional ByVal nWindowStyle As VbAppWinStyle = vbNormalFocus, _
    Optional nExitCode As Long = 0, Optional ByVal dwCreationFlags As Long = 0, _
    Optional ByVal strStartingPath As String = "", _
    Optional ByRef nProcessID As Long) As Boolean
    
    Dim bSuccess As Boolean, strPath$, i&
    Dim StartInfo As STARTUPINFO
    Dim ProcInfo As PROCESS_INFORMATION
    Dim strTempBatFile$, strTempDoneFile$, strCmd$

    On Error GoTo RunProcess_Error

    ' specify how to show window
    StartInfo.cb = Len(StartInfo)
    StartInfo.wShowWindow = nWindowStyle
    StartInfo.dwFlags = StartInfo.dwFlags Or STARTF_USESHOWWINDOW
    
    ' get startup path (so process runs from this directory)
    strProgram = Trim(strProgram)
    If Left(strProgram, 2) = ".\" Then strProgram = App.Path & Mid(strProgram, 2)
    i = At(strProgram, "\", -1)
    If i > 0 Then
        strPath = Trim(Left(strProgram, i - 1))
        If Right(strPath, 1) = ":" Then strPath = strPath & "\"
        If Left(strPath, 1) = Chr(34) Then strPath = Mid(strPath, 2)
    End If
    ' invalid if no drive specified (so default to App.Path)
    If InStr(strPath, ":") = 0 Then strPath = App.Path
    
    ' build the command line (it's better for Win2000 if we put it all together)
    If Len(strProgram) > 0 And Left(strProgram, 1) <> Chr(34) Then
        ' safer to put double-quotes around name of program
        strProgram = Chr(34) & strProgram & Chr(34)
    End If
    ' special handling for Batch files (6/28/2004: need to whether waiting or not)
    If UCase(Right(strProgram, 5)) = ".BAT" & Chr(34) Then ' And bWaitTillDone Then
        ' To wait for batch file to finish, we cannot depend on
        ' "GetExitCodeProcess" since the process won't exit until
        ' the user closes out the DOS window (if not set to auto-close),
        ' so instead we'll create a batch file to call this batch file
        ' and when it completes we'll create the ".DON" file.
        ''strTempBatFile = TempPath & "__$temp_.bat"
        strTempBatFile = TempPath & "Please CLOSE When Finished.BAT"
        strTempDoneFile = ReplaceFileExt(strTempBatFile, ".DON")
        KillFile strTempDoneFile
        strCmd = "Call " & Trim(strProgram & " " & strArgs) _
            & vbCrLf & "@copy " & Chr(34) & strTempBatFile & Chr(34) _
            & " " & Chr(34) & strTempDoneFile & Chr(34)
        FileFromString strTempBatFile, strCmd, True, False
        strProgram = strTempBatFile
    Else
        ' see if should try to find program associated with document
        If UCase(Right(strProgram, 5)) <> ".BAT" & Chr(34) _
                And UCase(Right(strProgram, 5)) <> ".EXE" & Chr(34) Then
            strCmd = Space(1024)
            If FindExecutable(StripStr(strProgram, Chr(34)), strPath, strCmd) <= 32 Then
                strCmd = ""
            End If
            FixNullTermStr strCmd
            strCmd = Trim(strCmd)
            If Len(strCmd) > 0 Then
                strProgram = Chr(34) & strCmd & Chr(34) & " " & Trim(strProgram)
            End If
        End If
        ' append args
        strProgram = Trim(strProgram & " " & strArgs)
    End If

    ' run the process
    ' (do NOT inherit handles, since doing so causes this program to not
    '  be able to completely close until the created process is also closed)
    DebugLog strProgram
    dwCreationFlags = dwCreationFlags Or CREATE_NEW_CONSOLE
    
    If Len(strStartingPath) = 0 Then strStartingPath = strPath
    
    If CreateProcess(ByVal 0&, strProgram, ByVal 0&, ByVal 0&, 0&, _
            dwCreationFlags, ByVal 0&, strStartingPath, StartInfo, ProcInfo) <> 0 Then
    
        bSuccess = True
        If bWaitTillDone Then
            ' wait until the process has terminated
            Do
                Sleep 0.1
                nExitCode = 0
                If GetExitCodeProcess(ProcInfo.hProcess, nExitCode) = 0 Then
                    Exit Do ' an error with this function
                End If
                ' special handling for batch files
                If Len(strTempBatFile) > 0 Then
                    If FileExist(strTempDoneFile) Then Exit Do
                End If
            Loop While nExitCode = 259 'STILL_ACTIV
        End If
        
        If Not IsMissing(nProcessID) Then nProcessID = ProcInfo.dwProcessId
        ' must close these handles when we don't need them anymore
        ' (even if process is still running)
        CloseHandle ProcInfo.hProcess
        CloseHandle ProcInfo.hThread
    Else
        DebugLog "CreateProcess FAILED"
    End If
    
RunProcess_Exit:
    If Len(strTempBatFile) > 0 And bWaitTillDone Then
        Sleep 0.2
        KillFile strTempBatFile
        KillFile strTempDoneFile
    End If
    RunProcess = bSuccess
    Exit Function

RunProcess_Error:
    Resume RunProcess_Exit
End Function

' Returns internet browser program (to use with RunProcess)
Public Function InternetBrowser() As String
    
    Dim strFile$
    Static strPgm$
    On Error Resume Next

    ' no need to check again if done within last minute
    If Len(strPgm) = 0 Then
        ' must create a temporary .HTM file so
        ' FindExecutable will be successful
        strFile = TempPath & "_TEMP_.HTM"
        FileFromString strFile, "temporary"
        strPgm = Space(1024)
        If FindExecutable(strFile, "", strPgm) <= 32 Then
            strPgm = ""
        End If
        FixNullTermStr strPgm
        KillFile strFile
    End If
    InternetBrowser = Trim(strPgm)
    
End Function

' To correctly set the value of various controls:
' check box, option button, list box, combo box
Public Sub SetCtl(ctl As Control, ByVal nValue&)
    
    ' first see if this is a "list-type" of control
    On Error GoTo DoesNotHaveAList
    If nValue < ctl.ListCount Then
        ctl.ListIndex = nValue
    End If
    Exit Sub

DoesNotHaveAList:
    On Error Resume Next
    If nValue = 0 Then
        ctl = 0
    ElseIf TypeOf ctl Is CheckBox Or TypeOf ctl Is ctlUniCheckXP Then 'RH added ctlUniCheckXP
        ctl = 1
    ' Sheridan controls no longer being used
    'ElseIf TypeOf ctl Is SSCheck Then
        'ctl = 1
    ElseIf TypeOf ctl Is OptionButton Then
        ctl = True
    'ElseIf TypeOf ctl Is SSOption Then
        'ctl = True
    End If

End Sub

' To make window stay on top (float), or not
Sub SetFormTopmost(frm As Form, ByVal OnTop As Boolean)

    If OnTop Then
        SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

End Sub

' shuffles (randomizes) elements of array
Public Sub Shuffle(vArray() As Variant, ByVal Min&, ByVal Max&)

    Dim i&, irnd&, Swap As Variant
    For i = Min To Max
        irnd = RandomNum(Min, Max)
        Swap = vArray(i)
        vArray(i) = vArray(irnd)
        vArray(irnd) = Swap
    Next

End Sub

' Goes into an idle state for number of seconds
' - dSeconds >= 0: fully idles for specified # of seconds
'    (0 = quickest full idle which yields timeslice to other threads)
' - dSeconds < 0: if a full idle was done within specified
'    # of seconds, then does a quicker "in-thread idle"
'    which does not yield a timeslice (like "DoEvents")
' - set bSkipFullIdle = True to skip the "WaitMessage" in all cases
' - set bEvinIfUnloaded = True to sleep even if Trade Nav is currently unloading
Public Sub Sleep(Optional ByVal dSeconds As Double = 0, Optional ByVal bSkipFullIdle As Boolean = False, Optional ByVal bEvenIfUnloading As Boolean = False)

#If TRADENAV_EXE Then
    ' for a long sleep in TradeNav, need to check intermittantly for g.bUnloading
    If bEvenIfUnloading = False Then
        Dim dWaitUntil#
        If dSeconds >= 0.25 Then
            dWaitUntil = gdTickCount + dSeconds * 1000#
            Do
                If g.bUnloading Then Exit Sub
                IdleSleep 125, bSkipFullIdle
            Loop While gdTickCount < dWaitUntil
            Exit Sub
        End If
    End If
#End If
    
    IdleSleep CLng(dSeconds * 1000#), bSkipFullIdle

End Sub

' Strips a string of all characters in second string.
Public Function StripStr(ByVal strToStrip$, ByVal strStripThese$) As String

    Dim i&, iPos&, strRtrn$, strChar$, iNewLen&
    
    iNewLen = 0
    strRtrn = Space(Len(strToStrip))
    For i = 1 To Len(strToStrip)
        strChar = Mid(strToStrip, i, 1)
        If InStr(strStripThese, strChar) = 0 Then
            iNewLen = iNewLen + 1
            Mid(strRtrn, iNewLen, 1) = strChar
        End If
    Next

    StripStr = Left(strRtrn, iNewLen)
End Function

' To validate user keystrokes into a text box, call this
' from the "KeyUp" event.
'  ctl:  name of control
'  char_filter:  mask of allowable characters -- a "0"
'      will allow all digits, an "A" will allow all alpha
'      (use just "0" for numeric, "-0" to include negatives)
'  str_length:  max length of text box (0 = unlimited)
'  max_decimals:  if numeric, max digits after the decimal
'
' e.g.  UserText Text1, "-0", 7, 2  '// allows -999.99
Sub UserText(ctl As Control, ByVal strCharFilter$, ByVal nStrLength&, ByVal nMaxDecimals&)

    Dim strOrig$, strText$, i&, c$, nCursor&
    
    ' build character filter
    If Len(strCharFilter) > 0 Then
        strCharFilter = UCase(strCharFilter)
        If nMaxDecimals > 0 Then strCharFilter = strCharFilter + "."
        If InStr(strCharFilter, "0") Then strCharFilter = strCharFilter + "123456789"
        If InStr(strCharFilter, "A") Then strCharFilter = strCharFilter + "BCDEFGHIJKLMNOPQRSTUVWXYZ"
    End If

    ' save original text and cursor location
    strOrig = ctl
    nCursor = ctl.SelStart
    
    ' check each character of existing string
    strText = ""
    For i = 1 To Len(strOrig)
        c = Mid(strOrig, i, 1)
        ' make sure this character may be included
        If Len(strCharFilter) = 0 Or InStr(strCharFilter, UCase(c)) > 0 Then
            strText = strText + c
        Else
            nCursor = nCursor - 1 ' to bump cursor back
        End If
    Next
    
    ' check number of digits after the decimal
    If nMaxDecimals > 0 Then
        i = InStr(strText, ".")
        If i > 0 Then strText = Left(strText, i + nMaxDecimals)
    End If
    
    ' check total string length
    If nStrLength > 0 Then strText = Left(strText, nStrLength)
    
    ' replace if text is different than original
    If strText <> strOrig Then
        Beep ' to let user know something was invalid
        ctl = strText
        ' set cursor back to where it was
        If nCursor < 0 Then nCursor = 0
        ctl.SelStart = nCursor
    End If

End Sub

' returns the week number (# weeks from 12/31/1899)
Public Function WkNum%(ByVal JDate&)

    If JDate > 99999 Then JDate = JulFromLong(JDate)
    WkNum = (JDate - 1) \ 7

End Function

' Returns the user's "WINDOWS" directory.
Public Function WindowsPath(Optional ByVal bStripLastSlash As Boolean = False) As String

    Dim iStrLen%
    Static strPath$
    If strPath = "" Then
        strPath = String(255, 0)
        iStrLen = GetWindowsDirectory(strPath, Len(strPath))
        strPath = AddSlash(Left(strPath, iStrLen))
    End If
    If bStripLastSlash And Right(strPath, 1) = "\" Then
        WindowsPath = Left(strPath, Len(strPath) - 1)
    Else
        WindowsPath = strPath
    End If
End Function

' Returns the user's "SYSTEM" directory.
Public Function WinSysPath(Optional ByVal bStripLastSlash As Boolean = False) As String

    Dim iStrLen%
    Static strPath$
    If strPath = "" Then
        strPath = String(255, 0)
        iStrLen = GetSystemDirectory(strPath, Len(strPath))
        strPath = AddSlash(Left(strPath, iStrLen))
    End If
    If bStripLastSlash And Right(strPath, 1) = "\" Then
        WinSysPath = Left(strPath, Len(strPath) - 1)
    Else
        WinSysPath = strPath
    End If
End Function

' Returns the user's "TEMP" directory.
Public Function TempPath(Optional ByVal bStripLastSlash As Boolean = False) As String

    Dim iStrLen%
    Static strPath$
    If strPath = "" Then
        strPath = String(255, 0)
        iStrLen = GetTempPath(Len(strPath), strPath)
        strPath = AddSlash(Left(strPath, iStrLen))
    End If
    If bStripLastSlash And Right(strPath, 1) = "\" Then
        TempPath = Left(strPath, Len(strPath) - 1)
    Else
        TempPath = strPath
    End If
End Function

' Returns the path for one of the special folders.
Public Function SpecialFolderPath(ByVal iFolderID As CSIDL_FOLDERS, Optional ByVal bStripLastSlash As Boolean = False) As String

    On Error Resume Next ' since SHGetFolderPath is not fully supported on Win95, Win98 or ME
    Dim strPath$, i&, bQuickLaunch As Boolean
    
    ' since the older OS's don't really support the "All Users" folders,
    ' just translate them to the "User" folders which do exist
    If Is9598orMe Then
        Select Case iFolderID
        Case CSIDL_ALLUSERS_STARTMENU
            iFolderID = CSIDL_USER_STARTMENU
        Case CSIDL_ALLUSERS_STARTUP
            iFolderID = CSIDL_USER_STARTUP
        Case CSIDL_ALLUSERS_PROGRAMS
            iFolderID = CSIDL_USER_PROGRAMS
        Case CSIDL_ALLUSERS_DESKTOP
            iFolderID = CSIDL_USER_DESKTOP
        End Select
    End If
    
    ' the "Quick Launch" folder is simply built from the USER_APPDATA
    If iFolderID = CSIDL_QUICKLAUNCH Then
        iFolderID = CSIDL_USER_APPDATA
        bQuickLaunch = True
    End If
    
    strPath = String(255, 0)
    i = -999
    i = SHGetFolderPath(0, iFolderID, 0, 0, strPath)
    If i = 0 Then
        FixNullTermStr strPath
        strPath = AddSlash(Trim(strPath))
        If bQuickLaunch Then
            strPath = strPath & "Microsoft\Internet Explorer\Quick Launch\"
            If Not DirExist(strPath) Then
                strPath = ""
            End If
        End If
        If bStripLastSlash And Right(strPath, 1) = "\" Then
            strPath = Left(strPath, Len(strPath) - 1)
        End If
    Else
        strPath = ""
    End If
    SpecialFolderPath = strPath
    
End Function

' Returns the user's "COMMON FILES" directory.
Public Function CommonFilesPath(Optional ByVal bStripLastSlash As Boolean = False) As String

    Static strPath$
    If strPath = "" Then
        strPath = SpecialFolderPath(CSIDL_PROGRAMFILES_COMMONFILES, bStripLastSlash)
        If Len(strPath) = 0 Then
            strPath = Left(WindowsPath, 2) & "\Program Files\Common Files\"
        End If
    End If
    If bStripLastSlash And Right(strPath, 1) = "\" Then
        CommonFilesPath = Left(strPath, Len(strPath) - 1)
    Else
        CommonFilesPath = strPath
    End If
End Function

' DAJ 03/04/2013: We believe that this is only used in Parse, so we are moving into Parse so that
' we can make modifications to it.  We are going to leave it here just in case something else is
' using it...
#If 0 Then
' checks if 2 floating point numbers are "equivalent"
' (if within four decimal places or .005% of each other).
Function Equiv(ByVal num1 As Double, ByVal num2 As Double) As Boolean

    Dim rtrn As Boolean

    rtrn = True
    ' first check difference (to 4 decimal places)
    If Abs(num1 - num2) >= 0.000051 Then
        ' override if "whole" parts not equal
        ' (e.g. to make sure 64001 <> 64002)
        If Int(num1 + 0.5) <> Int(num2 + 0.5) Or num2 = 0 Then
            rtrn = False
        Else
            ' then check ratio (within .005%)
            If Abs(num1 / num2 - 1) > 0.00005 Then rtrn = False
        End If
    End If

    Equiv = rtrn
End Function
#End If

' usually called from LostFocus event of a text box
' to check for a valid date
Public Function DateCheck(txtDate As Control, ByVal bAllowEmpty As Boolean) As Boolean

    Dim strText$, strTemp$, nYear&, nMonth&, nDay&, nDate&, v As Variant
    Dim bRtrn As Boolean
    Static bInProgress As Boolean
    
    On Error Resume Next

    If bInProgress Then
        ' when an invalid date, the "InfBox" will cause the
        ' LoseFocus event to fire again which will then
        ' reenter this function while it's "in_progress",
        ' in which case we just want to simply return false
        ' so any code expecting a valid date will not run.
        bRtrn = False
    Else
        bInProgress = True
        strText = txtDate
        bRtrn = True
        If Len(Trim(strText)) > 0 Or bAllowEmpty = False Then
            If Not IsDate(strText) Then
                MoveFocus txtDate
                InfBox "i=[Invalid] ; Not a valid date:|" + strText
                'MoveFocus txtDate
                bRtrn = False
            Else
                ' see if the century was included
                v = CVDate(strText)
                nMonth = Month(v)
                nDay = Day(v)
                nYear = Year(v)
                If InStr(strText, Str(nYear)) = 0 Then
                    ' since it wasn't explicity specified,
                    ' let's set the century to within the
                    ' best "window": -75 to +25 years
                    Do While nYear < Year(Date) - 75
                        nYear = nYear + 100
                    Loop
                End If
                nDate = nYear * 10000# + nMonth * 100 + nDay
                strTemp = DateFormat(nDate)
                If strTemp <> strText Then
                    txtDate = strTemp
                End If
            End If
        End If
        bInProgress = False
    End If

    DateCheck = bRtrn
End Function

' Formats a date with the specified number of digits while
' maintaining regional settings (order of parts and delimiter)
' - can pass d1 = "Format" to return the format itself
Public Function DateFormat(dDate As Variant, Optional ByVal eFormat As eDateFormat = MM_DD_YYYY, _
        Optional ByVal eTime As eTimeFormat = NO_TIME, Optional ByVal eAmPm As eAmPmFormat = NO_AMPM, _
        Optional ByVal bDontShowTimeIfZero As Boolean = False) As String

    Dim iPos%, iPos2%, strFmt$, bUseEnglishMonth As Boolean, d As Variant
    Static nChk&, strDates$(5), bAmPm As Boolean

    ' first time: store date formats for all modes
    If nChk = 0 Then
        ' get system format of known date: 12/25/80  'M/d/yy
        nChk = 29580
        strFmt = Format(nChk, "Short Date")

        ' parse it out
        iPos = InStr(strFmt, "12")
        If iPos = 0 Then
            ' look for month name (alpha)
            For iPos = 1 To Len(strFmt)
                If IsAlpha(strFmt, iPos) Then
                    For iPos2 = iPos To Len(strFmt)
                        If Not IsAlpha(strFmt, iPos2) Then
                            strFmt = Left(strFmt, iPos - 1) & "12" & Mid(strFmt, iPos2)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
            iPos = InStr(strFmt, "12")
        End If
        If iPos > 0 Then
            strFmt = Left(strFmt, iPos - 1) + "MM" + Mid(strFmt, iPos + 2)
            iPos = InStr(strFmt, "25")
            If iPos > 0 Then
                strFmt = Left(strFmt, iPos - 1) + "dd" + Mid(strFmt, iPos + 2)
                iPos = InStr(strFmt, "1980")
                If iPos > 0 Then
                    strFmt = Left(strFmt, iPos - 1) + "yyyy" + Mid(strFmt, iPos + 4)
                    nChk = -1 'success flag
                Else
                    iPos = InStr(strFmt, "80")
                    If iPos > 0 Then
                        strFmt = Left(strFmt, iPos - 1) + "yyyy" + Mid(strFmt, iPos + 2)
                        nChk = -1 'success flag
                    End If
                End If
            End If
        End If
        If nChk <> -1 Then strFmt = "MM/dd/yyyy" ' default for unknown!
        
        ' save date format for each mode
        strDates(0) = strFmt
        
        iPos = InStr(strFmt, "yyyy")
        strFmt = Left(strFmt, iPos - 1) + "yy" + Mid(strFmt, iPos + 4)
        strDates(1) = strFmt
        
        iPos = InStr(strFmt, "MM")
        strFmt = Left(strFmt, iPos - 1) + "M" + Mid(strFmt, iPos + 2)
        iPos = InStr(strFmt, "dd")
        strFmt = Left(strFmt, iPos - 1) + "d" + Mid(strFmt, iPos + 2)
        strDates(2) = strFmt
        
        iPos = InStr(strFmt, "yy")
        If iPos = 1 Then
            strFmt = Mid(strFmt, 3)
            Do
                If InStr("MD", UCase(Left(strFmt, 1))) Then Exit Do
                strFmt = Mid(strFmt, 2)
            Loop While Len(strFmt) > 0
        Else
            strFmt = Left(strFmt, iPos - 1)
            Do
                If InStr("MD", UCase(Right(strFmt, 1))) Then Exit Do
                strFmt = Left(strFmt, Len(strFmt) - 1)
            Loop While Len(strFmt) > 0
        End If
        strDates(3) = strFmt
        
        strDates(4) = "MMM-YY"
        
        ' see if AM/PM is used
        If Not IsDBCS Then
            strFmt = Trim(FormatDateTime(Date + 0.75, vbLongTime))
            If Not IsDigit(Right(strFmt, 1)) Then
                bAmPm = True
            End If
        End If
    End If

    ' get date from what was passed in
    d = DateOf(dDate)

    ' build time format
    Select Case eTime
    Case HH_MM_SS
        strFmt = "HH:MM:SS"
    Case HH_MM
        strFmt = "HH:MM"
    Case H_MM_SS
        strFmt = "H:MM:SS"
    Case H_MM
        strFmt = "H:MM"
    Case Else
        strFmt = ""
    End Select
    If Len(strFmt) > 0 Then
        If d > 0 And bDontShowTimeIfZero And d = Int(d) Then
            strFmt = ""
        ElseIf bAmPm Then
            Select Case eAmPm
            Case AP_UPPER
                strFmt = strFmt & "A/P"
            Case AP_LOWER
                strFmt = strFmt & "a/p"
            Case AMPM_UPPER
                strFmt = strFmt & " AM/PM"
            Case AMPM_LOWER
                strFmt = strFmt & " am/pm"
            End Select
        End If
    End If
    
    ' append date format
    If eFormat > UBound(strDates) Then
        eFormat = 0
    End If
    If eFormat >= 0 Then
        'strFmt = Trim(Format(d, strDates(eFormat)) & " " & strFmt)
        strFmt = Trim(strDates(eFormat) & " " & strFmt)
    End If
    
    ' apply format to date
    If d > 0 Then
        If IsDBCS Then
            If InStr(strFmt, "MMM") > 0 Then
                bUseEnglishMonth = True
                strFmt = Replace(strFmt, "MMM", "XXX")
            End If
        End If
        strFmt = Format(d, strFmt)
        If bUseEnglishMonth Then
            strFmt = Replace(strFmt, "XXX", MonthName(Month(d), True))
        End If
    ElseIf VarType(dDate) <> vbString Then
        strFmt = ""
    ElseIf UCase(dDate) <> "FORMAT" Then
        strFmt = ""
    End If

    DateFormat = strFmt
End Function

Public Function DateAndTime(vDate As Variant) As String
    
    DateAndTime = DateFormat(vDate, MM_DD_YYYY, HH_MM, AMPM_LOWER)
    
End Function

Public Function DateOf(vDate As Variant) As Variant
    
    Dim rdate As Variant, dDate#
    
    rdate = Date ' just to set variant to DATE type
    
    If VarType(vDate) = vbDate Then
        rdate = vDate
    ElseIf VarType(vDate) = vbString Then
        If IsDate(vDate) Then
            rdate = CVDate(vDate)
        ElseIf IsNumeric(vDate) Then
            dDate = Val(vDate)
            If dDate > 200000 Then
                rdate = JulFromLong(dDate)
            Else
                rdate = dDate
            End If
        Else
            rdate = 0
        End If
    'ElseIf vDate <= 0 Then
    '    rdate = 0
    ElseIf vDate > 999999 Then
        ' from Long # (YYYYMMDD)
        rdate = JulFromLong(CLng(vDate))
    Else
        rdate = vDate ' julian
    End If

    DateOf = rdate

End Function

' returns Day of Week: 1=Sun, 2=Mon, ..., 6=Fri, 7=Sat
' (can't always trust "FirstDayOfWeek" arg of "Weekday" function)
Public Function Weekday(vDate As Variant) As VbDayOfWeek

    'Note: don't use CLng() since it rounds!
    Dim JDate As Long
    JDate = Int(DateOf(vDate))
    Weekday = ((JDate - 1) Mod 7) + 1

End Function

' returns Name for the Weekday according to regional settings
' (can't always trust "FirstDayOfWeek" arg of "WeekdayName" function)
Public Function WeekdayName(vDate As Variant, Optional ByVal bAbbrev As Boolean = True) As String

    Dim s$, d#

    d = DateOf(vDate)
    If d > 0 Then
        If IsDBCS Then
            ' if DBCS then just use English name (Asian characters don't really work)
            Select Case Weekday(d)
            Case vbSunday: s = "Sunday"
            Case vbMonday: s = "Monday"
            Case vbTuesday: s = "Tuesday"
            Case vbWednesday: s = "Wednesday"
            Case vbThursday: s = "Thursday"
            Case vbFriday: s = "Friday"
            Case vbSaturday: s = "Saturday"
            End Select
            If bAbbrev Then s = Left(s, 3)
        Else
            s = VBA.WeekdayName(Weekday(d), bAbbrev, vbSunday)
        End If
    End If
    
    WeekdayName = s
End Function

' returns Name for the Month according to regional settings
Public Function MonthName(ByVal iMonth As Integer, Optional ByVal bAbbrev As Boolean = True, _
                Optional ByVal bForceEnglish As Boolean = False) As String

    Dim s$

    If iMonth >= 1 And iMonth <= 12 Then
        If IsDBCS Or bForceEnglish Then
            ' if DBCS then just use English name (Asian characters don't really work)
            Select Case iMonth
            Case 1: s = "January"
            Case 2: s = "February"
            Case 3: s = "March"
            Case 4: s = "April"
            Case 5: s = "May"
            Case 6: s = "June"
            Case 7: s = "July"
            Case 8: s = "August"
            Case 9: s = "September"
            Case 10: s = "October"
            Case 11: s = "November"
            Case 12: s = "December"
            End Select
            If bAbbrev Then s = Left(s, 3)
        Else
            s = VBA.MonthName(iMonth, bAbbrev)
        End If
    End If
    
    MonthName = s
End Function

' returns Month (1-12, 0=unknown) from a month name
Public Function MonthNumber(ByVal strMonth As String) As Integer

    Dim i&, s$
    ' first check English/Spanish/German/French/Italian month names
    strMonth = UCase(strMonth)
    Select Case Left(strMonth, 3)
        Case "JAN", "ENE", "GEN"
            MonthNumber = 1
        Case "FEB", "FEV"
            MonthNumber = 2
        Case "MAR"
            MonthNumber = 3
        Case "APR", "ABR", "AVR"
            MonthNumber = 4
        Case "MAY", "MAI", "MAG"
            MonthNumber = 5
        Case "JUN", "GIU"
            MonthNumber = 6
        Case "JUL", "LUG"
            MonthNumber = 7
        Case "JUI" ' first 3 letters of french for Jun & Jul are the same
            If InStr(strMonth, "N") > 0 Then
                MonthNumber = 6
            ElseIf InStr(strMonth, "L") > 0 Then
                MonthNumber = 7
            End If
        Case "AUG", "AGO", "AG"
            MonthNumber = 8
        Case "SEP", "SET"
            MonthNumber = 9
        Case "OCT", "OKT", "OUT", "OTT"
            MonthNumber = 10
        Case "NOV"
            MonthNumber = 11
        Case "DEC", "DIC", "DEZ"
            MonthNumber = 12
        Case Else
            ' if not a match, then also check each abbreviated regional month name
            For i = 1 To 12
                s = UCase(VBA.MonthName(i, True))
                If s = Left(strMonth, Len(s)) Then
                    MonthNumber = i
                    Exit For
                End If
            Next
    End Select

End Function

' returns True if Monday-Friday,  False if a weekend
Public Function IsWeekday(vDate As Variant) As Boolean

    Dim iWkDay As Integer
    iWkDay = Weekday(vDate)
    If iWkDay = 1 Or iWkDay = 7 Then
        IsWeekday = False
    Else
        IsWeekday = True
    End If

End Function

' returns sum of filesizes for files
Public Function DirSize(ByVal strPath$, ByVal strFileMasks$) As Double

    Dim dTotal#, strFileName$
    
    On Error Resume Next
    dTotal = 0
    strPath = AddSlash(strPath)
    If Len(strFileMasks) = 0 Then strFileMasks = "*.*"
    strFileName = Dir(strPath + strFileMasks)
    Do While Len(Trim(strFileName)) > 0
        dTotal = dTotal + FileLength(strPath + strFileName)
        strFileName = Dir
    Loop
    DirSize = dTotal
    
End Function

' To paint a "faded" screen (like setup programs)
Public Sub FadeForm(frm As Form, Optional ByVal strColor$ = "B")

    On Error Resume Next

    Dim iSaveScale%, iSaveStyle%, iSaveRedraw%
    Dim i&, j&, X&, Y&, nPixels&, iRed%, iGreen%, iBlue%

    ' Save current settings.
    iSaveScale = frm.ScaleMode
    iSaveStyle = frm.DrawStyle
    iSaveRedraw = frm.AutoRedraw

    ' Determine color.
    iRed = 0
    iGreen = 0
    iBlue = 0
    Select Case UCase(Left(Trim(strColor), 1))
    Case "R" 'Red
        iRed = 1
    Case "G" 'Green
        iGreen = 1
    Case "C" 'Cyan
        iGreen = 1
        iBlue = 1
    Case "P" 'Purple
        iRed = 1
        iBlue = 1
    Case "Y" 'Yellow
        iRed = 1
        iGreen = 1
    Case "W" 'White
        iRed = 1
        iBlue = 1
        iGreen = 1
    Case Else 'Blue
        iBlue = 1
    End Select

    ' Paint screen.
    frm.ScaleMode = 3
    nPixels = Screen.Height \ Screen.TwipsPerPixelY
    X = nPixels \ 64# + 0.5
    frm.DrawStyle = 5
    frm.AutoRedraw = True
    For j = 0 To nPixels Step X
        Y = 240 - 245 * j \ nPixels
        If Y < 0 Then Y = 0
        frm.Line (-2, j - 2)-(Screen.Width + 2, j + X + 3), RGB(iRed * Y, iGreen * Y, iBlue * Y), BF
    Next 'j

    ' Reset settings.
    frm.ScaleMode = iSaveScale
    frm.DrawStyle = iSaveStyle
    frm.AutoRedraw = iSaveRedraw

End Sub

Public Function FileDate(ByVal strFileName$) As Variant

    Dim rtrn As Variant

    rtrn = 0
    On Error Resume Next
    strFileName = Trim(strFileName)
    FixNullTermStr strFileName
    If FileExist(strFileName) Then
        rtrn = CVDate(FileDateTime(strFileName))
    End If

    FileDate = rtrn
End Function

Public Function FileFromArray(ByVal strFileName$, strArray$()) As Boolean

    Dim fh%, i&

    On Error GoTo FileFromArrayError

    fh = FreeFile
    Open strFileName For Output As #fh
    For i = 1 To UBound(strArray)
        Print #fh, strArray(i)
    Next
    Close #fh

    FileFromArray = True
    Exit Function

FileFromArrayError:
    FileFromArray = False
    Exit Function
End Function

Public Sub FileFromString(ByVal strFileName$, strText$, Optional ByVal bAddCrLf As Boolean = False, _
        Optional ByVal bAppend As Boolean = False, Optional ByVal bBinaryFile As Boolean = False)

    Dim fh&

    On Error Resume Next
    If Not bBinaryFile Then
        fh = FreeFile
        If bAppend Then
            Open strFileName For Append As #fh
        Else
            KillFile strFileName
            Open strFileName For Output As #fh
        End If
        If bAddCrLf Then
            Print #fh, strText
        Else
            Print #fh, strText;
        End If
        Close #fh
    Else
        If bAppend Then
            fh = FileOpen(strFileName, "a+b")
        Else
            fh = FileOpen(strFileName, "w+b")
        End If
        If fh <> 0 Then
            If FileBinaryIO(fh, ByVal strText, Len(strText), True) <> 0 Then
                'ToFile = True
            End If
            FileClose fh
        End If
    End If

End Sub

' Returns length of file, or -1 if file does not exist
Public Function FileLength(ByVal strFileName$) As Double

    On Error Resume Next
    Dim dLength#
    dLength = -1
    dLength = FileLen(strFileName)
    If dLength < -1 Then
        dLength = dLength + (2# ^ 32)
    End If
    FileLength = dLength
    
End Function


Public Function FileToArray(ByVal strFileName$, strArray$()) As Long

    Dim fh%, i&, strTemp$, nCount&, nSize&

    On Error GoTo FileToArrayError

    nSize = 255
    nCount = 0
    If FileExist(strFileName) Then
        fh = FreeFile
        Open strFileName For Input As #fh
        Do While Not EOF(fh)
            Line Input #fh, strTemp
            nCount = nCount + 1
            If nCount > nSize Or nCount = 1 Then
                'If nSize >= 16384 Then GoTo FileToArrayError
                nSize = nSize * 2
                ReDim Preserve strArray$(nSize)
            End If
            strArray(nCount) = strTemp
        Loop
        Close #fh
        ' back out of last empty strings
        For i = nCount To 1 Step -1
            If Len(Trim(strArray(i))) > 0 Then Exit For
            nCount = i - 1
        Next
    End If
    
FileToArrayExit:
    ReDim Preserve strArray$(nCount)
    FileToArray = nCount
    Exit Function

FileToArrayError:
    ' too many items!
    nCount = 0
    Resume FileToArrayExit:
End Function

Public Function FileToString(ByVal strFileName$, Optional ByVal nMaxBytes& = -1, _
        Optional ByVal bFirstLineOnly As Boolean = False, _
        Optional ByVal bBinaryFile As Boolean = False) As String
    
    Dim fh&, nBytes&, strText1$, strText2$
    On Error Resume Next
    nBytes = FileLength(strFileName)
    If nBytes > 0 Then
        If nBytes > nMaxBytes And nMaxBytes >= 0 Then nBytes = nMaxBytes
        On Error Resume Next
        If Not bBinaryFile Then
            ' read Text data
            fh = FreeFile
            Open strFileName For Input As #fh
            ' 7/1/98: first read all except last character
            ' just in case last character is EOF/asc(26).
            If nBytes > 1 Then strText1 = Input$(nBytes - 1, fh)
            ' now try to read last character
            strText2 = Input$(1, fh)
            Close #fh
            FileToString = strText1 & strText2
            If bFirstLineOnly Then
                nBytes = InStr(FileToString, Chr(13))
                If nBytes > 0 Then
                    FileToString = Left(FileToString, nBytes - 1)
                End If
                nBytes = InStr(FileToString, Chr(10))
                If nBytes > 0 Then
                    FileToString = Left(FileToString, nBytes - 1)
                End If
            End If
        Else
            ' read Binary data
            fh = FileOpen(strFileName, "rb")
            If fh <> 0 Then
                strText1 = Space(nBytes + 1)
                If FileBinaryIO(fh, ByVal strText1, nBytes, False) <> 0 Then
                    FileToString = Left(strText1, nBytes)
                End If
                FileClose fh
            End If
        End If
    End If

End Function

Sub FixNullTermStr(strText$)

    Dim i&
    i = InStr(strText, Chr(0))
    If i > 0 Then strText = Left(strText, i - 1)

End Sub

' Works like InStr, but can any occurance (1st, 2nd, 3rd, etc.)
' - set occurance to negative to search from end of string (e.g. -2 means 2nd occurance from right)
' - returns position of where found, or 0 if not found
Public Function At(strSrch$, strLookFor$, ByVal nOccurNum&) As Long

    Dim nStartPos&, nEndPos&, nStepDir%, i&, nOccurance&

    ' Set defaults.
    If nOccurNum > 0 Then
        nStartPos = 1
        nEndPos = Len(strSrch)
        nStepDir = 1
    Else
        nStartPos = Len(strSrch)
        nEndPos = 1
        nStepDir = -1
        nOccurNum = -nOccurNum
    End If
    nOccurance = 0

    ' Find occurance of LookFor$.
    For i = nStartPos To nEndPos Step nStepDir
        If Mid(strSrch$, i, Len(strLookFor)) = strLookFor Then
            nOccurance = nOccurance + 1
            If nOccurance = nOccurNum Then
                At = i
                Exit Function
            End If
        End If
    Next 'i

    ' Not found.
    At = 0
End Function

'Tracks elapsed time between calls.
'- to start benchmarking, call BenchMark with no parm
'- thereafter, call BenchMark with description of process just timed
'   (or if not want to display, pass no parm and use returned value)
Function BenchMark(Optional ByVal strDesc$ = "") As Long

    Static StartTime#, ElapsedTime#

    ElapsedTime = GetTickCount() - StartTime
    If ElapsedTime < 0 Then
        'TickCount wrapped around (every 49 days)
        ElapsedTime = ElapsedTime + 2# ^ 32
    End If
    If Len(strDesc) > 0 Then
        strDesc = Format(ElapsedTime / 1000#, "#####.###") + " seconds for:|" + strDesc
        InfBox "i=t ; h=BenchMark ; " + strDesc
    End If
    StartTime = GetTickCount()
    'return # milliseconds
    If ElapsedTime >= 2147483648# Then ElapsedTime = -1 '(invalid for returning)
    BenchMark = CLng(ElapsedTime)
End Function

' To center a control (horiz, vert, or both) within either a form or another control
' - ThisControl must be a control
' - InThisFormOrControl can be either another control or a form or the Screen object
' - strMode can be either "H" (horiz) or "V" (vert) or "B" (both)
Public Sub CenterTheControl(ThisControl As Control, InThisFormOrControl As Object, Optional ByVal strMode$ = "B")

    Dim iSavScaleMode%

    On Error Resume Next
    strMode = UCase(Trim(strMode))
    If InStr(strMode, "B") > 0 Or Len(strMode) = 0 Then
        strMode = "HV" & strMode
    End If
    If (TypeOf InThisFormOrControl Is Screen) Or InStr(strMode, "S") > 0 Then
        ' center in screen
        If InStr(strMode, "V") > 0 Then ThisControl.Top = (Screen.Height - ThisControl.Height) \ 2
        If InStr(strMode, "H") > 0 Then ThisControl.Left = (Screen.Width - ThisControl.Width) \ 2
    ElseIf TypeOf InThisFormOrControl Is Form Then
        ' center in form (must temporarily set the form's ScaleMode to twips)
        iSavScaleMode = InThisFormOrControl.ScaleMode
        InThisFormOrControl.ScaleMode = 1 ' (twips)
        If InStr(strMode, "V") > 0 Then ThisControl.Top = (InThisFormOrControl.ScaleHeight - ThisControl.Height) \ 2
        If InStr(strMode, "H") > 0 Then ThisControl.Left = (InThisFormOrControl.ScaleWidth - ThisControl.Width) \ 2
        InThisFormOrControl.ScaleMode = iSavScaleMode
    ElseIf ThisControl.Container Is InThisFormOrControl Then
        ' center in the control that contains it
        If InStr(strMode, "V") > 0 Then ThisControl.Top = (InThisFormOrControl.Height - ThisControl.Height) \ 2
        If InStr(strMode, "H") > 0 Then ThisControl.Left = (InThisFormOrControl.Width - ThisControl.Width) \ 2
    Else
        ' center relative to a control which is not the container
        If InStr(strMode, "V") > 0 Then ThisControl.Top = InThisFormOrControl.Top + (InThisFormOrControl.Height - ThisControl.Height) \ 2
        If InStr(strMode, "H") > 0 Then ThisControl.Left = InThisFormOrControl.Left + (InThisFormOrControl.Width - ThisControl.Width) \ 2
    End If

End Sub

'To center the form in the MDI parent (if exists) or Screen
'- strMode can specify "V" for vertical or "H" for horizontal (otherwise both)
Public Sub CenterTheForm(frm As Form, Optional ByVal strMode$ = "")
'modified 5/10/99 to accomodate multiple-monitors wide
'modified 11/7/00 to always center in MDI parent (if exists)
    Dim nLeft#, nTop#, nWidth#, nHeight#, i&
    Dim frmMDI As MDIForm
    Dim bMDIadjust As Boolean

    On Error Resume Next
    strMode = UCase(Left(strMode, 1))
    
    'Find MDI parent if exists
    If Not TypeOf frm Is MDIForm Then
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                Set frmMDI = Forms(i)
                With frmMDI
                    If IsMDIChild(frm) Then
                        nWidth = .ScaleWidth
                        nHeight = .ScaleHeight
                    Else
                        'for non-child forms, don't use MDI parent
                        'if it's not visible yet (e.g. splash screen),
                        'minimized, or off the screen (e.g. sometimes
                        'when the app is starting up)
                        If Not .Visible Or WindowStateX(frmMDI) = wsMinimized _
                                Or .Top + .Height <= 0 Then
                            Set frmMDI = Nothing '(don't use it)
                        Else
                            'must adjust for non-child forms using MDI
                            bMDIadjust = True
                            nWidth = .Width
                            nHeight = .Height
                        End If
                    End If
                End With
                Exit For
            End If
        Next
    End If
    'if no MDI parent, center in screen
    If frmMDI Is Nothing Then
        nWidth = Screen.Width
        nHeight = Screen.Height
    End If
    
    'Center horizontally
    If strMode = "V" Then
        nLeft = frm.Left 'leave where it is
    Else
        nLeft = (nWidth - frm.Width) \ 2
        If bMDIadjust Then
            nLeft = nLeft + frmMDI.Left
        End If
    End If
    
    'Center vertically
    If strMode = "H" Then
        nTop = frm.Top 'leave where it is
    Else
        nTop = (nHeight - frm.Height) \ 2
        If bMDIadjust Then
            nTop = nTop + frmMDI.Top
        End If
    End If
    
    'Perform move (more efficient to set both together
    'using .Move than to set .Left and .Top separately)
    If InStr(UCase(strMode), "X") > 0 Then
        MoveForm frm, nLeft, nTop
    Else
        frm.Move nLeft, nTop
    End If
    MoveFormOnScreen frm ', True
    
    Set frmMDI = Nothing
End Sub

' Sets drive and directory to new path.
' (returns false if not successful)
Public Function ChangePath(ByVal strNewPath$) As Boolean

    On Error Resume Next
    strNewPath = UCase(Trim(strNewPath))
    
    If SetCurrentDirectory(strNewPath) <> 0 Then
        ChangePath = True
    Else
        ChangePath = False
    End If

End Function


Public Sub KillFile(ByVal strFileMask As String, Optional ByVal bEvenIfReadOnlyOrHidden As Boolean = False)

    Dim nNumFiles&
    ' to be backwards-compatible with old routine, make sure to ignore
    ' paths (ending in a backslash), else will delete all files in folder!
    If Right(strFileMask, 1) <> "\" Then
        nNumFiles = DeleteFiles(strFileMask, bEvenIfReadOnlyOrHidden)
    End If

End Sub

Public Sub KillFolder(ByVal strFolder As String, Optional ByVal bDeleteFiles As Boolean = False)

    On Error Resume Next
    If DirExist(strFolder) Then
        If bDeleteFiles Then
            'KillFile AddSlash(strFolder) & "*.*", True
            DeleteFolder strFolder, True
        End If
        RmDir strFolder
    End If

End Sub

' To write a property to an INI file.
Public Function SetIniFileProperty(ByVal strPropName$, ByVal vPropValue As Variant, ByVal strSection$, ByVal strIniFile$) As Variant

    Dim rc&

    If Len(strIniFile) = 0 Then
        Beep
        InfBox "No .INI file specified ; t=30 ; i=E"
    Else
        If Len(Trim(strSection)) = 0 Then
            strSection = "General"
        End If
        ' PUT into INI file.
        rc = WritePrivateProfileString(strSection, strPropName, Str(vPropValue), strIniFile)
    End If

    SetIniFileProperty = vPropValue
End Function

' To read a property from an INI file.
Public Function GetIniFileProperty(ByVal strPropName$, ByVal vDefaultValue As Variant, ByVal strSection$, ByVal strIniFile$) As Variant

    Dim iStrLen&, strValue$, nMaxStringLen&
    
    If Len(strIniFile) = 0 Then
        Beep
        InfBox "No .INI file specified ; t=30 ; i=E"
    Else
        If Len(Trim(strSection)) = 0 Then
            strSection = "General"
        End If
        ' GET string from INI file (keep increasing buffer until big enough)
        nMaxStringLen = 250
        Do While nMaxStringLen < 10000000
            strValue = String(nMaxStringLen, 0)
            iStrLen = GetPrivateProfileString(strSection, strPropName, " ", strValue, nMaxStringLen, strIniFile)
            ' if buffer is too small, return value could be either nMaxStringLen-1 or nMaxStringLen-2
            If iStrLen < nMaxStringLen - 2 Then
                ' if not found, then just return default value
                If iStrLen > 0 Then
                    strValue = Trim(Left(strValue, iStrLen))
                    If Len(strValue) > 0 Then
                        If VarType(vDefaultValue) = vbString Then
                            vDefaultValue = strValue
                        Else
                            vDefaultValue = ValOfText(strValue, False)
                        End If
                    End If
                End If
                Exit Do
            End If
            nMaxStringLen = nMaxStringLen * 10
        Loop
    End If

    GetIniFileProperty = vDefaultValue
End Function

Public Function IsDigit(strChk$, Optional ByVal iPos& = 1) As Boolean

    Dim a%
    If iPos <= 0 Then iPos = 1
    If Len(strChk) >= iPos Then
        a = Asc(Mid(strChk, iPos, 1))
        If a >= 48 And a <= 57 Then IsDigit = True
    End If

End Function

' Returns true if an alpha character (A-Z or a-z)
' - if iPos > 0, then checks only the character at the specified position
' - if iPos = 0, then returns true if ANY alpha character in entire string
' - if iPos = -1 then return true if ALL alpha characters
Public Function IsAlpha(ByVal strText$, Optional ByVal iPos& = 0) As Boolean

    Dim a%
    If iPos = 0 Then
        ' return true if ANY alpha character in string
        For iPos = 1 To Len(strText)
            a = Asc(UCase(Mid(strText, iPos, 1)))
            If a >= 65 And a <= 90 Then
                IsAlpha = True
                Exit For
            End If
        Next
    ElseIf iPos > 0 And Len(strText) >= iPos Then
        ' check character at the specified position
        a = Asc(UCase(Mid(strText, iPos, 1)))
        If a >= 65 And a <= 90 Then
            IsAlpha = True
        End If
    ElseIf iPos = -1 Then
        ' return true if ALL alpha characters in string
        IsAlpha = True
        For iPos = 1 To Len(strText)
            a = Asc(UCase(Mid(strText, iPos, 1)))
            If (a < 65) Or (a > 90) Then
                IsAlpha = False
                Exit For
            End If
        Next
    End If

End Function

Public Function MakeDir(ByVal strNewDir$, Optional ByVal bBaseMustExist As Boolean = True) As Boolean

    Dim i&, bExist As Boolean

    strNewDir = Trim(strNewDir)
    If Right(strNewDir, 1) = "\" And Right(strNewDir, 2) <> ":\" Then
        strNewDir = Left(strNewDir, Len(strNewDir) - 1)
    End If

    ' Does path currently exist?
    On Error Resume Next
    bExist = DirExist(strNewDir)
    If bExist Then
        MakeDir = True
        Exit Function
    End If

    ' Try to make the directory.
    On Error GoTo MakeDirError
    MkDir (strNewDir)
    'If Len(Dir(NewDir, 16)) = 0 Then GoTo MakeDirError
    If DirExist(strNewDir) = 0 Then GoTo MakeDirError
    MakeDir = True
    Exit Function

MakeBase:
    ' Parse path and make directories.
    bBaseMustExist = True
    On Error Resume Next
    strNewDir = AddSlash(strNewDir)
    i = 1
    Do While True
        i = InStr(i + 1, strNewDir, "\")
        If i = 0 Then Exit Do
        If Mid(strNewDir, i - 1, 1) <> ":" Then
            MkDir Left(strNewDir, i - 1)
        End If
    Loop
    ' Check if exists now.
    On Error Resume Next
    bExist = DirExist(strNewDir)
    If bExist Then
        MakeDir = True
        Exit Function
    End If
    MakeDir = False
    Exit Function

MakeDirError:
    If Not bBaseMustExist Then Resume MakeBase
    MakeDir = False
    Exit Function

End Function

' for debugging
Public Sub DebugLog(strLogLine$, Optional ByVal bReinit As Boolean = False)

#If TRADENAV_EXE Then
    
    ' for TradeNav, use the cLogFile method:
    Static LogFile As cLogFile
    If (LogFile Is Nothing) Or bReinit Then
        Set LogFile = New cLogFile
        LogFile.OpenFile App.Path & "\Debug.Log", False
    End If
    LogFile.WriteText strLogLine

#Else

    Dim fh%
    Static strLogFile$, iLogging%
    
    If bReinit Then iLogging = 0
    If iLogging < 0 Then Exit Sub ' logging turned off
    If iLogging = 0 Then
        ' first time only
        strLogFile = AddSlash(App.Path) & "DEBUG.LOG"
        If Not FileExist(strLogFile) Then
            ' turn logging off if file not there to begin with
            iLogging = -1
            Exit Sub
        End If
        ' kill file to start fresh
        KillFile strLogFile
        iLogging = 1
    End If

    On Error Resume Next
    fh = FreeFile
    Open strLogFile For Append Shared As #fh
    If fh Then
        Print #fh, Format(Now, "hh:mm:ss") & " - " & strLogLine
        Close #fh
    End If
    
#End If

End Sub


Public Sub ListAdd(lst As Control, strItem$)
' adds item to list and moves highlight to bottom

    On Error Resume Next
    If Len(strItem) > 0 Then lst.AddItem strItem
    If lst.ListCount > 0 Then
        lst.ListIndex = lst.ListCount - 1
    End If
    lst.Refresh

End Sub

' Efficient method for finding a string in a list box
Public Function ListFind(lst As Control, ByVal strFind$) As Long

    Dim rc&

    strFind = strFind + Chr(0)
    rc = SendMessage(lst.hWnd, LB_FINDSTRINGEXACT, 0, ByVal strFind)

    ListFind = rc
End Function

Public Function ListFromFile(lst As Control, FileName$, _
        Optional append_to_list As Boolean = False) As Boolean

    Dim fh%, i&, Item$, rtrn As Boolean, is_list As Boolean

    On Error Resume Next
    rtrn = True
    If FileExist(FileName) Then
        If TypeOf lst Is ListBox Then
            is_list = True
        ElseIf TypeOf lst Is ctlUniComboImageXP Then
            is_list = True
        Else
            is_list = False
        End If
        fh = FreeFile
        Open FileName For Input As #fh
        If is_list Then
            ' list or combo box
            If append_to_list = False Then lst.Clear
            Do While Not EOF(fh)
                Line Input #fh, Item
                Item = Trim(Item)
                If Len(Item) > 0 Then lst.AddItem Item
            Loop
        Else
            ' text box
            If append_to_list = False Then lst = ""
            If FileLen(FileName) > 32000 Then
                i = 32000
            Else
                i = FileLen(FileName)
            End If
            lst = Input(i, fh)
        End If

        Close #fh
    Else
        rtrn = False
    End If

    ListFromFile = rtrn
End Function

Public Function ListItem(lst As Control) As String

    Dim strRtrn$
    strRtrn = ""
    On Error Resume Next
    If lst.ListCount > 0 And lst.ListIndex >= 0 Then
        strRtrn = lst.List(lst.ListIndex)
    End If
    ListItem = strRtrn

End Function

Public Sub ListRemove(lst As Control)

    On Error Resume Next
    If lst.ListCount > 0 And lst.ListIndex >= 0 Then
        lst.RemoveItem lst.ListIndex
        lst.Refresh
    End If

End Sub

Public Sub ListReplace(lst As Control, Item$, ByVal Num&)

    If Num >= lst.ListCount Then Num = lst.ListCount - 1
    If Num >= 0 Then lst.List(Num) = Item
    lst.Refresh

End Sub

Public Function ListToFile(lst As Control, FileName$, Optional ByVal item_start& = -1, Optional ByVal item_end& = -1) As Boolean

    Dim fh%, i&

    On Error GoTo ErrorListToFile
    fh = FreeFile
    Open FileName For Output As #fh

    If TypeOf lst Is ListBox Then
        ' list box
        If item_start < 0 Then item_start = 0
        If item_end < 0 Or item_end >= lst.ListCount Then item_end = lst.ListCount - 1

        For i = item_start To item_end
            Print #fh, lst.List(i)
        Next 'i
    Else
        ' text box
        Print #fh, lst
    End If

    Close #fh
    ListToFile = True
    Exit Function

ErrorListToFile:
    On Error Resume Next
    Close #fh
    ListToFile = False
End Function

' Returns value of formatted text (can choose to use regional settings
'   or not -- i.e. to recognize comma or period as decimal point), e.g.:
' "$10,000.00" returns 10000,  "($123.50)" returns -123.5
' "10%" returns 0.1,  "-1.2%" returns -0.012
' "4.35e3" returns 4350,  "4.35e-3" returns 0.00435
' "43.5K" returns 43500,  "43.5KB" returns 44544 (43.5 * 1024)
' "43.5M" returns 43500000,  "43.5MB" returns 45613056 (43.5 * 1024 * 1024)
' "43.5B" returns 43500000000,  "43.5GB" returns 46707769344
' "True" or "Yes" or "On" returns -1,  "False" or "No" or "Off" returns 0
Public Function ValOfText(ByVal strText As String, Optional ByVal bUseRegionalSettings As Boolean = True) As Double

    Dim dNum#, strTemp$, iPos%, iPower%, iFirstAlpha%
    Static strRegCheck$
    
    On Error Resume Next
    If bUseRegionalSettings Then
        ' always strip out $ regardless of regional settings
        strText = UCase(Trim(StripStr(strText, "$")))
        ' and strip out spaces and chr(160) if the group separator is a space (e.g. Czech
        ' settings seem to use the chr(160) as the group separator)
        If Len(strRegCheck) = 0 Then
            strRegCheck = Format(50000, "#,##0") ' (only need to get this once)
        End If
        If InStr(strRegCheck, Chr(160)) > 0 Or InStr(strRegCheck, " ") > 0 Then
            strText = StripStr(strText, Chr(160) & " ")
        End If
    Else
        ' if not using regional settings, strip out commas and $
        strText = UCase(Trim(StripStr(strText, "$,")))
    End If
    If Len(strText) > 0 Then
        ' check for "True", "Yes", and "On"
        If strText = "TRUE" Or strText = "T" Or strText = "YES" Or strText = "Y" Or strText = "ON" Then
            dNum = True
        Else
            ' find first alpha (A-Z), if exists, and truncate
            ' at first "bad" character (tab, new line, etc.)
            For iPos = 1 To Len(strText)
                strTemp = Mid(strText, iPos, 1)
                If strTemp < " " Or strTemp > "~" Then
                    strText = Trim(Left(strText, iPos - 1))
                    Exit For
                ElseIf iFirstAlpha = 0 Then
                    If (strTemp >= "A" And strTemp <= "Z") Or strTemp = "%" Then
                        iFirstAlpha = iPos
                    ' if not using regional settings, consider a space as end of number
                    ElseIf strTemp = " " And Not bUseRegionalSettings Then
                        iFirstAlpha = iPos + 1
                    End If
                End If
            Next
            If iFirstAlpha = 0 Then iFirstAlpha = Len(strText) + 1
                
            ' convert to a number
            dNum = 0
            If bUseRegionalSettings Then
                ' if using regional settings, use "CDbl" to convert as a number,
                ' but keep chopping rightmost character until doesn't error
                For iPos = iFirstAlpha - 1 To 1 Step -1
                    Err.Clear
                    dNum = CDbl(Trim(Left(strText, iPos)))
                    If Err.Number = 0 Then Exit For
                    dNum = 0
                Next
            Else
                ' if not using regional settings, use "Val" to convert as a number,
                ' (check for surrounding parentheses to indicate a negative number)
                If Left(strText, 1) = "(" And Right(strText, 1) = ")" Then
                    dNum = Val(Mid(strText, 2, iFirstAlpha - 2))
                    If dNum > 0 Then dNum = -dNum
                Else
                    dNum = Val(Left(strText, iFirstAlpha - 1))
                End If
            End If
                
            If dNum <> 0 And iFirstAlpha <= Len(strText) Then
                ' now adjust based on alpha portion
                Select Case Mid(strText, iFirstAlpha, 1)
                Case "%" '(percent)
                    dNum = dNum / 100#
                Case "E" '(scientific notation)
                    iPower = 0 '(in case of error on next line)
                    iPower = Val(Mid(strText, iFirstAlpha + 1))
                    If iPower <> 0 Then
                        dNum = dNum * (10# ^ iPower)
                    End If
                Case "K" '(kilo, kb's)
                    If Mid(strText, iFirstAlpha + 1, 1) = "B" Then
                        dNum = dNum * 1024#
                    Else
                        dNum = dNum * 1000#
                    End If
                Case "M" '(millions, megs)
                    If Mid(strText, iFirstAlpha + 1, 1) = "B" Then
                        dNum = dNum * 1024# * 1024#
                    Else
                        dNum = dNum * 1000000#
                    End If
                Case "B" '(billions)
                    dNum = dNum * 1000000000#
                Case "G" '(gigs)
                    If Mid(strText, iFirstAlpha + 1, 1) = "B" Then
                        dNum = dNum * 1024# * 1024# * 1024#
                    End If
                End Select
            End If
        End If
    End If

    ValOfText = dNum
End Function

' Gets text from a window using API call (Note: using "SendMessage" will
' lock up if encounters a non-responding top-level window, whereas
' "GetWindowText" will work for top-level windows without locking)
Public Function vbGetWindowText(ByVal hWnd As Long, Optional ByVal nMaxLen& = 512, _
        Optional ByVal bUseSendMessage As Boolean = False) As String

    Dim strText$, rc&
    
    If hWnd = 0 Then hWnd = GetActiveWindow()
    strText = Space(nMaxLen)
    If bUseSendMessage Then
        rc = SendMessage(hWnd, &HD, Len(strText), ByVal strText)
    Else
        rc = GetWindowText(hWnd, strText, Len(strText))
    End If
    If rc <= 0 Then
        vbGetWindowText = ""
    Else
        vbGetWindowText = Left(strText, rc)
    End If
    
End Function

#If 0 Then
' returns date from "long date" (e.g. 19991203)
Public Function JulFromLong(ByVal lDate As Long) As Long
' used to use C Dll function (much more efficient),
' but have to use this for now in 32-bit env.

    Dim strDate$, iLen%
    strDate = Trim(Str(lDate))
    iLen = Len(strDate)
    'If DateEuropean() Then
    '    ndate = Mid(sdate, slen - 1, 2) + "/" + Mid(sdate, slen - 3, 2) + "/" + Left(sdate, slen - 4)
    'Else
        strDate = Mid(strDate, iLen - 3, 2) + "/" + Mid(strDate, iLen - 1, 2) + "/" + Left(strDate, iLen - 4)
    'End If
    ' convert string to julian date
    If IsDate(strDate) Then
        JulFromLong = CVDate(strDate)
    Else
        JulFromLong = 0
    End If

End Function
#End If

#If 0 Then
' used to call DLL function in 16-bit env.
Public Function DirExist(strPath$) As Boolean
    
    Dim bExist%
    On Error Resume Next
    bExist = True
    If Len(Dir(strPath, 16)) = 0 Then bExist = False
    
    DirExist = bExist
End Function
#End If

' To store a function pointer in a variable or member of a structure
' (required for some DLL calls), we must use this wrapper function
' (since the "AddressOf" operator is only valid as a function argument).
' e.g.      Dim mt As MyType
'           mt.MyPtr = FunctionPtrToLong(AddressOf MyCallBackFunction)
Public Function FunctionPtrToLong(ByVal FunctionPtr As Long) As Long
    FunctionPtrToLong = FunctionPtr
End Function

' to add century to long date (YYMMDD), if not already
Public Function AddCentury(ByVal lDate As Long, _
    Optional ByVal crossover_year As Long = 15) As Long
    
    If lDate <= 0 Then
        lDate = 0 ' invalid
    ' check "crossover" for 6-digit dates
    ElseIf lDate < crossover_year * 10000 Then
        lDate = lDate + 20000000
    ' else check up to year 299, to accomodate
    ' 2 or 3-digit years (like CSI/MS7 Y2K dates)
    ElseIf lDate <= 2991231 Then
        lDate = lDate + 19000000
    End If
    
    AddCentury = lDate

End Function

'Zip actions: 'C'=Create, 'A'=Append, 'D'=Delete, 'U'=Unzip,
'        'N'=Number, 'S'=Size(uncompressed), 'F'=Find matching files
'(TLB 4/17/2008: now returns a Double in case 'S'ize is > signed long)
Public Function ZipExecute(ByVal strAction$, _
        ByVal strZipFile$, _
        ByVal strUnzippedPath$, _
        Optional ByVal strFileMask$ = "", _
        Optional ByVal bRecursive As Boolean = False, _
        Optional ByVal bOverwriteNewer As Boolean = False, _
        Optional ByVal vFirstDate As Variant = 0, _
        Optional ByVal vLastDate As Variant = 0, _
        Optional ByVal hFilesWin As Long = 0, _
        Optional ByVal hProgressWin As Long = 0, _
        Optional strErrMsg$) As Double
    
    Dim zargs As zip_args, dVal#, nSaveMouse%
    Dim vDate As Date

    nSaveMouse = Screen.MousePointer
    Screen.MousePointer = 11

    If strAction = "A" Then
        If Not FileExist(strZipFile) Then
            strAction = "C"
        End If
    End If

    zargs.Action = UCase(Left(Trim(strAction), 1))
    zargs.zip_file = Trim(strZipFile) & Chr(0)
    zargs.file_masks = Trim(strFileMask) & Chr(0)
    zargs.unzip_path = AddSlash(Trim(strUnzippedPath)) & Chr(0)
    zargs.recursive = bRecursive
    zargs.strip_paths = Not bRecursive
    zargs.overwrite_newer = bOverwriteNewer
    zargs.file_window = hFilesWin
    zargs.progress_window = hProgressWin
    
    vDate = DateOf(vFirstDate)
    If Year(vDate) <= 1900 Then
        zargs.start_date = 0
    Else
        zargs.start_date = JulToLong(CLng(vDate), True)
    End If
    
    vDate = DateOf(vLastDate)
    If Year(vDate) <= 1900 Then
        zargs.end_date = 0
    Else
        zargs.end_date = JulToLong(CLng(vDate), True)
    End If
    
    dVal = GenZip(zargs)
    
    ' fix wrap-around when size is being returned
    If zargs.Action = "S" And dVal < 0 Then
        dVal = dVal + 2# ^ 32
    End If
    
    If Not IsMissing(strErrMsg) Then
        strErrMsg = zargs.err_msg
        FixNullTermStr strErrMsg
    End If

    Screen.MousePointer = nSaveMouse
    ZipExecute = dVal
End Function

' scans subdirectories from "root" for matching files
' (outputs size,date,time,name of files to "asc_file")
Public Function ScanForFiles(ByVal strAscFile$, _
        ByVal strRoot$, _
        Optional ByVal strFileMask$ = "", _
        Optional ByVal bRecursive As Boolean = False, _
        Optional ByVal vFirstDate As Variant = 0, _
        Optional ByVal vLastDate As Variant = 0) As Long
    
    ScanForFiles = ZipExecute("F", strAscFile, strRoot, _
        strFileMask, bRecursive, False, vFirstDate, vLastDate)
End Function

' Rounds twips to nearest pixel
Public Function RoundTwips(ByVal nTwips As Long) As Long
    Dim pix#
    pix = Screen.TwipsPerPixelX
    RoundTwips = Int(nTwips / pix + 0.5) * pix
End Function


Public Sub ClientToScreenTwips(ctlFrom As Control, lpPoint As POINTAPI, _
        Optional nPixelSubtract As Long = 2)
    
    Dim rc&
    ' first convert to pixels
    lpPoint.X = RoundTwips(lpPoint.X) \ Screen.TwipsPerPixelX
    lpPoint.Y = RoundTwips(lpPoint.Y) \ Screen.TwipsPerPixelY
    ' convert to screen coordinates
    rc = ClientToScreen(ctlFrom.hWnd, lpPoint)
    ' subtract 2 pixels (don't know why!)
    lpPoint.X = lpPoint.X - nPixelSubtract
    lpPoint.Y = lpPoint.Y - nPixelSubtract
    ' convert back to twips
    lpPoint.X = lpPoint.X * Screen.TwipsPerPixelX
    lpPoint.Y = lpPoint.Y * Screen.TwipsPerPixelY

End Sub

' Loads string array with all matching files in path.
Public Function GetMatchingFiles(aFiles() As String, ByVal strPath$, _
        Optional ByVal strFileMask$ = "*.*") As Long

    Dim nCount&, strFile$
    On Error Resume Next
    strFile = Dir(AddSlash(strPath) & strFileMask)
    Do While Len(strFile) > 0
        If nCount Mod 100 = 0 Then
            ReDim Preserve aFiles(nCount + 101) As String
        End If
        nCount = nCount + 1
        aFiles(nCount) = strFile
        strFile = Dir
    Loop
    ReDim Preserve aFiles(nCount) As String

    GetMatchingFiles = nCount
End Function

' Allow user to select color from common dialog.
' (returns -1 if user canceled selection)
Public Function CommonDialogColor(ctlCommonDialog As Control) As Long
    
    On Error GoTo ErrNotCommonDlg
    ' Note: "ctlCommonDialog" passed in must be a Common Dialog control
    ' (couldn't use "as CommonDialog" since a project which
    ' never calls this might not even have one of these controls)
    
    With ctlCommonDialog
        ' Set Cancel to True.
        .CancelError = True
        ' Set the Flags property.
        .Flags = 1 'cdlCCRGBInit
    End With
    ' Display the Color dialog box.
    On Error GoTo CancelClicked
    ctlCommonDialog.ShowColor
    ' Set color to the selected color.
    CommonDialogColor = ctlCommonDialog.Color
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function

ErrNotCommonDlg:
    ' (see note above)
    CommonDialogColor = -1
    Err.Raise vbObjectError + 999, "CommonDialogColor", "PROGRAM ERROR: CommonDialogColor must be passed a Common Dialog control!"
    Exit Function

CancelClicked:
    ' User pressed Cancel button.
    CommonDialogColor = -1
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function
End Function

' Allow user to select a file using the common dialog.
' - returns blank if user canceled selection
' - default flag (-1) is to hide read-only files only when saving
' - specify delimiter to put dialog into multi-select mode ( delimiter will be used in return value )
Public Function CommonDialogFile(ctlCommonDialog As Control, ByVal bSaveFile As Boolean, _
        Optional ByVal strFilter$ = "", _
        Optional ByVal strInitialFileOrFolder$ = "", _
        Optional ByVal strTitle$ = "", _
        Optional ByVal lFlags As Long = -1, _
        Optional ByVal strDelimiter As String = "") As String
    
    On Error GoTo ErrNotCommonDlg
    ' Note: "ctlCommonDialog" passed in must be a Common Dialog control
    ' (couldn't use "as CommonDialog" since a project which
    ' never calls this might not even have one of these controls)

    Dim strReturn As String             ' Return value for the function
    Dim strFixed As String              ' Fixed version of the return value
    Dim strPath As String               ' Path for the selected files
    Dim strFile As String               ' File selected
    Dim strToAdd As String              ' Path and file to add to the return value
       
    ' default flag
    If lFlags = -1 Then
        If bSaveFile Then
            lFlags = 4 'cdlOFNHideReadOnly
        Else
            lFlags = 0
        End If
    End If
    
    With ctlCommonDialog
        ' Set Cancel to True.
        .CancelError = True
        ' Set the properties
        .Flags = lFlags Or 524288 'cdlOFNExplorer
        If Len(strDelimiter) > 0 Then
            .Flags = .Flags Or 512 'cdlOFNAllowMultiselect
        End If
        
        If Len(strTitle) > 0 Then
            .DialogTitle = strTitle
        ElseIf bSaveFile Then
            .DialogTitle = "Save as ..."
        Else
            .DialogTitle = "Select file ..."
        End If
        If Len(strFilter) > 0 Then
            .Filter = strFilter '"Wave Files (*.wav)|*.wav"
        Else
            .Filter = "All files|*.*"
        End If
        .FilterIndex = 1
        If DirExist(strInitialFileOrFolder) Then
            .FileName = ""
            .InitDir = strInitialFileOrFolder
        Else
            .FileName = strInitialFileOrFolder
        End If
    End With
    ' Display the File dialog box.
    On Error GoTo CancelClicked
    If bSaveFile Then
        ctlCommonDialog.ShowSave
    Else
        ctlCommonDialog.ShowOpen
    End If
    
    strReturn = ctlCommonDialog.FileName
    If (InStr(strReturn, Chr(0)) > 0) And (Len(strDelimiter) > 0) Then
        strPath = AddSlash(Parse(strReturn, Chr(0), 1))
        strFile = Parse(strReturn, Chr(0), 0)
        Do While Len(strFile) > 0
            If (strDelimiter = ",") Or (strDelimiter = " ") Then
                strToAdd = Chr(34) & strPath & strFile & Chr(34)
            Else
                strToAdd = strPath & strFile
            End If
            
            If Len(strFixed) = 0 Then
                strFixed = strToAdd
            Else
                strFixed = strFixed & strDelimiter & strToAdd
            End If
            
            strFile = Parse(strReturn, Chr(0), 0)
        Loop
        
        strReturn = strFixed
    End If
    
    ' return filename
    CommonDialogFile = strReturn
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function

ErrNotCommonDlg:
    ' (see note above)
    CommonDialogFile = ""
    Err.Raise vbObjectError + 999, "CommonDialogColor", "PROGRAM ERROR: CommonDialogColor must be passed a Common Dialog control!"
    Exit Function

CancelClicked:
    ' User pressed Cancel button.
    CommonDialogFile = ""
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function
End Function

' Allow user to select font from common dialog.
' (returns -1 if user canceled selection)
Public Function CommonDialogFont(ctlCommonDialog As Control, Font As StdFont) As Boolean

    On Error GoTo ErrNotCommonDlg
    ' Note: "ctlCommonDialog" passed in must be a Common Dialog control
    ' (couldn't use "as CommonDialog" since a project which
    ' never calls this might not even have one of these controls)
    
    With ctlCommonDialog
        ' Set Cancel to True.
        .CancelError = True
        ' Set the Flags property.
        .Flags = &H1 Or &H400 'cdlCFScreenFonts Or cdlCFANSIOnly
        ''ctlCommonDialog.Flags = ctlCommonDialog.Flags Or &H100 'cdlCFEffects
        ' Set the current font settings
        .Font.Name = Font.Name
        .Font.Size = Font.Size
        .Font.Bold = Font.Bold
        .FontItalic = Font.Italic
        .FontUnderline = Font.Underline
        .FontStrikethru = Font.Strikethrough
    End With
    ' Display the Font dialog box.
    On Error GoTo CancelClicked
    ctlCommonDialog.ShowFont
    With ctlCommonDialog
        ' Set to new font settings
        Font.Name = CheckSSFont(.Font.Name)
        Font.Size = .Font.Size
        Font.Bold = .Font.Bold
        Font.Italic = .FontItalic
        Font.Underline = .FontUnderline
        Font.Strikethrough = .FontStrikethru
        CommonDialogFont = True
    End With
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function

ErrNotCommonDlg:
    ' (see note above)
    Err.Raise vbObjectError + 999, "CommonDialogFont", "PROGRAM ERROR: CommonDialogFont must be passed a Common Dialog control!"
    Exit Function

CancelClicked:
    ' User pressed Cancel button.
    ChangePath App.Path ' TLB 10/26/2011: should set current path back, since the common dialog can change it
    Exit Function
End Function

' To restore and activate the previous instance
' of this program instead of starting a new one,
' put the following as the first line of code ...
'   If ActivatePrevInstance then End
Public Function ActivatePrevInstance() As Boolean
    
    Dim hWnd&, strTitle$, bActivatePrev As Boolean

    If App.PrevInstance Then
        'check for flag file to allow 2nd instance
        If Not FileExist(App.Path & "\Allow.2nd") Then
            bActivatePrev = True
        End If
    End If
        
    If bActivatePrev Then
        'save the title of the application
        strTitle = App.Title
        'rename so FindWindow will not find this instance
        App.Title = "_unwanted instance_"
        'attempt to get window handle using VB6 class name
        hWnd = FindWindow("ThunderRT6Main", strTitle)
        'get handle to previous window
        If IsWindow(hWnd) <> 0 Then
            hWnd = GetWindow(hWnd, GW_HWNDPREV)
            If IsWindow(hWnd) <> 0 Then
                'if minimized, then restore to prior state
                If IsIconic(hWnd) Then ShowWindow hWnd, SW_OTHERUNZOOM
                'and activate it
                SetForegroundWindow hWnd
            End If
        End If
    End If
    
    ActivatePrevInstance = bActivatePrev
End Function

' Can use this to initially size an MDI child form
' (sizes to bottom-right corner of a hidden control on form)
Public Sub SizeFormToControl(frm As Form, ctl As Control)
    frm.Width = ctl.Left + ctl.Width + frm.Width - frm.ScaleWidth
    frm.Height = ctl.Top + ctl.Height + frm.Height - frm.ScaleHeight
End Sub

' returns file extension
Public Function FileExt(ByVal strFileName$) As String
    Dim i&, strExt$
    ' find last dot
    i = At(strFileName, ".", -1)
    If i > 0 Then
        ' if a backslash after this dot, then it's not the extension
        If InStr(i, strFileName, "\") = 0 Then
            strExt = Mid(strFileName, i + 1)
        End If
    End If
    FileExt = strExt
End Function

' returns file path (with a terminating slash)
Public Function FilePath(ByVal strFileName$) As String
    Dim i&
    ' find last slash
    i = At(strFileName, "\", -1)
    ' or check for drive (in case passed as "C:filename")
    If i = 0 Then i = InStr(strFileName, ":")
    If i > 0 Then
        FilePath = AddSlash(Left(strFileName, i))
    Else
        FilePath = ""
    End If
End Function

' returns filename with path and extension stripped off
Public Function FileBase(ByVal strFileName$) As String
    Dim i&
    ' find last slash
    i = At(strFileName, "\", -1)
    ' or check for drive (in case passed as "C:filename")
    If i = 0 Then i = InStr(strFileName, ":")
    If i > 0 Then strFileName = Mid(strFileName, i + 1)
    ' find last dot
    i = At(strFileName, ".", -1)
    If i > 0 Then strFileName = Left(strFileName, i - 1)
    FileBase = strFileName
End Function

' Replace file extension of a filename (passed file may or may not have an extension)
Public Function ReplaceFileExt(ByVal strFileName$, ByVal strExt$) As String
    Dim i&
    i = Len(FileExt(strFileName))
    If i > 0 Then strFileName = Left(strFileName, Len(strFileName) - i - 1)
    If Left(strExt, 1) <> "." Then strExt = "." & strExt
    ReplaceFileExt = strFileName & strExt
End Function

' Returns elapsed # seconds between two TickCounts -- using GetTickCount()
' (handles 49-day tick count wrap-around problem)
' - leave off second arg if want elapsed time to Now
Public Function ElapsedSeconds(ByVal FromTickCount&, Optional ByVal ToTickCount) As Double
    Dim dTo#
    If IsMissing(ToTickCount) Then
        dTo = GetTickCount() ' to Now
    Else
        dTo = ToTickCount
    End If
    If FromTickCount > dTo + 2# ^ 16 Then
        dTo = dTo + 2# ^ 32 ' fixes wrap-around problem
    End If
    ElapsedSeconds = (dTo - FromTickCount) / 1000# ' as # of seconds
End Function

'To see if a specific bit (1-32) of a value is on or off.
Public Function GetBit(ByVal OfValue As Variant, ByVal BitNum As Long) As Boolean
    Dim FilterMask As Long
    If BitNum < 0 Then '(if negative, the filter mask was passed in directly)
        FilterMask = Abs(BitNum)
    ElseIf BitNum = 32 Then '(must handle this way to avoid overflow)
        FilterMask = -2 ^ (BitNum - 1)
    Else
        FilterMask = 2 ^ (BitNum - 1)
    End If
    GetBit = CLng(OfValue) And FilterMask
End Function

'To set a specific bit (1-32) of a value either on or off.
Public Sub SetBit(OfValue As Variant, ByVal BitNum As Long, ByVal SetToOn As Boolean)
    Dim FilterMask As Long
    If BitNum < 0 Then '(if negative, the filter mask was passed in directly)
        FilterMask = Abs(BitNum)
    ElseIf BitNum = 32 Then '(must handle this way to avoid overflow)
        FilterMask = -2 ^ (BitNum - 1)
    Else
        FilterMask = 2 ^ (BitNum - 1)
    End If
    If SetToOn Then
        OfValue = CLng(OfValue) Or FilterMask
    Else
        OfValue = CLng(OfValue) And Not FilterMask
    End If
End Sub

'To get the value of just specific bits (1-32)
Public Function GetValOfBits(ByVal dwFlags As Long, ByVal AtBitNum As Integer, ByVal NumBits As Integer) As Long
    Dim iBit&, iResult&, iPow2&
    ' if want all 32 bits, just return it
    If AtBitNum = 1 And NumBits = 32 Then
        iResult = dwFlags
    ' else make sure requested bits are in range (i.e. 1-32)
    ElseIf AtBitNum >= 1 And AtBitNum <= 32 And NumBits > 0 And AtBitNum + NumBits - 1 <= 32 Then
        iPow2 = 1
        For iBit = AtBitNum To 32
            If GetBit(dwFlags, iBit) Then
                iResult = iResult + iPow2
            End If
            If iBit = AtBitNum + NumBits - 1 Then Exit For ' need to check now so next line won't overflow
            iPow2 = iPow2 * 2
        Next
    End If
    GetValOfBits = iResult
End Function

'To set the value of just specific bits (1-32) within the Flags variable
Public Sub SetValOfBits(dwFlags As Long, ByVal AtBitNum As Integer, ByVal NumBits As Integer, ByVal ValOfBits As Long)
    Dim iBit&, iResult&, iPow2&
    ' make sure requested bits are in range (i.e. 1-32)
    If AtBitNum >= 1 And AtBitNum <= 32 And NumBits > 0 And AtBitNum + NumBits - 1 <= 32 Then
        For iBit = 1 To NumBits
            SetBit dwFlags, AtBitNum + iBit - 1, GetBit(ValOfBits, iBit)
        Next
    End If
End Sub

'Returns vValue if it's not Null, else returns vIfNull
'(the vIfNull is optional ONLY for strings -- it should
' NOT be considered optional for numbers and dates).
Public Function NullChk(ByVal vValue As Variant, _
        Optional ByVal vIfNull As Variant = "") As Variant
    If IsNull(vValue) Then
        NullChk = vIfNull
    Else
        NullChk = vValue
    End If
End Function

'Returns True if form is a StingRay FormX MDIchild
Public Function IsFormX(frm As Form) As Boolean
    On Error GoTo FuncExit
    IsFormX = False
    If frm.FormX1.Edge = -99999 Then DoEvents
    IsFormX = True
FuncExit:
    Exit Function
End Function

' Returns MDI's active form (works if using StingRay or not)
Public Function MDIActiveForm() As Form

    Dim i&
    Static frmMDI As Object 'could be form or control
    
    On Error Resume Next
    If frmMDI Is Nothing Then
        ' if first time, find the MDI parent
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                'see if form has a StingRay MDIFormX control
                Set frmMDI = Forms(i).MDIFormX1
                If frmMDI Is Nothing Then
                    'if not, then set to form itself
                    Set frmMDI = Forms(i)
                End If
                Exit For
            End If
        Next
        If frmMDI Is Nothing Then
            Set MDIActiveForm = Nothing
            Exit Function
        End If
    End If
    
    Set MDIActiveForm = frmMDI.ActiveForm
End Function

' Returns the MDI form
Public Function MDIForm() As MDIForm

    Dim i&
    Static frmMDI As MDIForm
    
    On Error Resume Next
    If frmMDI Is Nothing Then
        ' if first time, find the MDI parent
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                Set frmMDI = Forms(i)
                Exit For
            End If
        Next
    End If
    
    Set MDIForm = frmMDI
End Function

'Returns True if form is an MDIChild (works for either a
' a normal VB child, or a StingRay FormX MDIChild)
Public Function IsMDIChild(frm As Form) As Boolean

    Dim i&, bChild As Boolean
    On Error Resume Next
    If Not frm Is Nothing Then
        With frm
            'see if form is a normal VB MDIchild
            i = .MDIChild
            If i Then
                bChild = True
            Else 'see if form is a StingRay FormX MDIchild
                bChild = IsFormX(frm)
            End If
        End With
    End If
    IsMDIChild = bChild
End Function

Public Property Get WindowStateX(frm As Form) As vbWindowState
    ' see if StingRay FormX control exists
    On Error Resume Next
    WindowStateX = -1
    WindowStateX = frm.FormX1.WindowState
    If WindowStateX < 0 Then WindowStateX = frm.WindowState
End Property

Public Property Let WindowStateX(frm As Form, ByVal ws As vbWindowState)
    Dim iChk%
    ' see if StingRay FormX control exists
    On Error Resume Next
    iChk = -1
    iChk = frm.FormX1.Edge
    If iChk >= 0 Then
        ' if so, must set its WindowState property
        If frm.FormX1.WindowState <> ws Then
            frm.FormX1.WindowState = ws
        End If
    ElseIf frm.WindowState <> ws Then
        ' otherwise, set property of form itself
        frm.WindowState = ws
    End If
End Property

'Moves a form whether a StingRay or not
Public Sub MoveForm(frm As Form, Optional ByVal nLeft& = NULL_DATA, Optional ByVal nTop& = NULL_DATA, Optional ByVal nWidth& = NULL_DATA, Optional ByVal nHeight& = NULL_DATA)
    'more efficient to use "Move" to set all at once
    '(first see if a StingRay form)
    If nLeft <= NULL_DATA Then nLeft = frm.Left
    If nTop <= NULL_DATA Then nTop = frm.Top
    If nWidth < 0 Then nWidth = frm.Width
    If nHeight < 0 Then nHeight = frm.Height
    If IsFormX(frm) Then
        frm.FormX1.DoMDI nLeft, nTop, nWidth, nHeight
    Else
        frm.Move nLeft, nTop, nWidth, nHeight
    End If
End Sub

Public Function IntAsUnsigned(ByVal i As Integer) As Long
    Dim u As Long
    u = i
    If u < 0 Then u = u + 65536
    IntAsUnsigned = u
End Function

Public Function LongAsUnsigned(ByVal i As Long) As Double
    Dim u As Double
    u = i
    If u < 0 Then u = u + 4294967296#
    LongAsUnsigned = u
End Function

'Returns True only if the specified item exists in the collection.
Public Function ItemExists(ByVal objCollection As Object, ByVal vItem As Variant) As Boolean
    On Error GoTo FuncExit
    ItemExists = False
    If objCollection Is Nothing Then GoTo FuncExit
    If objCollection.Item(vItem) Is Nothing Then GoTo FuncExit
    'will only get here if the item exists in the collection
    ItemExists = True
FuncExit:
    Exit Function
End Function

Public Sub ShowFormLog(frm As Form, ByVal bShowingNow As Boolean)

    On Error Resume Next
    Dim strMessage$
    If bShowingNow Then
        strMessage = "0" & vbTab & ShowFormLogMsg(frm)
    Else
        strMessage = "-1" & vbTab & ShowFormLogMsg(frm)
    End If
    ShowFormLogToFile strMessage

End Sub

Private Sub ShowFormLogToFile(ByVal strMessage$)
On Error Resume Next

    Dim fh%
    Dim strPath$
    Static bAlreadyDone As Boolean
    
    If Len(strMessage) = 0 Then Exit Sub

    strPath = AddSlash(App.Path) & "ShowFormLogs\"
    If Not bAlreadyDone Then
        bAlreadyDone = True
        MakeDir strPath
        ' clean out old files
        'KillFile strPath & "*.LOG /o=-90"
    End If

    fh = FreeFile
    Open strPath & Format(Date, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ")" & vbTab & strMessage
        Close #fh
    End If

End Sub

Private Function ShowFormLogMsg(frm As Form) As String

    On Error Resume Next
    Dim strLogMsg$
    
    If Not frm Is Nothing Then
        strLogMsg = frm.Name & vbTab & frm.Caption
        If TypeOf frm Is frmAsk Then
            strLogMsg = strLogMsg & vbTab & Replace(frm.lblMessage, Chr(13), " | ")
        End If
    End If
    ShowFormLogMsg = strLogMsg

End Function

' Use this to show forms FAST (esp. when non-modal)
' (FYI: ActModal is needed for things like a form with the Tradesense editor control
'  which has a non-modal popup window being displayed overtop the control)
Public Sub ShowForm(frm As Form, Optional ByVal eModal As eShowFormMode = eForm_Nonmodal, _
        Optional ByVal frmOwner As Form = Nothing, Optional ByVal bAllowOffScreen As Boolean = False, _
        Optional ByVal gridAltRowColor As Long = 0)
    
    Dim i&, hMenu&, nMinShowing&, hWnd&, strLogMsg$
    Dim frmPrev As Form
    Dim aMinimized() As Boolean
    Static FormsActingModal As New Collection   '(to keep a "stack" of modal-acting forms)
    
    eModal = Abs(eModal) '(to be backwards-compatible)
    If Not frm Is Nothing Then
        ' TLB 4/5/2005: replace default font with TrueType font for all controls on form
        FixFormControls frm, gridAltRowColor
    
        'getting WindowState now will also make sure
        'the form gets loaded before showing
        If WindowStateX(frm) = wsMinimized Then WindowStateX(frm) = wsNormal
        ' if current form is already acting modal, then this form needs to be acting modal as well
        If FormsActingModal.Count > 0 Then
            eModal = eForm_ActModal
        ElseIf eModal = eForm_ActModal Then
            ' TLB 3/16/2009: otherwise, if MDI is already modal when trying to
            ' do the first "act modal", then just treat this as a regular modal
            If Not MDIForm Is Nothing Then
                If MDIForm.Enabled = False Then
                    eModal = eForm_Modal
                End If
            End If
        End If
        If eModal = eForm_Modal Then
            'can't do child forms modally
            If IsMDIChild(frm) Then eModal = eForm_Nonmodal
        End If
        
        ' make sure form will show on-screen
        If eModal <> eForm_Nonmodal Then
            'don't want to allow a modal form to be minimized
            If frm.MinButton Then
                hMenu = GetSystemMenu(frm.hWnd, 0)
                If hMenu <> 0 Then RemoveMenu hMenu, SC_MINIMIZE, 0
            End If
            'make sure entire modal form will show on-screen
            If WindowStateX(frm) = wsNormal Then
                MoveFormOnScreen frm
            End If
            ' TLB 10/7/2011: needed this when showing any modal form in case desktop painting had been locked
            LockWindowUpdate 0
        ElseIf Not bAllowOffScreen And Not IsMDIChild(frm) And WindowStateX(frm) = wsNormal Then
            'make sure form will show at least partially on-screen
            MoveFormOnScreen frm
        End If
        
        ' TLB 9/7/2012: to log each time a form is shown
        strLogMsg = ShowFormLogMsg(frm)
            
        ' show the form
        If eModal = eForm_Modal Then
            ' TLB - needed to provide a workaround solution for the following problem:
            ' if minimize an editor to the taskbar, then bring up toolbox (modal), then click on the
            ' minimized editor, the focus leaves the toolbox (making app appear locked) -- you have to
            ' click on TradeNav in the taskbar from 1-3 times to get the focus back on the toolbox.
            ' (if the form on the taskbar is not minimized then it's not a problem)
            If Not MDIForm Is Nothing Then
                If MDIForm.Enabled Then
                    ' if MDI is enabled (meaning no modal form up yet), then find all the
                    ' minimized non-child forms -- make them not visible and restore later
                    ReDim aMinimized(Forms.Count) As Boolean
                    For i = 1 To Forms.Count
                        Set frmPrev = Forms(i - 1)
                        If Not frmPrev Is Nothing And Not frmPrev Is MDIForm Then
                            If frmPrev.WindowState = 1 And frmPrev.Visible And Not frmPrev.MDIChild Then
                                aMinimized(i) = True
                                frmPrev.Visible = False
                            End If
                        End If
                    Next
                    Set frmPrev = Nothing
                Else
                    ReDim aMinimized(0) As Boolean
                End If
            End If
        
            ' show modally
            ShowFormLogToFile Str(eModal) & vbTab & strLogMsg
            If frmOwner Is Nothing Then
                frm.Show 1, MDIForm  '(pass MDI Parent, if exists, so App shows in Alt-Tab list)
            Else
                frm.Show 1, frmOwner '(use specified owner)
            End If
            ShowFormLogToFile "-1" & vbTab & strLogMsg
            DoEvents '(to help hide it faster)
            
            If Not MDIForm Is Nothing Then
                ' restore any minimized forms that had been temporarily made not visible
                For i = 1 To UBound(aMinimized)
                    If aMinimized(i) And i <= Forms.Count Then
                        Set frmPrev = Forms(i - 1)
                        If frmPrev.WindowState = 1 And Not frmPrev.Visible And Not frmPrev.MDIChild Then
                            frmPrev.Visible = True
                        End If
                    End If
                Next
                Set frmPrev = Nothing
            End If
        Else
            'if no owner specified, then default to the MDI parent
            'as owner (if it exists), unless set to "ShowInTaskbar"
            frm.Enabled = True
            ShowFormLogToFile Str(eModal) & vbTab & strLogMsg
            If frm.MDIChild Then
                frm.Show 0 '(cannot specify owner for MDI child)
            ElseIf Not frmOwner Is Nothing Then
                frm.Show 0, frmOwner
            ElseIf frm.ShowInTaskbar Then
                frm.Show 0
            Else
                frm.Show 0, MDIForm
            End If
            'immediately do a Refresh to get painted now
            frm.Refresh
            On Error Resume Next
            ' TLB 9/13/2006: Comment out the .SetFocus since we're not doing StingRay anymore
            ''frm.SetFocus '(StingRay forms don't always do this)
            
            If eModal = eForm_ActModal Then
                ' add to "acting modal" stack (do this after showing the form in case of an error)
                FormsActingModal.Add frm
                ' "act modal" (disable all other forms and loop until not visible)
                EnableAllForms False, frm
                hWnd = frm.hWnd '(use variable so won't reload form in loop if already unloaded)
                i = 0
                Do  ' TLB 5/28/2009: use "Sleep" (full idle) so won't peg CPU,
                    ' but only once every 100 times otherwise the keystrokes
                    ' don't always work right on the form (e.g. the Tab key)
                    i = i + 1
                    If i < 100 Then
                        DoEvents '(cannot use "Sleep" here)
                    Else
                        Sleep 0
                        i = 0
                    End If
                    If IsWindow(hWnd) = 0 Then Exit Do
                Loop While IsWindowVisible(hWnd) <> 0
                ' remove form from "acting modal" stack
                FormsActingModal.Remove FormsActingModal.Count
                If FormsActingModal.Count = 0 Then
                    ' if no more modal-acting forms on the stack, then enable all other forms
                    EnableAllForms True
                    
                    ' In this case, we end up not having any form in the application with the focus,
                    ' so set the focus to the owner if given (DAJ: 10/24/2007)...
                    If Not frmOwner Is Nothing Then
                        MoveFocus frmOwner
                    End If
                Else
                    ' else re-enable and activate the modal-acting form on top of the stack
                    Set frmPrev = FormsActingModal(FormsActingModal.Count)
                    frmPrev.Enabled = True
                    MoveFocus frmPrev
                    Set frmPrev = Nothing
                End If
                ShowFormLogToFile "-1" & vbTab & strLogMsg
            End If
        End If
    End If
    
End Sub

'When minimizing an MDI parent, call this in order
'to hide any non-child non-modal forms that may
'still be showing underneath.
Public Sub HideAllNonChildForms(Optional ByVal strExceptions$ = "")
    Dim frm As Form, i&
    On Error Resume Next
    'look through all loaded forms
    strExceptions = vbTab & Trim(UCase(strExceptions)) & vbTab
    For i = Forms.Count - 1 To 0 Step -1
        Set frm = Forms(i)
        With frm
            'if not an MDI child or MDI parent
            If Not IsMDIChild(frm) Then
                If Not TypeOf frm Is MDIForm Then '(not MDI parent)
                    If InStr(strExceptions, vbTab & UCase(frm.Name) & vbTab) = 0 Then
                        If .Visible Then
                            .Visible = False 'hide it
                        End If
                    End If
                End If
            End If
        End With
    Next
    Set frm = Nothing
End Sub

'Sets form placement from string returned by GetFormPlacement
'(string containing Placement, windowState, Visible)
Public Sub SetFormPlacement(frm As Form, ByVal strFormSize As String, _
        Optional ByVal strToDo = "P", Optional ByVal bAllowNegatives As Boolean = True)

    Dim s$, nLeft#, nTop#, nWidth#, nHeight#, nWindowState#, nVisible#
    
    On Error Resume Next
    nLeft = NULL_DATA: nTop = NULL_DATA: nWidth = NULL_DATA: nHeight = NULL_DATA
    nWindowState = -1
    nVisible = -9
        
    'parse placement #'s from string
    strFormSize = StripStr(strFormSize, " ")
    If Len(strFormSize) = 0 Then Exit Sub
    s = Parse(strFormSize, ";", 1)
    If Len(s) > 0 Then
        nLeft = Val(s)
        If nLeft < 0 And Not bAllowNegatives Then
            nLeft = 0
        End If
    End If
    s = Parse(strFormSize, ";", 2)
    If Len(s) > 0 Then
        nTop = Val(s)
        If nTop < 0 And Not bAllowNegatives Then
            nTop = 0
        End If
    End If
    s = Parse(strFormSize, ";", 3)
    If Len(s) > 0 Then nWidth = Val(s)
    s = Parse(strFormSize, ";", 4)
    If Len(s) > 0 Then nHeight = Val(s)
    s = Parse(strFormSize, ";", 5)
    If Len(s) > 0 Then nWindowState = Int(Val(s))
    s = Parse(strFormSize, ";", 6)
    If Len(s) > 0 Then nVisible = Int(Val(s))
    
    'fix sizes which we're not doing or are invalid
    strToDo = UCase(Trim(strToDo))
    If InStr(strToDo, "P") > 0 Then strToDo = strToDo & "LTWH"
    If InStr(strToDo, "L") = 0 Or nLeft <= NULL_DATA Then nLeft = frm.Left
    If InStr(strToDo, "T") = 0 Or nTop <= NULL_DATA Then nTop = frm.Top
    If InStr(strToDo, "W") = 0 Or nWidth <= 0 Then nWidth = frm.Width
    If InStr(strToDo, "H") = 0 Or nHeight <= 0 Then nHeight = frm.Height
    
    'more efficient to use "Move" to set all at once
    MoveForm frm, nLeft, nTop, nWidth, nHeight
    
    'set WindowState
    If InStr(strToDo, "S") > 0 And nWindowState >= 0 And nWindowState <= 2 Then
        WindowStateX(frm) = nWindowState
    End If
    
    'set Visibility
    If InStr(strToDo, "V") > 0 And nVisible >= -1 Then
        If nVisible = 0 Then
            frm.Hide
        ElseIf Not frm.Visible Then
            frm.Show
        End If
    End If

End Sub

'Gets form placement: string containing Left, Top, Width,
'  Height, windowState, Visible (to store in INI file, etc)
Public Function GetFormPlacement(frm As Form) As String
    GetFormPlacement = Str(frm.Left) & ";" & Str(frm.Top) _
        & ";" & Str(frm.Width) & ";" & Str(frm.Height) & ";" _
        & Str(WindowStateX(frm)) & ";" & Str(CInt(frm.Visible))
End Function

'Call this from form's Resize event as follows ...
'  If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) then Exit Sub
Public Function LimitFormSize(frm As Form, ByVal lMinScaleWidth As Long, ByVal lMinScaleHeight As Long, _
        Optional lMaxWidth As Long = 0, Optional lMaxHeight As Long = 0) As Boolean
        
    Dim lWidth&, lHeight&, lMinWidth&, lMinHeight&
    Dim tppWidth#, tppHeight#, bResize As Boolean
    
    Select Case WindowStateX(frm)
    'If form is maximized, just return false now
    Case wsMaximized
        LimitFormSize = False
        Exit Function
    'If form is minimized, then just return true right
    'now so the form's resize event will be exited
    'before all the code for resizing the controls.
    Case wsMinimized
        LimitFormSize = True
        Exit Function
    End Select
    
    'first round parms to nearest pixel in twips
    tppWidth = Screen.TwipsPerPixelX
    tppHeight = Screen.TwipsPerPixelY
    lMinScaleWidth = Int(lMinScaleWidth / tppWidth + 0.5) * tppWidth
    lMinScaleHeight = Int(lMinScaleHeight / tppHeight + 0.5) * tppHeight
    lMaxWidth = Int(lMaxWidth / tppWidth + 0.5) * tppWidth
    lMaxHeight = Int(lMaxHeight / tppHeight + 0.5) * tppHeight
    
    'convert min "scale" (inner) sizes to form (outer) sizes
    lWidth = frm.Width
    lHeight = frm.Height
    lMinWidth = lMinScaleWidth + (lWidth - frm.ScaleWidth)
    lMinHeight = lMinScaleHeight + (lHeight - frm.ScaleHeight)
    
    'check if width is outside limits
    If lWidth < lMinWidth Then
        lWidth = lMinWidth
        bResize = True
    ElseIf lWidth > lMaxWidth And lMaxWidth > 0 And lMaxWidth >= lMinWidth Then
        lWidth = lMaxWidth
        bResize = True
    End If
        
    'check if height is outside limits
    If lHeight < lMinHeight Then
        lHeight = lMinHeight
        bResize = True
    ElseIf lHeight > lMaxHeight And lMaxHeight > 0 And lMaxHeight >= lMinHeight Then
        lHeight = lMaxHeight
        bResize = True
    End If
        
    'if need to resize, more efficient to do both at same time
    If bResize Then
        frm.Move frm.Left, frm.Top, lWidth, lHeight
        LimitFormSize = True
    Else
        LimitFormSize = False
    End If
End Function

' Returns true if specified form is loaded
' (pass the form name, e.g. "frmMain")
Public Function FormIsLoaded(ByVal strFormName$) As Boolean

    Dim i&
    On Error Resume Next
    strFormName = UCase(strFormName)
    For i = 0 To Forms.Count - 1
        If UCase(Forms(i).Name) = strFormName Then
            FormIsLoaded = True
            Exit Function
        End If
    Next

End Function

Public Sub EnableContainer(ByVal plHwnd As Long, ByVal pbEnabled As Boolean)
On Error Resume Next

    Static iCount As Integer
    Dim lHwnd As Long
    
    lHwnd = GetWindow(plHwnd, GW_CHILD)
    If lHwnd = 0 Then Exit Sub
    
    Do While lHwnd <> 0&
        EnableContainer lHwnd, pbEnabled
        EnableWindow lHwnd, pbEnabled
        iCount = iCount + 1
        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
    Loop
    
    EnableWindow plHwnd, pbEnabled

End Sub

' Callback used by API calls in GetWindowHandles function
Private Function EnumWindowsCallback(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim nNumItems&
    ' get next item to store window handle
    nNumItems = EnumWindowHandles(0) + 1
    If nNumItems > UBound(EnumWindowHandles) Then
        ReDim Preserve EnumWindowHandles(nNumItems * 2) As Long
    End If
    ' store window handle at item# and store # items
    EnumWindowHandles(nNumItems) = hWnd
    EnumWindowHandles(0) = nNumItems
    EnumWindowsCallback = True
End Function

' To get array of window handles (returned array is 1-based)
' - if hWndParent = 0, will return array of all top-level window handles
' - if hWndParent <> 0, will return array of all child window handles (including their descendents)
Public Sub GetWindowHandles(aHwnd() As Long, Optional ByVal hWndParent& = 0)
    
    Dim i&
    ' reinitialize array of handles used by callback function
    ReDim EnumWindowHandles(0) As Long
    EnumWindowHandles(0) = 0
    ' call appropriate EnumWindows API
    If hWndParent = 0 Then
        i = EnumWindows(AddressOf EnumWindowsCallback, 0)
    Else
        i = EnumChildWindows(hWndParent, AddressOf EnumWindowsCallback, 0)
    End If
    ' copy handles to array to return
    ReDim aHwnd(EnumWindowHandles(0)) As Long
    aHwnd(0) = hWndParent
    For i = 1 To UBound(aHwnd)
        aHwnd(i) = EnumWindowHandles(i)
    Next
    ReDim EnumWindowHandles(0) As Long

End Sub

' To get array of Process ID's (returned array is 1-based)
Public Sub GetProcessIDs(aProcessIDs() As Long)
    
    Dim cb As Long, cbNeeded As Long, iNumElements As Long
    
    ' not supported for Win95, 98 or ME
    If Not Is9598orMe Then
        'Get the array containing the process id's for each process object
        cb = 8 * 32
        cbNeeded = 96 * 32
        Do While cb <= cbNeeded
            cb = cb * 2
            ReDim aProcessIDs(cb / 4) As Long
            EnumProcesses aProcessIDs(1), cb, cbNeeded
        Loop
    End If
    ReDim Preserve aProcessIDs(cbNeeded / 4) As Long

End Sub

' To get the name of a process
Public Function GetProcessName(ByVal nProcessID As Long, Optional ByVal bWithPath As Boolean = False) As String
    
    Dim cbNeeded As Long, nSize As Long, hProcess As Long
    Dim aModules(1 To 200) As Long
    Dim strModuleName As String
         
    ' not supported for Win95, 98 or ME
    If Not Is9598orMe Then
        'Get a handle to the Process
        hProcess = OpenProcess(1024 Or 16, 0, nProcessID) 'PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified process
            If EnumProcessModules(hProcess, aModules(1), 200, cbNeeded) <> 0 Then
                'Get the ModuleFileName of the process's first module
                strModuleName = Space(500)
                If bWithPath Then
                    nSize = GetModuleFileNameEx(hProcess, aModules(1), strModuleName, 500)
                Else
                    nSize = GetModuleBaseName(hProcess, aModules(1), strModuleName, 500)
                End If
                If nSize > 0 Then
                    GetProcessName = Left(strModuleName, nSize)
                End If
            End If
            'Close the handle to the process
            CloseHandle hProcess
        End If
    End If

End Function

' Returns True if specified key or mouse button is currently pressed
Public Function KeyIsPressed(ByVal vKey As Long) As Boolean
    '(need to ignore least significant bit if just
    ' want to see if it's currently pressed)
    If Abs(CLng(GetAsyncKeyState(vKey))) > 1 Then
        KeyIsPressed = True
    Else
        KeyIsPressed = False
    End If
End Function

' Returns True if any mouse button is currently pressed
Public Function MouseIsPressed(Optional ByVal bOnlyIfPressedSinceLastChecked As Boolean = False) As Boolean
    
    Dim lb&, rb&, mb&, b&
    lb = Abs(CLng(GetAsyncKeyState(VK_LBUTTON)))
    rb = Abs(CLng(GetAsyncKeyState(VK_RBUTTON)))
    mb = Abs(CLng(GetAsyncKeyState(VK_MBUTTON)))
    '(need to ignore least significant bit if just
    ' want to see if it's currently pressed)
    If lb > 1 Then
        b = lb
    ElseIf rb > 1 Then
        b = rb
    ElseIf mb > 1 Then
        b = mb
    End If
    If b > 1 Then
        If Not bOnlyIfPressedSinceLastChecked Then
            MouseIsPressed = True
        ElseIf (b And 1) Then
            MouseIsPressed = True
        End If
    End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FullPathName
'' Description: Converts relative paths to full paths
'' Inputs:      Filename/Path to convert
'' Returns:     Empty string on failure, Full Path on success
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FullPathName(ByVal strFileName As String) As String

    Dim strFullPath As String           ' Full path returned from API call
    Dim strFilePart As String           ' Only the filename from the path
    Dim lSize As Long                   ' Size of the full path
    
    strFullPath = Space(512)
    lSize = GetFullPathName(strFileName, 512, strFullPath, strFilePart)
    
    If lSize <= 0& Or lSize > 512& Then
        FullPathName = ""
    Else
        FullPathName = Left(strFullPath, lSize)
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FileVersion
'' Description: Given a filename, returns the version of the file
'' Inputs:      Filename to check
'' Returns:     Empty string on error, Version otherwise ("x.x.x.x")
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FileVersion(ByVal strFile As String) As String

    Dim lLength     As Long             ' Length of the information block
    Dim lHandle     As Long             ' Handle to the information block
    Dim lPtr        As Long             ' Pointer to the File Information
    Dim lVerSize    As Long             ' Size of the version information
    Dim vi As VS_FIXEDFILEINFO2 ' File information
    Dim strReturn   As String           ' Version of the file
    ' use byte array for binary data instead of string so will work on DBCS machines:
    Dim bytBuffer()   As Byte           ' Buffer for the version info
    
    strReturn = ""
    If FileExist(strFile) Then
        lLength = GetFileVersionInfoSize(strFile, lHandle)
        If lLength > 0 Then
            bytBuffer = String(lLength, 0)
            If GetFileVersionInfo(strFile, lHandle, lLength, bytBuffer(0)) Then
                If VerQueryValue(bytBuffer(0), "\\", lPtr, lVerSize) Then
                    CopyMemory vi, ByVal lPtr, Len(vi)
                    strReturn = vi.dwFileVersionMSh & "." & vi.dwFileVersionMSl & "." _
                              & vi.dwFileVersionLSh & "." & vi.dwFileVersionLSl
                End If
            End If
        End If
    End If

    FileVersion = strReturn

End Function

' To capitalize first letter of each part of path and file
' (each part "fixed" unless is already mixed case)
Public Function FileNameDisplay(ByVal strFile$) As String

    Dim iPart&, iPos&, iAsc&, iCase&
    Dim aParts() As String, strPart$, strExt$
    
    ' check each part of the filename (between backslashes)
    aParts = Split(Trim(strFile), "\")
    For iPart = 0 To UBound(aParts)
        strPart = Trim(aParts(iPart))
        If Len(strPart) > 0 Then
            ' deal with extension separately
            strExt = FileExt(strPart)
            If Len(strExt) > 3 Then strExt = ""
            strPart = Left(strPart, Len(strPart) - Len(strExt))
            ' see if has mixed case
            iCase = 0 '(default: do nothing)
            For iPos = 1 To Len(strPart)
                iAsc = Asc(Mid(strPart, iPos, 1))
                If iAsc >= 65 And iAsc <= 90 Then '(uppercase)
                    If iCase < 0 Then
                        iCase = 0 '(mixed case: do nothing)
                        Exit For
                    Else
                        iCase = 1 '(has uppercase)
                    End If
                ElseIf iAsc >= 97 And iAsc <= 122 Then '(lowercase)
                    If iCase > 0 Then
                        iCase = 0 '(mixed case: do nothing)
                        Exit For
                    Else
                        iCase = -1 '(has lowercase)
                    End If
                End If
            Next
            ' if all upper or all lower, then "fix"
            If iCase <> 0 Then
                ' make only first letter of each word uppercase
                strPart = UCase(Left(strPart, 1)) & LCase(Mid(strPart, 2))
                For iPos = 2 To Len(strPart)
                    '(consider any non-letter and non-digit a word separator)
                    Select Case Asc(Mid(strPart, iPos - 1, 1))
                    'if "0" - "9", "A" to "Z", "a" to "z"
                    Case 48 To 57, 65 To 90, 97 To 122
                        'same word, so leave lowercase
                    Case Else
                        'word delimiter, so make uppercase
                        Mid(strPart, iPos, 1) = UCase(Mid(strPart, iPos, 1))
                    End Select
                Next
            End If
            aParts(iPart) = strPart & LCase(strExt)
        End If
    Next
    FileNameDisplay = Join(aParts, "\")
    
End Function

' To activate and unminimize another program (if currently running)
' - vWindow: either the hWnd, or the app's Title to search for
' - iFindMode: 0 = exact match, 1 = first part matches, 2 = contained in
' - returns True if found and activated
Public Function ActivateOtherProgram(ByVal vWindow As Variant, Optional ByVal iFindMode% = 0) As Boolean

    Dim hWnd&, i&, s$
    Dim aHwnd() As Long
    
    'if passed the window title, then find the window handle
    If VarType(vWindow) = vbString Then
        vWindow = UCase(vWindow)
        GetWindowHandles aHwnd
        For i = 1 To UBound(aHwnd)
            s = UCase(vbGetWindowText(aHwnd(i), , False))
            Select Case iFindMode
            Case 1 'First part of title matches
                If Left(s, Len(vWindow)) = vWindow Then hWnd = aHwnd(i)
            Case 2 'Contained within title
                If InStr(s, vWindow) > 0 Then hWnd = aHwnd(i)
            Case Else 'Exact match
                If s = vWindow Then hWnd = aHwnd(i)
            End Select
            If hWnd <> 0 Then Exit For
        Next
    Else '(the window handle was passed)
        hWnd = CLng(vWindow)
    End If
    
    'see if a valid window
    If IsWindow(hWnd) <> 0 Then
        'if minimized, then restore to prior state
        If IsIconic(hWnd) Then ShowWindow hWnd, SW_OTHERUNZOOM
        'and activate it
        SetForegroundWindow hWnd
        ActivateOtherProgram = True
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetupGrid
'' Description: Sets up a grid generically with our standard settings in either
''              a grid mode, a list box style mode, or a tree mode
'' Inputs:      VSFlexGrid to set up, Mode to set it up in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetupGrid(FlexGrid As Control, ByVal eMode As eGridMode)

    Dim iSaveRedraw&
    With FlexGrid
        iSaveRedraw = .Redraw
        .Redraw = 0 'flexRDNone
        
        ' General settings
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = 1 'flexResizeColumns
        .Editable = 0 'flexEDNone
        .ExplorerBar = 7 'flexExSortShowAndMove
        .ExtendLastCol = True
        .Font.Name = CheckSSFont(.Font.Name)
        .ScrollTrack = True
        .SelectionMode = 3 'flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        ' Settings specific to the mode we want to be in
        Select Case eMode
            Case eGridMode_Grid
                .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspace
                .FixedRows = 1
                .GridLines = 1 'flexGridFlat
                .GridLinesFixed = 2 'flexGridInset
            Case eGridMode_List
                .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbWindowBackground
                .FixedRows = 0
                .GridLines = 0 'flexGridNone
                .GridLinesFixed = 0 'flexGridNone
            Case eGridMode_Tree
                .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbWindowBackground
        End Select
        
        .Redraw = iSaveRedraw
    End With

End Sub

' To clear the read-only flag from one or more files
Public Sub ClearReadOnlyFlags(ByVal strFileMask$)
    On Error Resume Next
    Dim strFile$, strPath$, iAttrib%
    strPath = AddSlash(FilePath(strFileMask))
    strFile = Dir(strFileMask, vbReadOnly)
    Do While Len(strFile) > 0
        strFile = strPath & strFile
        iAttrib = GetAttr(strFile)
        If iAttrib And vbReadOnly Then  '(must still check)
            SetAttr strFile, iAttrib - vbReadOnly
        End If
        strFile = Dir
    Loop
End Sub

' to fix problems in standard IsNumeric function
Public Function TextIsNumeric(ByVal strText$) As Boolean
    Dim iPos&
    ' check for scientific notation: 3.000e+005
    strText = UCase(Trim(strText))
    iPos = InStr(strText, "E")
    If iPos > 1 Then
        If IsNumeric(Left(strText, iPos - 1)) And IsNumeric(Mid(strText, iPos + 1)) Then
            TextIsNumeric = True
        End If
    ElseIf Len(strText) <= 1 Then
        If strText >= "0" And strText <= "9" Then
            TextIsNumeric = True
        End If
    ElseIf IsNumeric(strText) Then
        TextIsNumeric = True
    End If
End Function

' To force a call to the form's resize event
Public Sub FormResize(frm As Form)
    Dim lParam As Long
    With frm
        lParam = (.Height \ Screen.TwipsPerPixelY) * (2 ^ 16) _
            + (.Width \ Screen.TwipsPerPixelX)
        PostMessage .hWnd, WM_SIZE, 0, lParam
    End With
End Sub

' To get the percentage of free system resources
' - strMode: ""=least, "S"=system, "G"=gdi, "U"=user
' - returns -1 if unknown (e.g. if 16-bit program not found, or unsupported operating system)
' (Note: NT, 2000, XP always returns 90% or 0% since can grow as needed)
' NOTE: so should only display returned value if >= 0
Public Function GetFreeSystemResources(Optional ByVal strMode$ = "") As Long
    Dim nPerc&
    ' only run this for Win95, Win98, or ME
    nPerc = -1 '(if invalid operating system)
    If Is9598orMe Then
        If Not RunProcess(WinSysPath & "Free_Res.Exe", strMode, True, , nPerc) Then
            nPerc = -1 '(if couldn't run program successfully)
        End If
    End If
    GetFreeSystemResources = nPerc
End Function

Public Function WindowsVersionStr() As String
    
    Static strVersion As String
    
    If Len(strVersion) = 0 Then
        Select Case WindowsVersion
        Case 3.1
            strVersion = "Win32s"
        Case 3.5
            strVersion = "Win95"
        Case 3.8
            strVersion = "Win98"
        Case 3.9
            strVersion = "WinME"
        Case 4#
            strVersion = "WinNT"
        Case 5#
            strVersion = "Win2000"
        Case 5.1
            strVersion = "WinXP"
        Case 5.2
            strVersion = "Win2003"
        Case 6#
            strVersion = "Vista"
        Case 6.1
            strVersion = "Windows7"
        Case 6.2
            strVersion = "Windows8"
        Case 6.3
            strVersion = "Windows8.1"
        Case 6.4 ' Windows10 beta
            strVersion = "Windows10"
        Case Else ' Windows10 release (and thereafter?)
            strVersion = "Windows" & Format(WindowsVersion, "0.0##")
            If Right(strVersion, 2) = ".0" Then
                strVersion = Left(strVersion, Len(strVersion) - 2)
            End If
        End Select
    End If
    WindowsVersionStr = strVersion
    
End Function

' Returns Windows OS version (Major.Minor):
' 3.1 = Win32s on Windows 3.1
' 3.5 = Windows 95
' 3.8 = Windows 98
' 3.9 = Windows ME
' 4.0 = Windows NT
' 5.0 = Windows 2000
' 5.1 = XP
' 5.2 = Windows 2003
' 6.0 = Vista
' 6.1 = Windows 7
' 6.2 = Windows 8
Public Function WindowsVersion() As Double

    Static dVersion As Double
    Dim i&, d#, s$
    
    If dVersion = 0 Then
        Dim OsVersion As OSVERSIONINFO
        OsVersion.dwOSVersionInfoSize = Len(OsVersion)
        If GetVersionEx(OsVersion) <> 0 Then
            ' NOTE: version major/minor used differently for pre-NT vs. post-NT
            Select Case OsVersion.dwPlatformId
            Case VER_PLATFORM_WIN32s
                dVersion = 3.1 ' Win32s on Windows 3.1
            Case VER_PLATFORM_WIN32_WINDOWS
                If OsVersion.dwMinorVersion >= 90 Then
                    dVersion = 3.9 ' Windows ME
                ElseIf OsVersion.dwMinorVersion >= 10 Then
                    dVersion = 3.8 ' Windows 98
                Else
                    dVersion = 3.5 ' Windows 95
                End If
            Case Else  ' NT/2000 and later
                dVersion = OsVersion.dwMajorVersion
                If OsVersion.dwMinorVersion < 10 Then
                    dVersion = dVersion + OsVersion.dwMinorVersion / 10#
                ElseIf OsVersion.dwMinorVersion < 100 Then
                    dVersion = dVersion + OsVersion.dwMinorVersion / 100#
                Else
                    dVersion = dVersion + OsVersion.dwMinorVersion / 1000#
                End If
            End Select
        End If
        ' TLB 6/10/2015: according to Microsoft's strange backwards-compatibility design,
        ' their GetVersionEX will now never return anything greater than 6.2 (Windows 8.0)
        ' for any program not specifically designed for newer OS's (e.g. per a manifest).
        ' SO, we will now also start checking the version of the Kernel ...
        If dVersion = 0 Or dVersion > 5 Then
            s = UCase(Trim(FileVersion(AddSlash(WinSysPath) & "kernel32.dll")))
            i = InStr(s, ".")
            If i > 0 Then
                i = InStr(i + 1, s, ".")
                If i > 0 Then
                    s = Left(s, i - 1)
                    d = Val(s)
                    If d > dVersion Then
                        dVersion = d
                    End If
                End If
            End If
            ' TLB 1/15/2016: the above ended up not always working since it also seems to
            ' always return 6.2 (Windows 8.0) for all the system Dll's on Windows 10.
            ' So we will now try to get the new registry keys (which didn't exist prior to Windows 10).
            #If TRADENAV_EXE Then
                s = "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
                d = GetRegistryValue(rkLocalMachine, s, "CurrentMajorVersionNumber", 0)
                i = GetRegistryValue(rkLocalMachine, s, "CurrentMinorVersionNumber", 0)
                If d >= 8 And i >= 0 Then
                    If i < 10 Then
                        d = d + i / 10#
                    ElseIf i < 100 Then
                        d = d + i / 100#
                    Else
                        d = d + i / 1000#
                    End If
                    If d > dVersion Then
                        dVersion = d
                    End If
                End If
            #End If
        End If
    End If
    WindowsVersion = dVersion

End Function

Public Function Is9598orMe() As Boolean

    If WindowsVersion > 0 And WindowsVersion < 4 Then
        Is9598orMe = True
    End If

End Function

Public Function IsAtLeastVista() As Boolean

    If WindowsVersion >= 6 Then
        IsAtLeastVista = True
    End If

End Function

Public Function IsAtLeastXP() As Boolean

    If WindowsVersion >= 5.1 Then
        IsAtLeastXP = True
    End If

End Function

Public Sub SetLowFragHeap()

    Dim i As Long
    On Error Resume Next
    ' can only run this for XP and higher
    If IsAtLeastXP Then
        i = 2
        If HeapSetInformation(GetProcessHeap, 0, i, Len(i)) = 0 Then
            i = GetLastError
        End If
    End If

End Sub

' To create a delimited string of the font properties
' (for storing in registry, INI file, etc.)
Public Function FontToString(Font As StdFont) As String

    FontToString = Font.Name & "|" & Str(Font.Size) & "|" _
            & Str(Font.Bold) & "|" & Str(Font.Italic) & "|" _
            & Str(Font.Underline) & "|" & Str(Font.Strikethrough)
    
End Function

' To set font properties from a delimited string
' (to be used with FontToString)
Public Function FontFromString(Font As StdFont, ByVal strFont As String) As Boolean

    Dim strName$, nSize&
    
    On Error GoTo FontExit
    strName = Trim(Parse(strFont, "|", 1))
    nSize = Val(Parse(strFont, "|", 2))
    If Len(strName) > 0 And nSize > 0 Then
        Font.Name = CheckSSFont(strName)
        Font.Size = nSize
        Font.Bold = Val(Parse(strFont, "|", 3))
        Font.Italic = Val(Parse(strFont, "|", 4))
        Font.Underline = Val(Parse(strFont, "|", 5))
        Font.Strikethrough = Val(Parse(strFont, "|", 6))
        FontFromString = True
    End If
    
FontExit:
    Exit Function
End Function

Public Function IsValidFileBase(ByVal strName$, Optional ByVal bShowErrorMsg As Boolean = True) As Boolean
    
    Dim bValid As Boolean
    
    ' can't start or end with space
    If Len(strName) > 0 And Left(strName, 1) <> " " And Right(strName, 1) <> " " Then
        ' check for invalid characters
        If StripStr(strName, ":\/*?|><" & Chr(34)) = strName Then
            ' make sure name will save properly
            ' (try saving as a temp file)
            strName = TempPath & strName & ".tmp"
            If FileExist(strName) Then
                bValid = True '(must be a valid name)
            Else
                FileFromString strName, " " '(save to file)
                If FileExist(strName) Then
                    bValid = True
                    KillFile strName
                End If
            End If
        End If
    End If
    
    If (Not bValid) And bShowErrorMsg Then
        Beep
        InfBox "i=E ; h=Invalid Name ; Invalid characters in name."
    End If

    IsValidFileBase = bValid
End Function

' This function encrypts/decrypts a string (toggles).
' Created 11/24/95 by T.Birch.
'  "Xor" algorithm comes from Crescent's PowerPak Pro,
'  but has been enhanced to include a checksum of the
'  password (so not even close if the password is just
'  one letter off!), to vary which position of password
'  is used (the pattern is unique to each password!),
'  and to alter the password as encryption progresses
'  (to avoid any possible repeating patterns).
'NOTE #1: the VB routine is slower than the DLL, but is
'   more secure (since not passing the key)
'NOTE #2: this method is now OBSOLETE -- we should now
'   be using "gdEncrypt" using G32_GD.DLL
Public Sub VbEncrypt(ByVal strPassword$, strMemory$, _
        Optional ByVal nMemLen& = -1)

    Dim i&, iPos&, nChkSum&, nKeyLen&, nAddr&, iChar&
    Dim iXor As Byte
    Dim aKey() As Byte, aMem() As Byte

    If nMemLen < 0 Or nMemLen > Len(strMemory) Then nMemLen = Len(strMemory)
    nKeyLen = Len(strPassword)
    
    ' build Key array and ChkSum of key
    ReDim aKey(nKeyLen) As Byte
    For i = 0 To nKeyLen - 1
        iChar = Asc(Mid(strPassword, i + 1, 1))
        ' change space to underscore (in case later passed as command arg)
        If iChar = 32 Then iChar = Asc("_")
        aKey(i) = iChar
        If iChar > 127 Then iChar = iChar - 256
        ' this makes for a non-standard chksum
        If i Mod 2 <> 0 Then
            nChkSum = nChkSum + 2# * iChar
        Else
            nChkSum = nChkSum + iChar
        End If
    Next
    Do While nChkSum < 256
        nChkSum = nChkSum + 256
    Loop
    
    ' encrypt memory
    ReDim aMem(nMemLen) As Byte
    CopyMemory aMem(0), ByVal strMemory, nMemLen
    For i = 0 To nMemLen - 1
        ' determine which character in key to use
        ' (pattern is unique to each key and string)
        iPos = (nMemLen - i + nChkSum) Mod nKeyLen
        iChar = aKey(iPos)
        If iChar > 127 Then iChar = iChar - 256
        iPos = (iChar + nChkSum + nMemLen) Mod nKeyLen
        
        ' alter password as encryption progresses
        ' (based on ChkSum and string+key length)
        iXor = (aKey(iPos) + nChkSum + nMemLen + nKeyLen) Mod 256
        aKey(iPos) = iXor
        
        ' now Xor the current byte in the string
        aMem(i) = aMem(i) Xor iXor
    Next
    CopyMemory ByVal strMemory, aMem(0), nMemLen

End Sub

' Returns true if successfully renamed the file
Public Function RenameFile(ByVal strOldName$, ByVal strNewName$) As Boolean

    On Error GoTo RenameError
    If Len(strOldName) = 0 Or Len(strNewName) = 0 Then GoTo RenameError
    Name strOldName As strNewName
    RenameFile = True
    Exit Function

RenameError:
    RenameFile = False
    Exit Function
End Function

' Used for VB error-handling -- use the following template (can use "default mode"
' as long as only the event names have an underscore in them) ...
'   Private Sub/Function/Property RoutineName (Args)
'   On Error GoTo ErrSection:
'       ... (Code) ...
'       Exit Sub/Function/Property
'   ErrSection:
'       RaiseError Me.Name & ".RoutineName"
'   End Sub/Function/Property
Public Function RaiseError(Optional ByVal strErrSource$ = "", _
                    Optional ByVal Mode As eGDRaiseErrorMode = eGDRaiseError_Default, _
                    Optional ByVal strPath As String = "") As Boolean
                    
    Static lNumber As Long              ' Error number
    Static strSource As String          ' Error source
    Static strDesc As String            ' Error description
    Static strCallStack As String       ' Call stack (Sources stringed together)
    Static strFileName As String        ' File name of the output file
    Static dTimeRaised As Double        ' Time the error was initially raised
    Static bHasHadErrors As Boolean        ' True if any errors were handled
    Dim fh As Integer                   ' File handle for output file
    Dim iPos As Integer
    
    If Mode = eGDRaiseError_HasHadErrors Then
        RaiseError = bHasHadErrors
        Exit Function
    ElseIf Mode <> eGDRaiseError_Init Then
        bHasHadErrors = True
    End If
    
    If Mode = eGDRaiseError_Default Then
        ' Default: if an underscore after the last period then do a "show", else do a "raise"
        Mode = eGDRaiseError_Raise
        If InStr(UCase(strErrSource), ".CLASS_") = 0 Then
            For iPos = Len(strErrSource) To 1 Step -1
                If Mid(strErrSource, iPos, 1) = "_" Then
                    Mode = eGDRaiseError_Show
                    Exit For
                ElseIf Mid(strErrSource, iPos, 1) = "." Then
                    Exit For
                End If
            Next
        End If
    End If
    
    ' If no current error set, set the current error
    If (lNumber = 0 Or (Now > (dTimeRaised + 3 / 86400))) And Mode <> eGDRaiseError_Init Then
        lNumber = Err.Number
        If strErrSource = "" Then
            strSource = App.EXEName
        Else
            strSource = strErrSource
        End If
        strDesc = Err.Description
        dTimeRaised = Now
    End If
    
    ' If not showing the message, re-raise the error...
    If Mode = eGDRaiseError_Raise Then
        strCallStack = strCallStack & "," & strErrSource
        Err.Raise lNumber, strSource, strDesc
    Else
        If Mode = eGDRaiseError_Show Then
            strCallStack = strCallStack & "," & strErrSource
            Replace strDesc, vbCrLf, "|"
            
            ' Strip off leading comma if it has one...
            If Left(strCallStack, 1) = "," Then strCallStack = Mid(strCallStack, 2)
            
            ' Re-Initialize the logging file if necessary
            If Len(strFileName) = 0 Then
                If strPath = "" Then
                    strFileName = AddSlash(App.Path) & "Errors.LOG"
                Else
                    strFileName = AddSlash(strPath) & "Errors.LOG"
                End If
                KillFile strFileName
            End If

            ' Log the file to the Errors.LOG file...
            On Error Resume Next
            fh = FreeFile
            Open strFileName For Append Shared As #fh
            If fh Then
                Print #fh, Format(dTimeRaised, "hh:mm:ss") & vbTab;
                Print #fh, "Error: " & Str(lNumber) & vbTab & strDesc
                Print #fh, strCallStack
                Print #fh, "==========================="
                Close #fh
            End If
            
            ' Show the message to the user appropriately...
            If lNumber < 0 Then
                InfBox strDesc, , , "Error", , , , , , , , eGDAlign_Left
            Else
                InfBox "An unexpected error occurred.||Please report the following: " & _
                    "|Source:  " & strSource & _
                    "|Message: " & strDesc, , , "Error", , , , , , , , eGDAlign_Left
            End If
        End If
            
        ' Reset the static variables...
        lNumber = 0
        strSource = ""
        strDesc = ""
        strCallStack = ""
        dTimeRaised = 0
    End If

End Function

' primarily called from toolbar click in order to force the
' "Lost_Focus" to get called from currently active control
' (must pass a ToggleTo control that is always visible)
Public Sub ToggleFocus(frm As Form, ctlToggleTo As Control)

    On Error Resume Next
    Dim ctl As Control
    Set ctl = frm.ActiveControl
    If Not ctl Is Nothing Then
        MoveFocus ctlToggleTo
        MoveFocus ctl
        DoEvents
    End If

End Sub

' returns True if running from the Visual Basic IDE (i.e. not from compiled EXE)
Public Function IsIDE() As Boolean


    On Error GoTo ErrSection
    Debug.Print 3 / 0 ' (this will error if in IDE)
    Exit Function

ErrSection:
    IsIDE = True
    Resume Next
End Function

' returns True if running the DBCS version of Windows
' (usually only sold in China, Japan and Korea)
' - FYI: VB strings with binary data are interpreted differently
'   under the double-byte-character-set!
Public Function IsDBCS() As Boolean

    Static iIsDBCS As Integer
    If iIsDBCS = 0 Then
        If Asc(Chr(130)) <> 130 Then
            iIsDBCS = 1 'DBCS
        Else
            iIsDBCS = -1 'Unicode
        End If
    End If
    If iIsDBCS > 0 Then
        IsDBCS = True
    Else
        IsDBCS = False
    End If
    
End Function

' To build an "INI" type of string consistently (for various data types)
' - numbers, dates and booleans will be first converted to a double
'   (stored more consistently regardless of regional settings,
'    so "." will always be the decimal point when stored)
Public Function IniString(ByVal strProperty$, ByVal vValue As Variant) As String
On Error GoTo ErrSection:

    If Right(strProperty, 1) <> "=" Then
        strProperty = Trim(strProperty) & "="
    End If
    IniString = strProperty & Str(vValue)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IniString"
End Function

' To parse an "INI" type of string ("Prop=Value")
' - hands back the value as both a string and a number
' - regional settings will be ignored for converting the number
'   (so "." will always be interpreted as a decimal point)
Public Function ParseIniString(strProperty$, strValue$, dValue#) As Boolean
On Error GoTo ErrSection:

    Dim iPos&
    
    iPos = InStr(strProperty, "=")
    If iPos > 0 Then
        strValue = Trim(Mid(strProperty, iPos + 1))
        strProperty = Trim(Left(strProperty, iPos - 1))
        dValue = ValOfText(strValue, False)
        ParseIniString = True
    Else
        ParseIniString = False
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.ParseIniString"
End Function

' Converts a number, date, or boolean from a string (esp. from a text file)
' - always use this when converting from a string in a text file (so will
'       recognize "." as the decimal point regardless of regional settings)
' - this is an override to VB's "Val" function in order to recognize "True"
Public Function Val(ByVal strText As Variant) As Double

    On Error Resume Next '(so will return 0 in case of overflow)
    If VarType(strText) <> vbString Then
        Val = CDbl(strText) '(if not a string, just convert to a double)
    ' need to be backward-compatible with any booleans that had been written
    ' to a text file using the "CStr" function
    ElseIf UCase(Left(Trim(strText), 4)) = "TRUE" Then
        Val = True
    Else ' Val will always interpret a "." as the decimal point
        Val = VBA.Val(strText)
    End If

End Function

' Formats a number, date, or boolean (esp. for storing in a text file)
' - always use this instead of "CStr" when storing to a text file (so will
'       store "." as the decimal point regardless of regional settings)
' - this is an override to VB's "Str" function in order to trim the string
Public Function Str(ByVal vValue As Variant) As String

    On Error Resume Next
    If VarType(vValue) = vbString Then
        Str = vValue '(if already a string, just return it)
    Else
        ' need Dates and Booleans first converted to a Double,
        ' then use "Str" (instead of "CStr" since we want a decimal
        ' point instead of a comma regardless of regional settings)
        Str = Trim(VBA.Str(CDbl(vValue)))
    End If
    
End Function

#If 0 Then

' Converts a number, date, or boolean from a string (esp. from a text file)
' - always use this when converting from a string in a text file (so will
'       recognize "." as the decimal point regardless of regional settings)
' - this is an override to VB's "Val" function in order to recognize "True"
'       and to allow the option for using regional settings (i.e. the decimal point)
Public Function Val2(ByVal strText As Variant, Optional ByVal bUseRegionalSettings As Boolean = False) As Double

    On Error Resume Next ' returns 0 in case of errors (e.g. overflow, or an empty string to CDbl)
    If VarType(strText) <> vbString Then
        Val = CDbl(strText) '(if not a string, just convert to a double)
    ' need to be backward-compatible with any booleans that had been written
    ' to a text file using the "CStr" function
    ElseIf UCase(Left(Trim(strText), 4)) = "TRUE" Then
        Val = True
    ElseIf bUseRegionalSettings Then
        Val = CDbl(strText) ' "CDbl" will use regional settings to interpret the decimal point
    Else
        Val = VBA.Val(strText) ' "Val" will always interpret a "." as the decimal point
    End If

End Function


' Formats a number, date, or boolean (esp. for storing in a text file)
' - always use this instead of "CStr" when storing to a text file (so will
'       store "." as the decimal point regardless of regional settings)
' - this is an override to VB's "Str" function in order to trim the string
'       and to allow the option for using regional settings (i.e. the decimal point)
Public Function Str2(ByVal vValue As Variant, Optional ByVal bUseRegionalSettings As Boolean = False) As String

    On Error Resume Next
    If VarType(vValue) = vbString Then
        Str = vValue
    ElseIf bUseRegionalSettings Then
        Str = Trim(CStr(dValue))
    Else
        ' need Dates and Booleans converted to a Double,
        ' then use "Str" (instead of "CStr" since we must use a decimal
        ' point instead of a comma regardless of regional settings)
        Str = Trim(VBA.Str(CDbl(dValue)))
    End If
    
End Function

#End If

' FileCopy which will optionally overwrite a read-only file,
' handle wildcards, and raise a more descriptive error
Public Sub FileCopy(ByVal strSource$, ByVal strDest$, Optional ByVal bOverwriteReadOnly As Boolean = False)
On Error GoTo CopyError

    Dim strSrcPath$, strDestPath$, strFile$, strSrcFile$, strDestFile$, iAttrib%
    
    ' if dest is a folder, make sure ends with a backslash (or if source has wildcards)
    If DirExist(strDest) Or InStr(strSource, "*") > 0 Or InStr(strSource, "?") > 0 Then
        strDest = AddSlash(strDest)
    End If
    strSrcPath = AddSlash(FilePath(strSource))
    strDestPath = AddSlash(FilePath(strDest))
    
    ' copy each matching file
    strFile = Dir(strSource, vbReadOnly)
    Do While Len(strFile) > 0
        strSrcFile = strSrcPath & strFile
        If Right(strDest, 1) = "\" Then
            strDestFile = strDestPath & strFile
        Else
            strDestFile = strDest
        End If
        If bOverwriteReadOnly Then
            If FileExist(strDestFile) Then
                iAttrib = GetAttr(strDestFile)
                If iAttrib And vbReadOnly Then   '(must still check)
                    SetAttr strDestFile, iAttrib - vbReadOnly
                End If
            End If
        End If
        VBA.FileCopy strSrcFile, strDestFile
        
        strFile = Dir
    Loop
    
    Exit Sub

CopyError:
    If Not FileExist(strSource) Then
        Err.Raise vbObjectError + 999, , "Error copying file " & strSrcFile
    Else
        Err.Raise vbObjectError + 999, , "Error copying file to " & strDestFile
    End If
End Sub

'// Returns number of files copied.
'// - strFromSpec: filemask with special search options
'// - strTo: name of destination (path or filename)
'// - Mode:  'C'opy (copy all: overwrite if dest exists)
'//          'E'quate (only if dest is different time/size or does not exist)
'//          'U'pdate (only if dest is older or does not exist)
'//          'V'ersionUpdate (same as Update, but first compares file versions)
'//          'A'dd (only if dest does not exist)
'//          'F'reshen (only if dest exists and is older)
'// - bOverwriteReadOnly: copy even if dest is read-only
'// (Note: if the destination file does not exist, the source will be
'//      copied to a temp file in the dest folder, then be renamed.)
Public Function CopyFiles(ByVal strFromSpec$, ByVal strDest$, Optional ByVal bOverwriteReadOnly As Boolean = False, Optional ByVal strMode$ = "E") As Long

    ' if dest is a folder, make sure ends with a backslash (or if source has wildcards)
    If DirExist(strDest) Or InStr(strFromSpec, "*") > 0 Or InStr(strFromSpec, "?") > 0 Then
        strDest = AddSlash(strDest)
    End If
    
    CopyFiles = gdCopyFiles(strFromSpec, strDest, strMode, bOverwriteReadOnly, True)
    
End Function

'// Returns number of files moved.
'// - strFromSpec: filemask with special search options
'// - strTo: name of destination (path or filename)
'// - Mode: same as for CopyFiles -- default is 'A'dd (only if dest does not exist)
'// - bOverwriteReadOnly: move even if dest is read-only (N/A when Mode = "A")
'// - will first try a "rename" (which should work if dest is same volume)
'// - if "rename" does not work and if the source and dest paths are not
'//      the same, then will do the following: the source will be renamed
'//      to a temp file (so will be immediately "gone"), the temp file will
'//      be copied to the destination, then the temp file will be deleted.
Public Function MoveFiles(ByVal strFromSpec$, ByVal strDest$, Optional ByVal bOverwriteReadOnly As Boolean = False, Optional ByVal strMode$ = "A") As Long

    ' if dest is a folder, make sure ends with a backslash (or if source has wildcards)
    If DirExist(strDest) Or InStr(strFromSpec, "*") > 0 Or InStr(strFromSpec, "?") > 0 Then
        strDest = AddSlash(strDest)
    End If
    
    MoveFiles = gdMoveFiles(strFromSpec, strDest, strMode, bOverwriteReadOnly)
    
End Function

' returns a string with the list of all the available drive letters (e.g. "ACDKLMS")
Public Function GetAllDrives() As String

    Dim strDrives$, iLen&
    strDrives = String(100, " ")
    iLen = gdGetAllDrives(strDrives)
    GetAllDrives = Left(strDrives, iLen)

End Function

' To kill all processes of specified name
' (returns how many processes were killed)
Public Function KillProcess(ByVal strProcess As String, _
            Optional ByVal bJustCountButLeaveRunning As Boolean = False) As Long

    Dim i&, hWnd&, strWindow$, lPID&, hProcess&, rc&, iCount&
    Dim aHwnd() As Long
    
    ' look through all the top-level window handles
    GetWindowHandles aHwnd
    strProcess = UCase(Trim(strProcess))
    For i = LBound(aHwnd) To UBound(aHwnd)
        hWnd = aHwnd(i)
        If hWnd <> 0 Then
            ' get title of window (but make sure to use GetWindowText instead
            ' of SendMessage so won't lock-up on non-responding processes)
            strWindow = vbGetWindowText(hWnd, , False)
            If UCase(Trim(strWindow)) = strProcess Then
                If bJustCountButLeaveRunning Then
                    iCount = iCount + 1
                Else
                    ' get Process ID for this window
                    lPID = 0
                    rc = GetWindowThreadProcessId(hWnd, lPID)
                    If lPID <> 0 Then
                        ' get a Process handle in order to terminate
                        hProcess = OpenProcess(1, False, lPID)
                        If hProcess <> 0 Then
                            If TerminateProcess(hProcess, 0) <> 0 Then
                                iCount = iCount + 1
                            End If
                            CloseHandle hProcess
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    KillProcess = iCount
End Function

' To format numbers for display purposes ...
' - iDigitsAfterDecimal: >0 to force 'n' digits, <0 for up to 'n' digits (trailing zeroes will be removed)
' - bUseCommas: if true, will use thousands separator for specific regional settings
' - bUseParensIfNegative: if true, will use parentheses instead of a negative sign for negative numbers
Public Function FormatNum(ByVal dValue As Double, _
            Optional ByVal iDigitsAfterDecimal As Integer = -6, _
            Optional ByVal bUseCommas As Boolean = False, _
            Optional ByVal bUseParensIfNegative As Boolean = False) As String

    Dim strValue As String, strFmt As String
    Static strDot As String
    
    On Error Resume Next
    
    ' if first time: get regional settings for the decimal point
    If Len(strDot) = 0 Then
        strDot = Mid(Format(1.5, "0.0"), 2, 1)
    End If
    
    ' build the format string
    If bUseCommas Then
        strFmt = "#,##0"
    Else
        strFmt = "#0"
    End If
    If iDigitsAfterDecimal < 0 Then
        strFmt = strFmt & "." & String(Abs(iDigitsAfterDecimal), "#")
    ElseIf iDigitsAfterDecimal > 0 Then
        strFmt = strFmt & "." & String(iDigitsAfterDecimal, "0")
    End If

    ' get the formatted value
    strValue = Format(dValue, strFmt)
    
    ' strip off the decimal point if it's the last character
    If Right(strValue, 1) = strDot Then
        strValue = Left(strValue, Len(strValue) - 1)
    End If

    ' use parentheses for negative number?
    If bUseParensIfNegative And Left(strValue, 1) = "-" Then
        FormatNum = "(" & Mid(strValue, 2) & ")"
    Else
        FormatNum = strValue
    End If

End Function

Public Sub EnableAllForms(ByVal bEnable As Boolean, Optional frmExcept As Form = Nothing)

    Dim i&
    
    On Error Resume Next
    For i = 0 To Forms.Count - 1
        If Not Forms(i) Is frmExcept Then
            Forms(i).Enabled = bEnable
        End If
    Next

End Sub

' Converts the DateTime from one time zone to another.
' TimeZone can be an empty string for local time zone, "GMT" for GMT/UTC, "NY" for New York,
'      "CHI" for Chicago, or it can be a custom time zone specification formatted as follows ...
' - a pipe-delimited string with the GMT offset first, then optional DST rules in
'      chronological order: "GmtOffset|EarliestDstRule|NextDstRule|...|CurrentDstRule"
' - each DST rule is: "FromYear,DstStart,DstEnd,[DstOffset]" (DstOffset defaults to 60)
' - DstStart and DstEnd specifies when to change: Month/DayOfMonth (DayOfMonth uses "SetFromRule"
'      format, e.g. "3/1S" or "3/FS" = first Sunday of March, "10/5S" or "10/LS" = last Sun of Oct)
' - e.g. Sydney = "600|1987,10/LS,3/3S|1989,10/LS,3/1S|1995,10/LS,3/LS|2000,8/LS,3/LS|2001,10/LS,3/LS"
'      Chicago = "-360|1966,4/LS,10/LS|1974,1/1S,10/LS|1975,2/LS,10/LS|1976,4/LS,10/LS|1987,4/1S,10/LS"
Public Function ConvertTimeZone(Optional ByVal dDateTime# = 0, Optional ByVal strFromTimeZone$ = "", _
            Optional ByVal strToTimeZone$ = "NY") As Double
    
    If dDateTime = 0 Then dDateTime = Now
    ConvertTimeZone = gdConvertTimeZone(dDateTime, strFromTimeZone, strToTimeZone)
    
End Function

' To play a sound file (e.g. *.WAV)
' - will return immediately after starting the sound unless pass bWaitUntilFinished
' - can optionally make the sound loop until stopped
' - pass an empty string to stop the current sound
Public Sub PlaySoundFile(Optional ByVal strFile$ = "", Optional ByVal bWaitUntilFinished As Boolean = False, Optional ByVal bLoopUntilStopped As Boolean = False)

    ' TLB 1/30/2012: if specified filename is relative to the app path (e.g. ".\blah" or "..\blah"),
    ' then set the current directory to the app path (just in case it had gotten changed somehow)
    If Left(strFile, 1) = "." Then
        ChangePath App.Path
    End If

    If Len(Trim(strFile)) = 0 Then
        PlaySound ByVal 0&, 0, 0 ' pass a NULL pointer in order to stop current sound
    ElseIf bLoopUntilStopped Then
        PlaySound strFile, 0, &H20000 Or &H1 Or &H8 ' pass LOOP and ASYNC flags
    ElseIf bWaitUntilFinished Then
        PlaySound strFile, 0, &H20000 ' just pass FILENAME flag
    Else
        PlaySound strFile, 0, &H20000 Or &H1 ' pass ASYNC flag to immediately return
    End If

End Sub

' If passed font name is Sans Serif, replaces with the desired default
' (use TrueType version if exists on machine, else use the non-TrueType version)
Public Function CheckSSFont(Optional ByVal strFontName$ = "MS Sans Serif") As String

    Dim Font As StdFont, strFile$
    Static strSansSerifFont$
    
    On Error Resume Next
    If Len(strSansSerifFont) = 0 Then
        ' first time: see if TrueType version of Sans Serif exists on this machine
        Set Font = New StdFont
        strSansSerifFont = "Microsoft Sans Serif"
        Font.Name = strSansSerifFont
        If UCase(Font.Name) <> UCase(strSansSerifFont) Then
            ' if font file exists, try adding it as a resource and try again
            strFile = WindowsPath & "Fonts\micross.ttf"
            If FileExist(strFile) Then
                ' even if doesn't work for this instance, the AddFontResource seems to
                ' add the font to the registry so will work the next time the app is run
                AddFontResource strFile
                SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0&, ByVal 0&
                Font.Name = strSansSerifFont
            End If
            If UCase(Font.Name) <> UCase(strSansSerifFont) Then
                strSansSerifFont = "MS Sans Serif" ' else use the default (non-TrueType)
            End If
        End If
'strSansSerifFont = "MS Sans Serif"
        Set Font = Nothing
    End If
    
    If UCase(strFontName) = "MS SANS SERIF" Or UCase(strFontName) = "MICROSOFT SANS SERIF" Then
        ' use default Sans Serif font (hopefully the TrueType version)
        CheckSSFont = strSansSerifFont
    Else ' else return what was passed
        CheckSSFont = strFontName
    End If

End Function

' Fixes the Fonts and BackColors for the controls on one or all forms
' - tries to replace the default "MS Sans Serif" with the TrueType "Microsoft Sans Serif"
'       since this font prints better (esp. in grids)
' - sets all the BackColors to the App's custom BackColor (if it had been set)
Public Sub FixFormControls(Optional IfJustThisForm As Form = Nothing, Optional ByVal gridAltRowColor As Long = 0)
On Error GoTo Trap

    Dim iControl&, iForm&, strOrig$, d#, n&, bJustOneForm As Boolean
    Dim ctl As Control, Font As StdFont, frm As Form
    Dim nWhiteForeColor&
    Dim bSkipCtl As Boolean
    Dim bAlreadyDone As Boolean

    'On Error Resume Next
    
    nWhiteForeColor = RGB(225, 225, 225)
                
    If m_PrevAppBackColor = 0 Then
        m_PrevAppBackColor = -1 '(just set it to something invalid for now)
    End If
    
    ' do either one or all forms
    For iForm = 0 To Forms.Count - 1
        If IfJustThisForm Is Nothing Then
            Set frm = Forms(iForm)
        Else
            Set frm = IfJustThisForm
        End If
        
        'JM 10-27-2015: some control color info
        'default background color for most controls on a form appears to be:
        '   &H8000000F: light-gray system constant for Button Face (default for command buttons)
        '   &H8000000A: light-gray system constant for Active Border
        '   these 2 colors are very close, but the code below only process button face color: &H8000000F
        '   so to leave a control untouched set the background color to &H8000000A and vice versa
        '
        'default foreground color for most controls on a form appears to be:
        '   &H80000012: black system constant for Button Text (default for command buttons)
        '   &H80000008: black system constant for Window Text
        '   these 2 are both black, but the code below only process button text color: &H80000012
        '   forecolor is usually the text color for standard VB controls
        '   so to leave the text color untouched set the forecolor to &H80000008 and vice versa
        '
        '&H8000000D system highlight color used on some forms that gets changed to a blue text color
        '&HC00000 is a bright blue used by vsLinked as text color in frmSystemManger (and maybe elsewhere?)
        '
        'some RGBs & constants found for different shades of blue in code
        '   8388608 = rgb(0, 0, 128)
        '   10485760 = rgb(0, 0, 160)
        '
        'bright blue hex color in frmSymbolSelector & frmMonteCarlo: &H00C00000

        
        bAlreadyDone = False
        ' do all the controls on the form
        For iControl = 0 To frm.Controls.Count - 1
            Set ctl = frm.Controls(iControl)
            bSkipCtl = False
            
            ' see if should replace the Font for this control
            Set Font = Nothing
            Set Font = ctl.Font
            If Not Font Is Nothing Then
                strOrig = ""
                strOrig = Font.Name
                If Len(strOrig) > 0 Then
                    ' skip all date controls (spacing gets whacked out -- don't know why)
                    ' and skip combo boxes (causes all text to be highlighted at form load -- don't know why)
                    d = 0
                    If TypeOf ctl Is ctlUniComboImageXP Then
                        d = 1
                    Else
                        d = ctl.Year
                    End If
                    If d = 0 Then
                        Font.Name = CheckSSFont(strOrig)
                        If UCase(Font.Name) <> UCase(strOrig) Then
                            Set ctl.Font = Font '(this needs to be done at least for FlexGrids)
                        End If
                    End If
                End If
            End If
            
            'RH skip this, to old-fashioned
'            If IsAtLeastVista Then
'                If Not IfJustThisForm Is Nothing Then
'                    If Not bAlreadyDone Then
'                        n = SetWindowTheme(frm.hWnd, "", 0)
'                        SendMessage frm.hWnd, WM_THEMECHANGED, 0, 0
'                        bAlreadyDone = True
'                    End If
'                End If
'            End If
            
            If TypeOf ctl Is CommandButton Then
                ' skip command buttons -- since can't change BackColor if .Style = 0,
                ' and .Style is a read-only property -- plus it helps them stand out better
                bSkipCtl = True
            ElseIf (TypeOf ctl Is Frame Or TypeOf ctl Is ctlUniFrameWL) And ctl.Name = "fraLegend" Then
                bSkipCtl = True     'chart tab of performance report
            ElseIf (TypeOf ctl Is Label Or TypeOf ctl Is ctlUniLabelXP) And InStr(ctl.Name, "lblLegend") <> 0 Then
                bSkipCtl = True     'chart tab of performance report
            ElseIf TypeOf ctl Is PictureBox Then
                bSkipCtl = True
'            ElseIf ctl.Name = "txtPreview" Then
'                bSkipCtl = True     'txtPreview is color-coded trade sense
            ElseIf TypeName(frm) = "frmWindowLink" Then
                bSkipCtl = True
            ElseIf TypeName(ctl) = "RichTextBox" Then
                bSkipCtl = True
            ElseIf ctl.Name = "fgTickDistribution" Or ctl.Name = "fgBidDetail" Or ctl.Name = "fgAskDetail" Then
                bSkipCtl = True
            End If
            
            
            ' set BackColor to the App's custom backcolor
            If m_AppBackColor <> 0 And Not bSkipCtl Then
                n = 0
                n = ctl.BackColor
                
                'RH - OVERRIDER THIS
'                If n = &H8000000F Or n = m_PrevAppBackColor Then
'                    ctl.BackColor = m_AppBackColor
'                    n = 0
'                    n = ctl.ForeColor
'
'                    If m_bAppWhiteForeColor Then
'                        If n = &H8000000D Or n = RGB(0, 0, 128) Then
'                            ctl.ForeColor = RGB(224, 224, 224)  'special RGB flag
'                        ElseIf n = &H80000012 Or n = &HC00000 Or n = 0 Then
'                            ctl.ForeColor = nWhiteForeColor
'                        End If
'                    ElseIf n = nWhiteForeColor Then
'                        ctl.ForeColor = &H80000012 'reset
'                    ElseIf n = RGB(224, 224, 224) Then
'                        ctl.ForeColor = &H8000000D 'reset
'                    End If
'                End If
                
                ' some controls have more than one "back color" property
                Select Case UCase(TypeName(ctl))
                Case "VSINDEXTAB"
                    ctl.FrontTabColor = m_AppBackColor
                    ctl.ForeColor = &H80000012
                    n = ctl.FrontTabForeColor
                    
                    If m_bAppWhiteForeColor Then
                        If n = &H8000000D Or n = RGB(0, 0, 128) Then
                            ctl.FrontTabForeColor = RGB(224, 224, 224)  'special RGB flag
                        ElseIf n = &H80000012 Then
                            ctl.FrontTabForeColor = nWhiteForeColor
                        End If
                    ElseIf n = nWhiteForeColor Then
                        ctl.FrontTabForeColor = &H80000012 'reset
                    ElseIf n = RGB(224, 224, 224) Then
                        ctl.FrontTabForeColor = &H8000000D 'reset
                    End If
                    
                Case "VSFLEXGRID"
                    n = 0
                    n = ctl.BackColorFixed
                    If n = &H8000000F Or n = m_PrevAppBackColor Then
                        ctl.BackColorFixed = m_AppBackColor
                        If m_AppBackColor = kDarkThemeColor Or m_AppBackColor = vbWhite Then
                            ctl.BackColor = m_AppBackColor
                            
                            'RH - overridden by Kevins #f2f2f2 ligh grey
                            '''ctl.BackColorBkg = m_AppBackColor
                            ctl.BackColorBkg = g.Styler.GetColor(eGrid_Background)
                            
                            
                            
                        End If
                                            
                        n = 0
                        n = ctl.ForeColorFixed
                        If n = &H80000012 And m_bAppWhiteForeColor Then
                            ctl.ForeColorFixed = nWhiteForeColor
                            ctl.ForeColor = nWhiteForeColor
                        ElseIf n = nWhiteForeColor Then
                            ctl.ForeColorFixed = &H80000012     'reset
                        End If
                    End If

                    n = 0
                    n = ctl.GridColor
                    If n = &H8000000F Or n = m_PrevAppBackColor Then
                        If m_AppBackColor = kDarkThemeColor Then
                            ctl.GridColor = RGB(45, 45, 45)
                        ElseIf m_AppBackColor = vbWhite Then
                            ctl.GridColor = RGB(225, 225, 225)
                        Else
                            ctl.GridColor = m_AppBackColor
                        End If
                    End If
                    ctl.BackColorAlternate = gridAltRowColor

                Case "SSACTIVETOOLBARS"
                    If m_AppBackColor = vbWhite Then
                        ctl.BackColor = vbWhite
                        ctl.ForeColor = kDarkThemeColor
                    ElseIf m_AppBackColor = kDarkThemeColor Then
                        ctl.BackColor = kDarkThemeColor
                        ctl.ForeColor = vbWhite
                    Else
                        ctl.BackColor = m_AppBackColor ' GetSysColor(15)   'COLOR_BTNFACE
                        ctl.ForeColor = &H80000012  'GetSysColor(18)   'COLOR_BTNTEXT ' GetSysColor(9)    'COLOR_CAPTIONTEXT
                    End If

'                Case "RICHTEXTBOX"
'                    If m_AppBackColor = kDarkThemeColor Then
'                        ctl.BackColor = kDarkThemeColor
'                        ctl.SelStart = 0
'                        ctl.SelLength = Len(ctl)
'                        ctl.SelColor = nWhiteForeColor
'                        ctl.SelStart = 0        'reposition cursor to beginning of text
'                    Else
'                        ctl.BackColor = vbWhite
'                        ctl.SelStart = 0
'                        ctl.SelLength = Len(ctl)
'                        ctl.SelColor = 0
'                        ctl.SelStart = 0        'reposition cursor to beginning of text
'                    End If

                End Select
            End If
                        
            ' TLB 6/21/2012: turned out the system's "highlight" color wasn't really the right color
            ' to use as a ForeColor for all themes (e.g. XP Silver), so just change it to dark blue
            n = 0
            n = ctl.ForeColor
            If n = &H8000000D Then
                ctl.ForeColor = RGB(0, 0, 128)
            End If
            If UCase(TypeName(ctl)) = "VSINDEXTAB" Then
                n = 0
                n = ctl.FrontTabForeColor
                If n = &H8000000D Then
                    ctl.FrontTabForeColor = RGB(0, 0, 128)
                End If
            End If
            
            ''Debug.Print ctl.Name, UCase(TypeName(ctl))
            'RH - new Hexagora controls
            With g.Styler
            
                frm.BackColor = .GetColor(eForm_Background)
    
                'Frames
                If TypeOf ctl Is ctlUniFrameWL Then
                    ctl.BackColor = .GetColor(eFrame_Background)
                    
                    'Hexagora added a default caption for frames (the frame's name) so it needs to be removed
                    'also, frames with no captions should also be borderless (many buttons are sitting inside frames, these aren't meant to be visible
                    If ctl.Caption = ctl.Name Or Left$(ctl.Caption, 5) = "Frame" Then
                        ctl.Caption = ""
                    End If
                    
                    If ctl.Caption = "" Then
                        ctl.BorderColor = .GetColor(eFrame_Background)
                    Else
                        ctl.BorderColor = .GetColor(eFrame_Border)
                    End If
                    
                'Radio Buttons
                ElseIf TypeOf ctl Is ctlUniRadioXP Then
                    ctl.BackColor = .GetColor(eFrame_Background)
                    ctl.BorderColor = .GetColor(eFrame_Border)
                    
                     'Fix radio buttons height being a bit too short
                    ctl.Height = ctl.Height + 40
                
                'Labels
                ElseIf TypeOf ctl Is ctlUniLabelXP Then
                    ctl.BackColor = .GetColor(eFrame_Background)
                    ctl.BorderColor = .GetColor(eFrame_Border)
                
                
                'Checkboxes
                ElseIf TypeOf ctl Is ctlUniCheckXP Then
                    'Fix checkboxesheight being a bit too short
                    ctl.Height = ctl.Height + 60
                    ctl.BackColor = .GetColor(eFrame_Background)
                    ctl.BorderColor = .GetColor(eCheck_Border)
                
                
                'Buttons
                ElseIf TypeOf ctl Is ctlUniButtonImageXP Then
                    'button styles are already defaulted
                    ctl.BackColor = .GetColor(eButton_Background)
                    ctl.BorderColor = .GetColor(eButton_Border)
                    ctl.ForeColor = .GetColor(eButton_Text)
                    ctl.Style = iCtlBtnStyle_Flat
                    ctl.RoundedBorders = True
                End If
            End With
            
            
        
        Next iControl
        
        ' do the form's BackColor
        'RH - OVERRIDE THIS
'        If m_AppBackColor <> 0 Then
'            n = 0
'            n = frm.BackColor
'            If n = &H8000000F Or n = m_PrevAppBackColor Then
'                frm.BackColor = m_AppBackColor
'
'                n = 0
'                n = frm.ForeColor
'                If n = &H80000012 And m_bAppWhiteForeColor Then
'                    frm.ForeColor = nWhiteForeColor
'                ElseIf n = nWhiteForeColor Then
'                    ctl.ForeColor = &H80000012      'reset
'                End If
'            End If
'        End If
            
        
        If Not IfJustThisForm Is Nothing Then
            Exit For
        End If
    Next iForm
    
    Set ctl = Nothing
    Set Font = Nothing
    Set frm = Nothing
        
    Exit Sub
    
Trap:
    'Debug.Print Err.Description
    Resume Next
End Sub

' Can set the App's BackColor to a custom color
' (will change all forms and controls which are set to COLOR_BTNFACE,
' -- except for command buttons, combo boxes, and a few others)
Public Sub SetAppBackColor(ByVal nBackColor As Long, Optional ByVal bWhiteForeColor As Boolean = False)

    Dim i&
    On Error Resume Next
    
    ' can pass 0 to clear the custom backcolor
    If nBackColor = 0 Then
        nBackColor = &H8000000F ' COLOR_BTNFACE for Windows = GetSysColor(15)
    ElseIf nBackColor = 1 Or nBackColor = kDarkThemeColor Then
        bWhiteForeColor = True
    End If
    
    If nBackColor <> m_AppBackColor Or bWhiteForeColor <> m_bAppWhiteForeColor Then
        ' set module variables
        m_PrevAppBackColor = m_AppBackColor
        m_AppBackColor = nBackColor
        m_bAppWhiteForeColor = bWhiteForeColor
        ' fix colors for all controls on all forms
        FixFormControls
    End If
    ' reset Prev to an invalid color
    m_PrevAppBackColor = -1
End Sub

Public Function GetAppBackColor() As Long
    If m_AppBackColor = 0 Then
        GetAppBackColor = &H8000000F ' COLOR_BTNFACE for Windows = GetSysColor(15)
    Else
        GetAppBackColor = m_AppBackColor
    End If
End Function

' this allows one window to look active while another is receiving the input
Public Sub FakeActiveLook(ByVal hLookActive As Long, ByVal hRealActive As Long)
    'Make the fake active form look active.
    Call SendMessage(hLookActive, WM_NCACTIVATE, 1, ByVal &H0)
    'Ensure that the keyboard input is directed back to the right window
    Call SendMessage(hRealActive, WM_NCACTIVATE, 1, ByVal &H0)
End Sub

Public Function RoundToSecond(ByVal dDateTime As Double) As Double

    'RoundToSecond = gdFixDateTime(Int(dDateTime * 86400# + 0.5) / 86400#)
    RoundToSecond = gdFixDateTime2(dDateTime, True)
    
End Function

Public Function RoundToMinute(ByVal dDateTime As Double) As Double

    RoundToMinute = gdFixDateTime(Int(dDateTime * 1440# + 0.5) / 1440#)
    
End Function

' Returns number of milliseconds computer has been running
' - works like "GetTickCount" but returns a double
' - is high-resolution -- down to thousandths of a millisecond,
'   (but will use GetTickCount if high-res timer not supported)
' - avoids GetTickCount's 49-day wrap-around problem
' TLB 6/21/2006: the high-res method has issues on multi-processors,
' so let's not use it by default until we can figure something out
' - iUseLowRes: True = low-res, False = auto-detect, 2 = force high-res
Public Function gdTickCount(Optional ByVal iUseLowRes As Integer = True)

    Static dPrevHighRes#, lProcessMask&, lSystemMask&, bOneCPU As Boolean
    ' just the first time, check if there's only 1 CPU
    If lSystemMask = 0 Then
        On Error Resume Next
        If GetProcessAffinityMask(GetCurrentProcess, lProcessMask, lSystemMask) <> 0 Then
            ' 1 CPU involved only if ProcessMask and SystemMask both = 1
            If lProcessMask = 1 And lSystemMask = 1 Then
                bOneCPU = True
            End If
        End If
        lSystemMask = -1 ' (just so won't keep checking)
    End If
    
    ' check if High/Low resolution was specified
    Select Case Abs(iUseLowRes)
    Case 1:
        gdTickCount = gdTickCountVB(1)
    Case 2:
        gdTickCount = gdTickCountVB(0)
    Case Else:
        If bOneCPU Then
            ' can always try the high-res if there's only 1 CPU
            gdTickCount = gdTickCountVB(0)
            ' but if ever returns less than the previous, then high-res can't be used
            If gdTickCount < dPrevHighRes Then
                bOneCPU = False
                gdTickCount = gdTickCountVB(1)
            End If
            dPrevHighRes = gdTickCount
        Else
            gdTickCount = gdTickCountVB(1)
        End If
    End Select
    
End Function

Public Function SetStrToChar(ByVal strToConvert As String, ByVal strChar As String) As String
    If Len(strToConvert) > 0 Then
        SetStrToChar = Replace(Space(Len(strToConvert)), " ", strChar)
    End If
End Function

' this code derived from MSDN example titled "Positioning Objects on a Multiple Display Setup"
' (returns true if window was moved)
Public Function MoveFormOnScreen(frm As Form, Optional ByVal bCenter As Boolean = False, _
                Optional ByVal bOnPrimaryMonitor As Boolean = False) As Boolean

    Dim i&, w&, h&, hWnd&, hMon&, bMoved As Boolean
    Dim wrc As Rect, mrc As Rect
    Dim mi As MONITORINFO
    
    ' this routine is not really valid for MDI Child windows
    On Error Resume Next ' (since .MDIChild property doesn't exist for all forms)
    i = 0
    i = frm.MDIChild
    If i <> 0 Then Exit Function
    
    ' would rather simply not move the window than to bother raising an error
    On Error GoTo ErrSection
    
    ' get the monitor that this window is nearest (or the one it is most on)
    If Not frm Is Nothing Then
        If frm.WindowState = 0 Then
            hWnd = frm.hWnd
        End If
    End If
    If hWnd <> 0 Then
        If GetWindowRect(hWnd, wrc) <> 0 Then
            w = wrc.Right - wrc.Left
            h = wrc.Bottom - wrc.Top
            If bOnPrimaryMonitor Then
                hMon = MonitorFromRect(wrc, MONITOR_DEFAULTTOPRIMARY)
            Else
                hMon = MonitorFromRect(wrc, MONITOR_DEFAULTTONEAREST)
            End If
        End If
    End If
    
    If hMon <> 0 Then
        ' get info for that monitor
        mi.cbSize = Len(mi)
        If GetMonitorInfo(hMon, mi) <> 0 Then
            'mrc = mi.rcMonitor
            mrc = mi.rcWork
            If bCenter Then
                ' center this window in the monitor
                i = mrc.Left + (mrc.Right - mrc.Left - w) / 2
                If wrc.Left <> i Then
                    wrc.Left = i
                    bMoved = True
                End If
                i = mrc.Top + (mrc.Bottom - mrc.Top - h) / 2
                If wrc.Top <> i Then
                    wrc.Top = i
                    bMoved = True
                End If
            Else
                ' make this window be fully visible in that monitor
                i = (wrc.Right + wrc.Left) / 2
                If w > mrc.Right - mrc.Left And i > mrc.Left And i < mrc.Right Then
                    ' if window is wider than monitor and window's mid-X is between
                    ' the monitor's left and right, then leave the left where it is
                Else
                    i = mrc.Right - w
                    If wrc.Left < i Then i = wrc.Left
                    If mrc.Left > i Then i = mrc.Left
                    If wrc.Left <> i Then
                        wrc.Left = i
                        bMoved = True
                    End If
                End If
                i = (wrc.Bottom + wrc.Top) / 2
                If h > mrc.Bottom - mrc.Top And i > mrc.Top And i < mrc.Bottom Then
                    ' if window is taller than monitor and window's mid-Y is between
                    ' the monitor's top and bottom, then leave the top where it is
                Else
                    i = mrc.Bottom - h
                    If wrc.Top < i Then i = wrc.Top
                    If mrc.Top > i Then i = mrc.Top
                    If wrc.Top <> i Then
                        wrc.Top = i
                        bMoved = True
                    End If
                End If
            End If
            If bMoved Then
                ' move the window
                wrc.Right = wrc.Left + w
                wrc.Bottom = wrc.Top + h
                SetWindowPos hWnd, 0, wrc.Left, wrc.Top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
            End If
        End If
    End If
    
    MoveFormOnScreen = bMoved

ErrExit:
    Exit Function
    
ErrSection:
    ' would rather simply not move the window than to bother raising an error
    Resume ErrExit
End Function

' Returns amount of physical RAM
' - pass false for total RAM installed, pass true for RAM currently available
' NOTE: "Available" = TotalPhysical - Committed (which could be negative if well into the swap drive)
' - returns 0 if failed
Public Function PhysicalRAM(Optional ByVal bAvailable As Boolean = False, _
            Optional ByVal bInMegs As Boolean = True) As Double

    Dim MemTemp As Currency ' pseudo-64-bit integer buffer (we need this just so we can get a proper 64 bit storage space)
    Dim ms As MEMORY_STATUS
    Dim msx As MEMORY_STATUS_EX
    Dim dTotal#, dCommitted#

    On Error Resume Next
    If Is9598orMe Then
        ' for Win95,98,ME: must use older version (does not support > 2 gigs)
        ms.dwLength = Len(ms)
        GlobalMemoryStatus ms
        ' Committed = TotalPageFile - AvailPageFile
        dCommitted = CDbl(ms.dwTotalPageFile) - CDbl(ms.dwAvailPageFile)
        ' Total physical RAM
        dTotal = CDbl(ms.dwTotalPhys)
    Else
        ' for WinNT and above: can use newer version (supports > 2 gigs)
        msx.dwLength = Len(msx)
        If GlobalMemoryStatusEx(msx) <> 0 Then
            ' Committed = TotalPageFile - AvailPageFile
            CopyMemory MemTemp, msx.ullTotalPageFile, Len(MemTemp)
            dTotal = Round(CDbl(MemTemp) * 10000)
            CopyMemory MemTemp, msx.ullAvailPageFile, Len(MemTemp)
            dCommitted = dTotal - Round(CDbl(MemTemp) * 10000)
            ' Total physical RAM
            CopyMemory MemTemp, msx.ullTotalPhys, Len(MemTemp)
            dTotal = Round(CDbl(MemTemp) * 10000)
        End If
    End If
    
    If dTotal <= 0 Then
        PhysicalRAM = 0 ' error
    ElseIf bAvailable Then
        ' "Available" = TotalPhysical - Committed
        ' (which could be negative if well into the swap drive)
        PhysicalRAM = dTotal - dCommitted
    Else
        PhysicalRAM = dTotal
    End If
    If bInMegs Then PhysicalRAM = PhysicalRAM / 1048576#

End Function

' Returns a gradient color (partially between 2 other colors)
' - iPercentage must be between 0 and 100
' - typical usage is to return the 16 primary levels of Blue, Red, etc.
'      For i = 1 to 16 : iColor = GradientColor(100 * i / 16, RGB(0, 0, 255)) : Next
' - but can also do any number of levels between any 2 colors
'      For i = 1 to 7 : iColor = GradientColor(100 * i / 7, vbGreen, vbRed) : Next
Public Function GradientColor(ByVal iPercentage&, ByVal iFromColor&, Optional ByVal iToColor& = vbWhite) As Long

    Dim iRed&, iFromRed&, iToRed&, iGreen&, iFromGreen&, iToGreen&, iBlue&, iFromBlue&, iToBlue&
    
    If iPercentage <= 0 Then
        GradientColor = iFromColor
    ElseIf iPercentage >= 100 Then
        GradientColor = iToColor
    Else
        ' split From and To colors into their RGB components
        iFromRed = iFromColor Mod 256
        iFromColor = Int(iFromColor / 256)
        iFromGreen = iFromColor Mod 256
        iFromColor = Int(iFromColor / 256)
        iFromBlue = iFromColor Mod 256
        
        iToRed = iToColor Mod 256
        iToColor = Int(iToColor / 256)
        iToGreen = iToColor Mod 256
        iToColor = Int(iToColor / 256)
        iToBlue = iToColor Mod 256
        
        ' build as a percentage of the "distance" between (for each RGB component)
        iRed = Int(iFromRed + (iToRed - iFromRed) * iPercentage / 100# + 0.5)
        iGreen = Int(iFromGreen + (iToGreen - iFromGreen) * iPercentage / 100# + 0.5)
        iBlue = Int(iFromBlue + (iToBlue - iFromBlue) * iPercentage / 100# + 0.5)
        GradientColor = iRed + iGreen * 256 + iBlue * 256 * 256
    End If
    
End Function

' Returns a HeatMap color for a percentage value
' - pass a number between 0 and 100
' - returns the corresponding color from a standard 5-color heat map (blue, cyan, green, yellow, dark red)
' - can adjust the intensity percentage by using a percentage < 100 (to lighten all the colors)
' - can adjust how "evenly spread" the 5 heat map colors are (e.g. 0 = not evenly spread, which is
'      the "standard" method, but which severly limits the cyan and yellow colors)
Public Function GetHeatMapColor(ByVal dPercent As Double, Optional ByVal iIntensityPercent As Integer = 100, _
            Optional ByVal iEvenlySpreadPercent As Integer = 100) As Long
    
    Dim dFract#, dPower#, nRed&, nGreen&, nBlue&, nRed1&, nGreen1&, nBlue1&, nRed2&, nGreen2&, nBlue2&
    
    ' calculate the power to use based on how evenly spread to make all 5 colors
    If iEvenlySpreadPercent > 0 And iEvenlySpreadPercent <= 100 Then
        dPower = 1 + iEvenlySpreadPercent / 100
    Else
        dPower = 1
    End If
    
    ' based on: 0%=blue, 25%=cyan, 50%=green, 75%=yellow, 90%=red, 100%=brownish-red
    If dPercent <= 0 Then
        ' blue
        nBlue1 = 255
        nBlue2 = 255
        dPercent = 0
        dFract = 0
    ElseIf dPercent < 25 Then
        ' from blue to cyan
        nBlue1 = 255
        nGreen2 = 255
        nBlue2 = 255
        dFract = 1 - ((25 - dPercent) / 25) ^ dPower
    ElseIf dPercent < 50 Then
        ' from cyan to green
        nGreen1 = 255
        nBlue1 = 255
        nGreen2 = 255
        dFract = ((dPercent - 25) / 25) ^ dPower
    ElseIf dPercent < 75 Then
        ' from green to yellow
        nGreen1 = 255
        nRed2 = 255
        nGreen2 = 255
        dFract = 1 - ((75 - dPercent) / 25) ^ dPower
    ElseIf dPercent < 90 Then
        ' from yellow to red
        nRed1 = 255
        nGreen1 = 255
        nRed2 = 255
        dFract = ((dPercent - 75) / 15) ^ dPower
    ElseIf dPercent < 100 Then
        ' from red to brownish-red
        nRed1 = 255
        nRed2 = 192
        dFract = ((dPercent - 90) / 10) ^ dPower
    Else
        ' brownish-red
        nRed1 = 192
        nRed2 = 192
        dPercent = 100
        dFract = 0
    End If

    ' convert percent to a fraction of the distance between the closest two quarter points
    'dFract = dPercent / 25# - Int(dPercent / 25#)  ' for STANDARD method (but not very evenly spread)
    If dFract < 0 Then
        dFract = 0
    ElseIf dFract > 1 Then
        dFract = 1
    End If
    
    ' determine heat map color
    nRed = nRed1 + (nRed2 - nRed1) * dFract
    nGreen = nGreen1 + (nGreen2 - nGreen1) * dFract
    nBlue = nBlue1 + (nBlue2 - nBlue1) * dFract
    If iIntensityPercent > 0 And iIntensityPercent < 100 Then
        nRed = 255 - (255 - nRed) * iIntensityPercent / 100
        nGreen = 255 - (255 - nGreen) * iIntensityPercent / 100
        nBlue = 255 - (255 - nBlue) * iIntensityPercent / 100
    End If
    
    GetHeatMapColor = RGB(nRed, nGreen, nBlue)
    
End Function

Public Function WindowExists(ByVal strTitle As String, Optional ByVal bExactMatch As Boolean = False) As Boolean

    Dim aHwnd() As Long
    Dim lIndex As Long
    Dim strWindowTitle As String
    Dim bReturn As Boolean
    
    bReturn = False
    
    GetWindowHandles aHwnd()
    For lIndex = 1 To UBound(aHwnd)
        strWindowTitle = Trim(vbGetWindowText(aHwnd(lIndex)))
        If bExactMatch Then
            If UCase(strWindowTitle) = UCase(strTitle) Then
                bReturn = True
                Exit For
            End If
        Else
            If InStr(UCase(strWindowTitle), UCase(strTitle)) <> 0 Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    WindowExists = bReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxDouble
'' Description: Given two double values, return the maximum of the two
'' Inputs:      Value 1, Value 2
'' Returns:     Maximum of the two values
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxDouble(ByVal dValue1 As Double, ByVal dValue2 As Double) As Double
    If dValue1 > dValue2 Then
        MaxDouble = dValue1
    Else
        MaxDouble = dValue2
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MinDouble
'' Description: Given two double values, return the minimum of the two
'' Inputs:      Value 1, Value 2
'' Returns:     Minimum of the two values
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MinDouble(ByVal dValue1 As Double, ByVal dValue2 As Double) As Double
    If dValue1 < dValue2 Then
        MinDouble = dValue1
    Else
        MinDouble = dValue2
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MaxLong
'' Description: Given two Long values, return the maximum of the two
'' Inputs:      Value 1, Value 2
'' Returns:     Maximum of the two values
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MaxLong(ByVal lValue1 As Long, ByVal lValue2 As Long) As Long
    If lValue1 > lValue2 Then
        MaxLong = lValue1
    Else
        MaxLong = lValue2
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MinLong
'' Description: Given two Long values, return the minimum of the two
'' Inputs:      Value 1, Value 2
'' Returns:     Minimum of the two values
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MinLong(ByVal lValue1 As Long, ByVal lValue2 As Long) As Long
    If lValue1 < lValue2 Then
        MinLong = lValue1
    Else
        MinLong = lValue2
    End If
End Function

' To allow another App to set the foreground window (i.e. to get the focus)
' - setting hProcessID = -1 allows any App to set the focus
Public Function AllowSetForegroundWindow(Optional ByVal hProcessID As Long = -1&)
    On Error Resume Next ' since some older OS's do not recognize this function
    AllowSetForegroundWindow = AllowSetForegroundWindowAPI(hProcessID)
End Function

' To create a shortcut (.LNK file) on the desktop, start menu, etc.
' - can just pass "Desktop" as path to create a shortcut on the desktop
' - can just use '*' at beginning of path to indicate the StartMenu+Programs
'       (e.g. "*\Genesis" to put it into the Genesis folder within Programs)
Public Function CreateShortcut(ByVal strProgramFile As String, _
        Optional ByVal strShortcutPath As String = "DESKTOP", _
        Optional ByVal strShortcutTitle As String = "", _
        Optional ByVal strTooltip As String = "", _
        Optional ByVal strIconFile As String) As Boolean
    
On Error GoTo ErrSection ' since some older OS's may not recognize this functionality
    Dim strPath$
    Dim VbsObj As Object, MyShortcut As Object

    If Len(strProgramFile) > 0 And Len(strShortcutPath) > 0 Then
        ' build the name of the LNK file
        If UCase(Right(strShortcutPath, 4)) <> ".LNK" Then
            If Len(strShortcutTitle) = 0 Then
                strShortcutTitle = FileBase(strProgramFile)
            End If
            If UCase(strShortcutPath) = "DESKTOP" Then
                strShortcutPath = SpecialFolderPath(CSIDL_ALLUSERS_DESKTOP)
            ElseIf Left(strShortcutPath, 1) = "*" Then
                strShortcutPath = SpecialFolderPath(CSIDL_ALLUSERS_PROGRAMS, True) & Mid(strShortcutPath, 2)
            End If
            strShortcutPath = AddSlash(Trim(strShortcutPath)) & Trim(strShortcutTitle) & ".lnk"
        End If
    
        ' create the LNK file
        Set VbsObj = CreateObject("WScript.Shell")
        Set MyShortcut = VbsObj.CreateShortcut(strShortcutPath)
        MyShortcut.TargetPath = strProgramFile
        MyShortcut.WorkingDirectory = FilePath(strProgramFile)
        'MyShortcut.WindowStyle = 1
        'MyShortcut.Hotkey = "CTRL+SHIFT+F"
        If Len(strIconFile) > 0 Then
            MyShortcut.IconLocation = strIconFile & ", 0" '"notepad.exe, 0"
        Else
            MyShortcut.IconLocation = strProgramFile & ", 0" '"notepad.exe, 0"
        End If
        MyShortcut.Description = strTooltip
        MyShortcut.Save
        If FileExist(strShortcutPath) Then
            CreateShortcut = True
        End If
    End If

ErrExit:
    Set MyShortcut = Nothing
    Set VbsObj = Nothing
    Exit Function
    
ErrSection:
    Resume ErrExit ' would rather simply return false than to bother raising an error
End Function

Public Property Get CheckBoxValue(chk As ctlUniCheckXP) As Boolean 'RH was Checkbox
    CheckBoxValue = (chk.Value = vbChecked)
End Property
Public Property Let CheckBoxValue(chk As ctlUniCheckXP, ByVal bValue As Boolean)
    If bValue = True Then
        If chk.Value <> vbChecked Then '(just so won't do a Click event unless it changed)
            chk.Value = vbChecked
        End If
    ElseIf chk.Value <> vbUnchecked Then
        chk.Value = vbUnchecked
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MonthToCode
'' Description: Determine the month code for the given month
'' Inputs:      Month, ComStock Night Code?
'' Returns:     Month Code
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MonthToCode(ByVal lMonth As Long, Optional ByVal bComStockNightCode As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strMonths As String             ' String of month codes
    Dim strNight As String              ' String of ComStock night month codes
    
    strMonths = "FGHJKMNQUVXZ"
    strNight = "ABCDEILOPRST"
    
    If bComStockNightCode = True Then
        MonthToCode = Mid(strNight, lMonth, 1)
    Else
        MonthToCode = Mid(strMonths, lMonth, 1)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.MonthToCode"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CodeToMonth
'' Description: Determine the month for the given month code
'' Inputs:      Month Code, ComStock Night Code?
'' Returns:     Month
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CodeToMonth(ByVal strMonthCode As String, Optional ByVal bComStockNightCode As Boolean = False) As Long
On Error GoTo ErrSection:

    Dim strMonths As String             ' String of month codes
    Dim strNight As String              ' String of ComStock night month codes
    
    strMonths = "FGHJKMNQUVXZ"
    strNight = "ABCDEILOPRST"
    
    If bComStockNightCode = True Then
        CodeToMonth = InStr(strNight, UCase(strMonthCode))
    Else
        CodeToMonth = InStr(strMonths, UCase(strMonthCode))
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.CodeToMonth"
End Function

' To set the tab stops (column width) for a list box or text box (32 is the standard width)
Public Function SetColumnWidthForControl(ctl As Control, ByVal nColumnWidth As Long) As Boolean
    
    On Error Resume Next
    Dim hWnd As Long, rc&
    
    hWnd = ctl.hWnd
    If hWnd <> 0 Then
        If TypeOf ctl Is ctlUniTextBoxXP Or TypeOf ctl Is TextBox Then 'RH was TextBox
            rc = SendMessage(hWnd, EM_SETTABSTOPS, 1, nColumnWidth)
        ElseIf TypeOf ctl Is ListBox Then
            rc = SendMessage(hWnd, LB_SETTABSTOPS, 1, nColumnWidth)
        End If
    End If
    If rc <> 0 Then
        SetColumnWidthForControl = True
    End If

End Function

' Returns the lowest 16 bits of a 32 bit number
Public Function LoWord(ByVal iValue As Long, ByVal bAsUnsigned As Boolean) As Long

    Dim i&
    CopyMemory i, iValue, 2
    If i >= 32768 And Not bAsUnsigned Then
        i = i - 65536
    End If
    LoWord = i

End Function

' Returns the highest 16 bits of a 32 bit number
Public Function HiWord(ByVal iValue As Long, ByVal bAsUnsigned As Boolean) As Long

    Dim i&, p&
    p = GetAddress(iValue) + 2
    CopyMemory i, ByVal p, 2
    If i >= 32768 And Not bAsUnsigned Then
        i = i - 65536
    End If
    HiWord = i

End Function

' Returns the lowest 8 bits of a number
Public Function LoByte(ByVal iValue As Long, ByVal bAsUnsigned As Boolean) As Long

    Dim i&
    CopyMemory i, iValue, 1
    If i >= 128 And Not bAsUnsigned Then
        i = i - 256
    End If
    LoByte = i

End Function

' Returns the next 8 bits of a number (after the first 8 bits)
Public Function HiByte(ByVal iValue As Long, ByVal bAsUnsigned As Boolean) As Long

    Dim i&, p&
    p = GetAddress(iValue) + 1
    CopyMemory i, ByVal p, 1
    If i >= 128 And Not bAsUnsigned Then
        i = i - 256
    End If
    HiByte = i

End Function

' Calling this periodically will keep the computer from going into sleep/hibernation mode
' (e.g. should call this once a minute while streaming/downloading/distributing)
Public Sub DoNotHibernateNow()

    If IsAtLeastXP Then
        ' passing 1=ES_SYSTEM_REQUIRED resets the system's idle timer
        'SetThreadExecutionState 1 ' ES_SYSTEM_REQUIRED
        SetThreadExecutionState 3 ' ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    End If

End Sub

' to enable/disable the Aero effects in Vista and Windows7
Public Sub EnableAero(ByVal bEnable As Integer)

    Static bEnabled As Boolean
    On Error Resume Next
    If IsAtLeastVista Then
        If bEnable = 2 Then
            ' toggle it from previous state
            bEnabled = Not bEnabled
        ElseIf bEnable = 0 Then
            bEnabled = False
        Else
            bEnabled = True
        End If
        DwmEnableComposition Abs(bEnabled)
    End If
    
End Sub

' returns true if Aero is currently enabled
Public Function AeroIsEnabled() As Boolean
    
    Dim bEnabled As Long
    On Error Resume Next
    If IsAtLeastVista Then
        DwmIsCompositionEnabled bEnabled
        If bEnabled <> 0 Then
            AeroIsEnabled = True
        End If
    End If
    
End Function

' TLB: was trying this out, but can't really get it to work for what we want
Public Sub SetNCRendering(ByVal hWnd&, Optional iValue& = 1)

Exit Sub
    On Error Resume Next
    Dim iAttrib& ', iValue&
    If IsAtLeastVista And (hWnd <> 0) Then
        iAttrib = 2 ' DWMWA_NCRENDERING_POLICY
iAttrib = 4 ' DWMWA_ALLOW_NCPAINT
        'iValue = 1  ' DWMNCRP_DISABLED
        DwmSetWindowAttribute hWnd, iAttrib, iValue, Len(iValue)
    End If

End Sub

Public Function IsLeapYear(ByVal iYear As Integer) As Boolean
    
    Dim bReturn As Boolean              ' Return value for the function
    
    If iYear Mod 400 = 0 Then
        bReturn = True
    ElseIf iYear Mod 100 = 0 Then
        bReturn = False
    ElseIf iYear Mod 4 = 0 Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    IsLeapYear = bReturn

End Function

Public Sub ShowDropDown(cbo As ctlUniComboImageXP)

    SendMessage cbo.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&

End Sub


' to set the Opacity of a window (from 0 = transparent ... to 100 = opaque)
Public Sub SetWindowOpacity(ByVal hWnd As Long, ByVal iPercent As Integer)

    On Error Resume Next
    Dim dwFlags As Long
    
    If iPercent >= 0 And iPercent < 100 Then
        ' convert 0-99 percent to 0-255 range
        iPercent = Round(iPercent * 256# / 100#)
        dwFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
        dwFlags = dwFlags Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, dwFlags
        SetLayeredWindowAttributes hWnd, 0, iPercent, LWA_ALPHA
    Else
        ' turn transparency off
        dwFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
        dwFlags = dwFlags And Not WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, dwFlags
        SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
    End If
    
End Sub

' Turns pixels of this color in a window to Transparent
Public Sub SetWindowTransparencyColor(ByVal hWnd As Long, ByVal iTransparentColor As Long)

    On Error Resume Next
    Dim dwFlags As Long
    
    If iTransparentColor >= 0 Then
        dwFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
        dwFlags = dwFlags Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, dwFlags
        SetLayeredWindowAttributes hWnd, iTransparentColor, 0, LWA_COLORKEY
    Else
        ' turn transparency off
        dwFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
        dwFlags = dwFlags And Not WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, dwFlags
        SetLayeredWindowAttributes hWnd, 0, 0, LWA_COLORKEY
    End If
    
End Sub

Public Function IsTransparentWindow(ByVal hWnd As Long) As Boolean

    On Error Resume Next
    Dim dwFlags As Long
    dwFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (dwFlags And WS_EX_LAYERED) = WS_EX_LAYERED Then
        IsTransparentWindow = True
    Else
        IsTransparentWindow = False
    End If
    
End Function

'**************************************
' To Change Form Styles at Runtime
' By: Stephen Kent.  This code is copyrighted and has limited warranties. Please see
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=30084&lngWId=1 for details.
'
' Description: This sub-procedure will allow the developer to fairly easily switch between a
' form's border styles during runtime. Normally this isn't really possible because several of
' the attributes are read-only at runtime. This code overcomes those limitations.
' I have only tested this with VB6, but since it is basically just API calls it should be able
' to work with any version that supports API calls.
' (Thanks to Fred_CPP for the tip on using SWP_FRAMECHANGED instead of resizing the form.)
'
' Assumes: for certain buttons to work such as those in the control box they need to be
' enabled in design time (even if they are then hidden at runtime) otherwise there will
' be no handlers linked to those buttons and they will be useless. This applies to
' the What's This Button, Min Button, Max Button, and the Control Box.
' (What's this button has same restrictions on it as it does when used normally)
'**************************************
Public Sub ChangeFormBorder(frmForm As Form, _
    ByVal eNewBorder As FormBorderStyleConstants, _
    Optional ByVal bClipControls As Boolean = True, _
    Optional ByVal bControlBox As Boolean = True, _
    Optional ByVal bMaxButton As Boolean = True, _
    Optional ByVal bMinButton As Boolean = True, _
    Optional ByVal bShowInTaskBar As Boolean = True, _
    Optional ByVal bWhatsThisButton As Boolean = False, _
    Optional ByVal bVisible As Boolean = True)

    Dim lRet As Long
    Dim lStyleFlags As Long
    Dim lStyleExFlags As Long
    
    'Initialize our flags
    lStyleFlags = 0
    lStyleExFlags = 0
    
    'If we want ClipControls then add that flag and change the form property
    If bClipControls Then
        lStyleFlags = lStyleFlags Or WS_CLIPCHILDREN
        frmForm.ClipControls = True
    Else
        frmForm.ClipControls = False
    End If
   
    'If we want the control box then add the flag (property is read-only)
    If bControlBox Then lStyleFlags = lStyleFlags Or WS_SYSMENU
    
    'If we want the max button then add the flag (property is read-only)
    If bMaxButton Then lStyleFlags = lStyleFlags Or WS_MAXIMIZEBOX
    
    'If we want the min button then add the flag (property is read-only)
    If bMinButton Then lStyleFlags = lStyleFlags Or WS_MINIMIZEBOX
    
    'If we want the form to show in taskbar then add the flag (property is read-only)
    If bShowInTaskBar Then lStyleExFlags = lStyleExFlags Or WS_EX_APPWINDOW
    
    'If we want the what's this button then add the flag (property is read-only)
    If bWhatsThisButton Then lStyleExFlags = lStyleExFlags Or WS_EX_CONTEXTHELP
    
    'If the form is an MDI Child form then add the flag (Don't want to screw up the form)
    If frmForm.MDIChild Then lStyleExFlags = lStyleExFlags Or WS_EX_MDICHILD
        
    If bVisible Then lStyleFlags = lStyleFlags Or WS_VISIBLE
        
    'Now we need to set the flags for the border we are changing to
    Select Case eNewBorder
    Case vbBSNone
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS)
        '(no change to extended style flags)
    Case vbFixedSingle
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS Or WS_CAPTION)
        lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
    Case vbSizable
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
        lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
    Case vbFixedDialog
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS Or WS_CAPTION Or DS_MODALFRAME)
        lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME)
    Case vbFixedToolWindow
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS Or WS_CAPTION)
        lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
    Case vbSizableToolWindow
        lStyleFlags = lStyleFlags Or (WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
        lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
    End Select

    'WS_VISIBLE makes sure the form is visible
    'WS_CLIPSIBLINGS makes sure that when there are other windows with the same
    '       relative family that they do not draw over each other.
    'WS_CAPTION provides the form's caption
    'WS_THICKFRAME makes the form sizable
    'DS_MODALFRAME allows dialog forms to have 3d effect
    'WS_EX_WINDOWEDGE is for the border around the form
    'WS_EX_DLGMODALFRAME says the window has a double border and may or may not have a caption
    'WS_EX_TOOLWINDOW says we need a shorter caption and smaller font
    
    'Change our styles
    lRet = SetWindowLong(frmForm.hWnd, GWL_STYLE, lStyleFlags)
    lRet = SetWindowLong(frmForm.hWnd, GWL_EXSTYLE, lStyleExFlags)
    
    'Signal that the frame has changed
    lRet = SetWindowPos(frmForm.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
    
    'Make that we've changed the border in the form's property
    frmForm.BorderStyle = eNewBorder
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectComboByText
'' Description: Attempt to select the given value in the given combo box
'' Inputs:      Combo Box, Value
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectComboByText(cboCombo As ctlUniComboImageXP, ByVal strValue As String) As Boolean

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    If Len(strValue) > 0 Then
        For lIndex = 0 To cboCombo.ListCount - 1
            If cboCombo.List(lIndex) = strValue Then
                cboCombo.ListIndex = lIndex
                bReturn = True
                Exit For
            End If
        Next lIndex
    End If
    
    SelectComboByText = bReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectComboByItemData
'' Description: Attempt to select the given item data in the given combo box
'' Inputs:      Combo Box, Item Data
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SelectComboByItemData(cboCombo As ctlUniComboImageXP, ByVal lValue As Long) As Boolean

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 0 To cboCombo.ListCount - 1
        If cboCombo.ItemData(lIndex) = lValue Then
            cboCombo.ListIndex = lIndex
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    SelectComboByItemData = bReturn

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyComboBox
'' Description: Copy the contents of one combo box into another one
'' Inputs:      Source Combo Box, Destination Combo Box
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CopyComboBox(cboSource As ctlUniComboImageXP, cboDestination As ctlUniComboImageXP)

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To cboSource.ListCount - 1
        cboDestination.AddItem cboSource.List(lIndex)
        cboDestination.ItemData(cboDestination.NewIndex) = cboSource.ItemData(lIndex)
    Next lIndex

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceTheForm
'' Description: Place the given form appropriately
'' Inputs:      Form to Place, INI file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PlaceTheForm(FormToPlace As Form, ByVal strIniFile As String)

    Dim strPlacement As String          ' Form placement
    
    strPlacement = GetIniFileProperty(FormToPlace.Name, "", "Placement", strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm FormToPlace
    Else
        SetFormPlacement FormToPlace, strPlacement
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveTheFormPlacement
'' Description: Save the placement of the given form
'' Inputs:      Form to Save, INI file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveTheFormPlacement(FormToSave As Form, ByVal strIniFile As String)

    SetIniFileProperty FormToSave.Name, GetFormPlacement(FormToSave), "Placement", strIniFile

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetEditorCaption
'' Description: Sets the caption of the form to "Object Type -->  Object Name"
'' Inputs:      Form, Object Type, Object Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditorCaption(frm As Form, ByVal strObjectType As String, ByVal strObjectName As String)

    Dim strCaption$
    
    If Len(Trim(strObjectName)) > 0 Then
        strCaption = Trim(strObjectName) & "   [" & Trim(strObjectType) & "]"
    Else
        strCaption = "New " & Trim(strObjectType)
    End If
    
    frm.Caption = strCaption

End Sub

' To allow use to select a folder
Public Function BrowseForFolder(ByVal strInitialPath$, Optional ByVal strTitle$ = "Select folder ...", Optional ByVal bShowNewFolderButton As Boolean = True) As String

    Dim pidl&, strPath$, iPos&
    Dim bi As BROWSEINFO
    
    ' set owner to active form (so folder dialog won't disappear behind)
    If Not Screen.ActiveForm Is Nothing Then
        bi.hOwner = Screen.ActiveForm.hWnd
    End If
    
    ' store initial path to be used by the callback function
    m_InitialBrowseForFolder = strInitialPath
    
    ' set the BrowseInfo properties
    bi.pidlRoot = 0 'CSIDL_USER_DESKTOP
    bi.lpszTitle = strTitle
    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE 'Or &H10
    If Not bShowNewFolderButton Then
        bi.ulFlags = bi.ulFlags Or BIF_NONEWFOLDERBUTTON
    End If
    bi.lpfnCallback = FunctionPtrToLong(AddressOf BFF_CallBack)
    
    ' call the dialog
    pidl = SHBrowseForFolder(bi)
    If pidl <> 0 Then
        ' if user returned a path, chop it at the null
        strPath = Space(512)
        If SHGetPathFromIDList(ByVal pidl, ByVal strPath) Then
            iPos = InStr(strPath, Chr(0))
            If iPos > 1 Then
                strPath = Left(strPath, iPos - 1)
            Else
                strPath = ""
            End If
        End If
        Call CoTaskMemFree(pidl)
    End If
    
    ' return selected folder
    BrowseForFolder = AddSlash(Trim(strPath))

End Function

'This is the callback function used by BrowseForFolder to set the initial path
Private Function BFF_CallBack(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
On Error Resume Next

    Dim lpIDList As Long
    Dim strBuffer As String
    
    If uMsg = BFFM_INITIALIZED Then
        ' set starting dir
        Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal m_InitialBrowseForFolder)
    ElseIf uMsg = BFFM_SELCHANGED Then
        ' change the status text (current folder) for better viewing
        strBuffer = Space(512)
        If Abs(SHGetPathFromIDList(lp, strBuffer)) Then
            Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, ByVal strBuffer)
        End If
    End If
    BFF_CallBack = 0
    
End Function

' Returns true if the Windows Updater indicates that the machine needs to be rebooted (e.g. to finish installing updates)
Public Function IsRebootRequired() As Boolean

    On Error Resume Next
    Dim objSysInfo As Object
    Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")
    If Not objSysInfo Is Nothing Then
        IsRebootRequired = objSysInfo.RebootRequired
    End If
    
End Function

'JM 12-01-2015 - not needed, using SetWindowTheme API instead
' To enable/disable the normal "theme" for the app's non-client area (i.e. window caption area).
'Public Sub EnableThemeForNonClientArea(frm As Form, ByVal bEnable As Boolean)
'    Dim uFlags&
'
'    If FileExist(WinSysPath & "UxTheme.dll") Then
'        uFlags = 2 Or 4 ' STAP_ALLOW_CONTROLS or STAP_ALLOW_WEBCONTENT
'        If bEnable Then
'            uFlags = uFlags Or 1 ' STAP_ALLOW_NONCLIENT
'        End If
'        SetThemeAppProperties uFlags
'
'        'If hWindowToNotify <> 0 Then
'        '    SendMessage hWindowToNotify, WM_THEMECHANGED, 0, 0
'        'End If
'    End If
'End Sub

' TLB 2/2/2016: the IDE running in Windows 8 is always erroring on any SendKeys command,
' so this override will at least allow us to skip over the errors even if SendKeys doesn't work
Public Sub SendKeys(ByVal strText As String, Optional ByVal bWait As Boolean = False)

    On Error Resume Next
    VBA.SendKeys strText, bWait

End Sub


