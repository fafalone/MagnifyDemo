Attribute VB_Name = "mDefs"
Option Explicit
#If VBA7 = 0 Then
Public Enum LongPtr
    [_]
End Enum
#End If

Public Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINT
    X As Long
    Y As Long
End Type

Public Enum BOOL
    CFALSE
    CTRUE
End Enum

Public Enum WindowStyles
    WS_OVERLAPPED = &H0
    WS_POPUP = &H80000000
    WS_CHILD = &H40000000
    WS_MINIMIZE = &H20000000
    WS_VISIBLE = &H10000000
    WS_DISABLED = &H8000000
    WS_CLIPSIBLINGS = &H4000000
    WS_CLIPCHILDREN = &H2000000
    WS_MAXIMIZE = &H1000000
    WS_BORDER = &H800000
    WS_DLGFRAME = &H400000
    WS_VSCROLL = &H200000
    WS_HSCROLL = &H100000
    WS_SYSMENU = &H80000
    WS_THICKFRAME = &H40000
    WS_GROUP = &H20000
    WS_TABSTOP = &H10000
    WS_MINIMIZEBOX = &H20000
    WS_MAXIMIZEBOX = &H10000
    WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
    WS_TILED = WS_OVERLAPPED
    WS_ICONIC = WS_MINIMIZE
    WS_SIZEBOX = WS_THICKFRAME
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
    WS_CHILDWINDOW = WS_CHILD
End Enum
Public Enum WindowStylesEx
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_TOPMOST = &H8
    WS_EX_ACCEPTFILES = &H10
    WS_EX_TRANSPARENT = &H20
    WS_EX_MDICHILD = &H40
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    WS_EX_LAYERED = &H80000
    WS_EX_NOINHERITLAYOUT = &H100000   ' Disable inheritence of mirroring by children
    WS_EX_NOREDIRECTIONBITMAP = &H200000
    WS_EX_LAYOUTRTL = &H400000   ' Right to left mirroring
    WS_EX_COMPOSITED = &H2000000
    WS_EX_NOACTIVATE = &H8000000
End Enum


Public Enum ShowWindow
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
    SW_MAX = 11
End Enum

Public Type Msg
    hWnd        As LongPtr
    message     As Long
    wParam      As LongPtr
    lParam      As LongPtr
    Time        As Long
    PT          As POINT
End Type


Public Const WM_KEYDOWN = &H100
Public Const VK_ESCAPE = &H1B
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MAXIMIZE = &HF030&
Public Const WM_DESTROY = &H2
Public Const WM_SIZE = &H5
Public Const SM_CXSCREEN = 0 ' 0x00
Public Const SM_CYSCREEN = 1 ' 0x01
Public Const SM_CXFRAME = 32 ' 0x20
Public Const SM_CYFRAME = 33 ' 0x21
Public Const SM_CYCAPTION = 4 ' 0x04
Public Const IDC_ARROW = 32512&
Public Const COLOR_BTNFACE = 15

Public Enum ClassStyles
     CS_VREDRAW = &H1
     CS_HREDRAW = &H2
     CS_DBLCLKS = &H8
     CS_OWNDC = &H20
     CS_CLASSDC = &H40
     CS_PARENTDC = &H80
     CS_NOCLOSE = &H200
     CS_SAVEBITS = &H800
     CS_BYTEALIGNCLIENT = &H1000
     CS_BYTEALIGNWINDOW = &H2000
     CS_GLOBALCLASS = &H4000
     CS_IME = &H10000
     CS_DROPSHADOW = &H20000
End Enum

Public Type WNDCLASSEXW
    cbSize As Long
    style As ClassStyles
    lpfnWndProc As LongPtr
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As LongPtr
    hIcon As LongPtr
    hCursor As LongPtr
    hbrBackground As LongPtr
    lpszMenuName As LongPtr
    lpszClassName As LongPtr
    hIconSm As LongPtr
End Type

Public Enum LayeredWindowAttributes
    LWA_COLORKEY = &H1
    LWA_ALPHA = &H2
End Enum

Public Enum SWP_Flags
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOOWNERZORDER = &H200
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSENDCHANGING = &H400
    
    SWP_DEFERERASE = &H2000
    SWP_ASYNCWINDOWPOS = &H4000
End Enum
Public Enum WindowZOrderDefaults
    HWND_DESKTOP = 0&
    HWND_TOP = 0&
    HWND_BOTTOM = 1&
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
End Enum

Public Enum GWL_INDEX
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
End Enum









Public Const WC_MAGNIFIERW = "Magnifier"
Public Const WC_MAGNIFIER = WC_MAGNIFIERW

Public Enum MagnifierWindowStyles
    MS_SHOWMAGNIFIEDCURSOR = &H1
    MS_CLIPAROUNDCURSOR = &H2
    MS_INVERTCOLORS = &H4
End Enum

Public Type MAGTRANSFORM
    v(0 To 2, 0 To 2) As Single
End Type


Public Type MAGCOLOREFFECT
    transform(0 To 4, 0 To 4) As Single
End Type

#If VBA7 Then
Public Declare PtrSafe Function MagInitialize Lib "magnification.dll" () As BOOL
Public Declare PtrSafe Function MagUninitialize Lib "magnification.dll" () As BOOL
Public Declare PtrSafe Function MagSetWindowTransform Lib "magnification.dll" (ByVal hwnd As LongPtr, pTransform As MAGTRANSFORM) As BOOL
Public Declare PtrSafe Function MagSetColorEffect Lib "magnification.dll" (ByVal hwnd As LongPtr, pEffect As MAGCOLOREFFECT) As BOOL
#If Win64 Then
Public Declare PtrSafe Function MagSetWindowSource Lib "magnification.dll" (ByVal hwnd As LongPtr, rect As RECT) As BOOL
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function MagSetWindowSource Lib "magnification.dll" (ByVal hwnd As LongPtr, ByVal rectLeft As Long, ByVal rectRight As Long, ByVal rectTop As Long, ByVal rectBottom As Long) As BOOL
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If
Public Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Any, ByVal bErase As BOOL) As BOOL
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As SWP_Flags) As Long
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As SHOWWINDOW) As Long
Public Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As BOOL
Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As BOOL
Public Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageW" (lpMsg As MSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare PtrSafe Function TranslateMessage Lib "user32" (ByRef lpmsg As MSG) As BOOL
Public Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByRef lpmsg As MSG) As LongPtr
Public Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As BOOL
Public Declare PtrSafe Function RegisterClassExW Lib "user32" (pcWndClassEx As WNDCLASSEXW) As Integer
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As LayeredWindowAttributes) As BOOL
Public Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal dwExStyle As WindowStylesEx, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As WindowStyles, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As BOOL
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function LoadCursorW Lib "user32" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
#Else
Public Declare Function MagInitialize Lib "magnification.dll" () As BOOL
Public Declare Function MagUninitialize Lib "magnification.dll" () As BOOL
Public Declare Function MagSetWindowTransform Lib "magnification.dll" (ByVal hWnd As LongPtr, pTransform As MAGTRANSFORM) As BOOL
Public Declare Function MagSetColorEffect Lib "magnification.dll" (ByVal hWnd As LongPtr, pEffect As MAGCOLOREFFECT) As BOOL
#If Win64 Then
Public Declare Function MagSetWindowSource Lib "magnification.dll" (ByVal hWnd As LongPtr, rect As rect) As BOOL
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare Function MagSetWindowSource Lib "magnification.dll" (ByVal hWnd As LongPtr, ByVal rectLeft As Long, ByVal rectRight As Long, ByVal rectTop As Long, ByVal rectBottom As Long) As BOOL
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Any, ByVal bErase As BOOL) As BOOL
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As SWP_Flags) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As ShowWindow) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As BOOL
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As BOOL
Public Declare Function GetMessage Lib "user32" Alias "GetMessageW" (lpmsg As Msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (ByRef lpmsg As Msg) As BOOL
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByRef lpmsg As Msg) As LongPtr
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As rect) As BOOL
Public Declare Function RegisterClassExW Lib "user32" (pcWndClassEx As WNDCLASSEXW) As Integer
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As LayeredWindowAttributes) As BOOL
Public Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As WindowStylesEx, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As WindowStyles, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As BOOL
Public Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare Function LoadCursorW Lib "user32" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr

#End If







Public Function CreateWindowW(ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As WindowStyles, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    CreateWindowW = CreateWindowExW(0, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, ByVal lpParam)
End Function

Public Function FARPROC(lpfn As LongPtr) As LongPtr
FARPROC = lpfn
End Function
