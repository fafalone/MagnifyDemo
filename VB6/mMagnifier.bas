Attribute VB_Name = "mMagnifier"
Option Explicit
Private Const MAGFACTOR As Single = 2#
Private Const InvertColors As Boolean = False
Private Const RESTOREDWINDOWSTYLES = WS_SIZEBOX Or WS_SYSMENU Or WS_CLIPCHILDREN Or WS_CAPTION Or WS_MAXIMIZEBOX

Private Const WindowClassName = "MagnifierWindow"
Private Const WindowTitle = "VB6/twinBASIC Magnifier Demo"
Private Const timerInterval = 16
Private hwndMag As LongPtr
Private hwndHost As LongPtr
Private magWindowRect As rect
Private hostWindowRect As rect
Private isFullScreen As Boolean
Sub Main()
    
    If MagInitialize() = CFALSE Then
        Debug.Print "Failed to initialize magnification API."
        Exit Sub
    End If
    
    If SetupMagnifier(App.hInstance) = CFALSE Then
        Debug.Print "Failed to initialize magnifier."
        Exit Sub
    End If
    
    ShowWindow hwndHost, SW_NORMAL
    UpdateWindow hwndHost
    
    Dim timerId As LongPtr: timerId = SetTimer(hwndHost, 0, timerInterval, AddressOf UpdateMagWindow)
    
    Dim tMSG As Msg
    Dim hr As Long
    Debug.Print "Entering message loop"
    hr = GetMessage(tMSG, 0, 0, 0)
    Do While hr <> 0
        If hr = -1 Then
        Debug.Print "Error: 0x" & Hex$(Err.LastDllError)
        Else
            TranslateMessage tMSG
            DispatchMessage tMSG
        End If
        hr = GetMessage(tMSG, 0, 0, 0)
    Loop
    Debug.Print "Exited message loop"
    KillTimer 0, timerId
    MagUninitialize
End Sub

Private Function HostWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_KEYDOWN
            If (wParam = VK_ESCAPE) Then
                If isFullScreen Then
                    GoPartialScreen
                End If
            End If
            
        Case WM_SYSCOMMAND
            If GET_SC_WPARAM(wParam) = SC_MAXIMIZE Then
                GoFullScreen
            Else
                HostWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
            End If
            
        Case WM_DESTROY
            PostQuitMessage (0)
            
        Case WM_SIZE
            GetClientRect hWnd, magWindowRect
            SetWindowPos hwndMag, 0, magWindowRect.Left, magWindowRect.Top, magWindowRect.Right, magWindowRect.Bottom, 0
            
        Case Else
            HostWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End Select
End Function
'Remove after next WDL update
Public Function GET_SC_WPARAM(ByVal wParam As LongPtr) As Long
    GET_SC_WPARAM = (CLng(wParam) And &HFFF0&)
End Function

Private Function RegisterHostWindowClass(hInstance As LongPtr) As Integer
 
    Dim wcex As WNDCLASSEXW

    wcex.cbSize = LenB(wcex)
    wcex.style = CS_HREDRAW Or CS_VREDRAW
    wcex.lpfnWndProc = FARPROC(AddressOf HostWndProc)
    wcex.hInstance = hInstance
    wcex.hCursor = LoadCursorW(0, IDC_ARROW)
    wcex.hbrBackground = (1 + COLOR_BTNFACE)
    wcex.lpszClassName = StrPtr(WindowClassName)

    RegisterHostWindowClass = RegisterClassExW(wcex)
End Function

Private Function SetupMagnifier(hInst As LongPtr) As BOOL
    ' // Set bounds of host window according to screen size.
    hostWindowRect.Top = 0
    hostWindowRect.Bottom = GetSystemMetrics(SM_CYSCREEN) / 4 ' top quarter of screen
    hostWindowRect.Left = 0
    hostWindowRect.Right = GetSystemMetrics(SM_CXSCREEN)
    
    RegisterHostWindowClass hInst
    
    hwndHost = CreateWindowExW(WS_EX_TOPMOST Or WS_EX_LAYERED, _
        StrPtr(WindowClassName), StrPtr(WindowTitle), RESTOREDWINDOWSTYLES, _
        0, 0, hostWindowRect.Right, hostWindowRect.Bottom, 0, 0, hInst, ByVal 0)
    
    If hwndHost = 0 Then
        Debug.Print "Failed to create main window."
        Exit Function
    End If
    
    SetLayeredWindowAttributes hwndHost, 0, 255, LWA_ALPHA
    
    GetClientRect hwndHost, magWindowRect
    hwndMag = CreateWindowW(StrPtr(WC_MAGNIFIER), StrPtr("MagnifierWindow"), _
        WS_CHILD Or MS_SHOWMAGNIFIEDCURSOR Or WS_VISIBLE, _
        magWindowRect.Left, magWindowRect.Top, magWindowRect.Right, magWindowRect.Bottom, hwndHost, 0, hInst, 0)
    
    If hwndMag = 0 Then
        Debug.Print "Failed to create magnifier window."
        Exit Function
    End If
    
    Dim matrix As MAGTRANSFORM
    matrix.v(0, 0) = MAGFACTOR
    matrix.v(1, 1) = MAGFACTOR
    matrix.v(2, 2) = 1
    
    Dim ret As BOOL: ret = MagSetWindowTransform(hwndMag, matrix)
    
    If ret Then
        If InvertColors Then
            Dim magEffectInvert As MAGCOLOREFFECT
            With magEffectInvert 'Reminder: x,y inverted vs cpp
                .transform(0, 0) = -1#: .transform(1, 0) = 0: .transform(2, 0) = 0: .transform(3, 0) = 0: .transform(4, 0) = 0
                .transform(0, 1) = 0: .transform(1, 1) = -1#: .transform(2, 1) = 0: .transform(3, 1) = 0: .transform(4, 1) = 0
                .transform(0, 2) = 0: .transform(1, 2) = 0: .transform(2, 2) = -1#: .transform(3, 2) = 0: .transform(4, 2) = 0
                .transform(0, 3) = 0: .transform(1, 3) = 0: .transform(2, 3) = 0: .transform(3, 3) = 1#: .transform(4, 3) = 0
                .transform(0, 4) = 1#: .transform(1, 4) = 1#: .transform(2, 4) = 1#: .transform(3, 4) = 0: .transform(4, 4) = 1#
            End With
            ret = MagSetColorEffect(hwndMag, magEffectInvert)
        End If
    End If
    
    SetupMagnifier = ret
End Function

Private Sub UpdateMagWindow(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    Dim mousePoint As POINT
    GetCursorPos mousePoint
    
    Dim width As Long: width = (magWindowRect.Right - magWindowRect.Left) / MAGFACTOR
    Dim height As Long: height = (magWindowRect.Bottom - magWindowRect.Top) / MAGFACTOR
    
    Dim sourceRect As rect
    sourceRect.Left = mousePoint.X - width / 2
    sourceRect.Top = mousePoint.Y - height / 2
    
    If (sourceRect.Left < 0) Then
        sourceRect.Left = 0
    End If
    If (sourceRect.Left > GetSystemMetrics(SM_CXSCREEN) - width) Then
        sourceRect.Left = GetSystemMetrics(SM_CXSCREEN) - width
    End If
    sourceRect.Right = sourceRect.Left + width

    If (sourceRect.Top < 0) Then
        sourceRect.Top = 0
    End If
    If (sourceRect.Top > GetSystemMetrics(SM_CYSCREEN) - height) Then
        sourceRect.Top = GetSystemMetrics(SM_CYSCREEN) - height
    End If
    sourceRect.Bottom = sourceRect.Top + height

    'Set the source rectangle for the magnifier control.
    #If Win64 Then
    MagSetWindowSource hwndMag, sourceRect
    #Else
    MagSetWindowSource hwndMag, sourceRect.Left, sourceRect.Top, sourceRect.Right, sourceRect.Bottom
    #End If

    'Reclaim topmost status, to prevent unmagnified menus from remaining in view.
    SetWindowPos hwndHost, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE

    'Force redraw.
    InvalidateRect hwndMag, ByVal 0, CTRUE
End Sub

Private Sub GoFullScreen()
 
        isFullScreen = True
        ' // The window must be styled As layered For proper rendering.
        ' // It Is styled As transparent so that it does Not capture mouse clicks.
        SetWindowLong hwndHost, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
        ' // Give the window a system menu so it can be closed On the taskbar.
        SetWindowLong hwndHost, GWL_STYLE, WS_CAPTION Or WS_SYSMENU

        ' // Calculate the span of the display area.
        Dim hDC As LongPtr: hDC = GetDC(0)
        Dim xSpan As Long: xSpan = GetSystemMetrics(SM_CXSCREEN)
        Dim ySpan As Long: ySpan = GetSystemMetrics(SM_CYSCREEN)
        ReleaseDC 0, hDC

        ' // Calculate the size of system elements.
        Dim xBorder As Long: xBorder = GetSystemMetrics(SM_CXFRAME)
        Dim yCaption As Long: yCaption = GetSystemMetrics(SM_CYCAPTION)
        Dim yBorder As Long: yBorder = GetSystemMetrics(SM_CYFRAME)

        ' // Calculate the window origin And span For full-screen mode.
        Dim xOrigin As Long: xOrigin = -xBorder
        Dim yOrigin As Long: yOrigin = -yBorder - yCaption
        xSpan = xSpan + 2 * xBorder
        ySpan = ySpan + 2 * yBorder + yCaption

        SetWindowPos hwndHost, HWND_TOPMOST, xOrigin, yOrigin, xSpan, ySpan, _
            SWP_SHOWWINDOW Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub
    ' //
    ' // FUNCTION: GoPartialScreen()
    ' //
    ' // PURPOSE: Makes the host window resizable And focusable.
    ' //
Private Sub GoPartialScreen()
 
        isFullScreen = False

        SetWindowLong hwndHost, GWL_EXSTYLE, WS_EX_TOPMOST Or WS_EX_LAYERED
        SetWindowLong hwndHost, GWL_STYLE, RESTOREDWINDOWSTYLES
        SetWindowPos hwndHost, HWND_TOPMOST, _
            hostWindowRect.Left, hostWindowRect.Top, hostWindowRect.Right, hostWindowRect.Bottom, _
            SWP_SHOWWINDOW Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

