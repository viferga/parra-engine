Attribute VB_Name = "modCreateWindows"
Option Explicit

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2

Public Const WM_DESTROY = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_SETCURSOR = &H20

Public Const IDC_ARROW = 32512&
Public Const ES_MULTILINE = &H4&
Public Const GWL_WNDPROC = (-4)
Public Const CW_USEDEFAULT = &H80000000
Public Const COLOR_WINDOW = 5
Public Const IDI_APPLICATION = 32512&

Public Const SW_SHOWNORMAL = 1

Public Const MB_OK = &H0&
Public Const MB_ICONEXCLAMATION = &H30&

Public Const ClassNameStr = "AnyClassName"
Public Const AppNameStr = "Parra-Engine v2"

'Public ButOldProcVal As Long
Public HwndVal As Long
'Public ButtonHwndVal As Long
'Public EditHwndVal As Long
'Public ScrollHwndVal As Long
'Public ComboHwndVal As Long

Public Function CreateWindows() As Boolean

'You can use the following classnames
'-BUTTON (This is a normal pushbutton)
'-COMBOBOX (This is a combobox)
'-EDIT (This is a textbox)
'-LISTBOX (This is a listbox)
'-MDICLIENT (This is a MDI Client window)
'-RICHEDIT (This is an Richtextbox V1.0 control)
'-RICHEDIT_CLASS (This is an Richtextbox V2.0 control)
'-SCROLLBAR (This is a HScrollbar)
'-STATIC (Like line controls)

'Initialize the windows
'Cambiar el 800x600 por la resolucion que usemos con el init de Dx
HwndVal& = CreateWindowEx(0&, ClassNameStr$, AppNameStr$, WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 800, 600, 0&, 0&, App.hInstance, ByVal 0&)

'EditHwndVal& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "A TextBox", WS_CHILD Or ES_MULTILINE, 50, 20, 200, 80, HwndVal&, 0&, App.hInstance, 0&)

'ButtonHwndVal& = CreateWindowEx(0&, "Button", "A Button", WS_CHILD, 50, 120, 200, 25, HwndVal&, 0&, App.hInstance, 0&)

'ScrollHwndVal& = CreateWindowEx(0&, "SCROLLBAR", "", WS_CHILD, 50, 150, 200, 25, HwndVal&, 0&, App.hInstance, 0&)

'ComboHwndVal& = CreateWindowEx(0&, "COMBOBOX", "A Combobox", WS_CHILD, 50, 190, 200, 20, HwndVal&, 0&, App.hInstance, 0&)

'Create the windows
Call ShowWindow(HwndVal&, SW_SHOWNORMAL)
'Call ShowWindow(ButtonHwndVal&, SW_SHOWNORMAL)
'Call ShowWindow(EditHwndVal&, SW_SHOWNORMAL)
'Call ShowWindow(ScrollHwndVal&, SW_SHOWNORMAL)
'Call ShowWindow(ComboHwndVal&, SW_SHOWNORMAL)

'ButOldProcVal& = GetWindowLong(ButtonHwndVal&, GWL_WNDPROC)

'Call SetWindowLong(ButtonHwndVal&, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc))

CreateWindows = (HwndVal& <> 0)

End Function

Public Sub CreateForms()

Dim wMsg As Msg

If RegisterWindowClass = False Then Exit Sub

If CreateWindows Then

Do While GetMessage(wMsg, 0&, 0&, 0&)

    Call TranslateMessage(wMsg)
    
    Call DispatchMessage(wMsg)
Loop

End If

Call UnregisterClass(ClassNameStr$, App.hInstance)

End Sub

Public Function RegisterWindowClass() As Boolean

Dim wc As WNDCLASS

wc.style = CS_HREDRAW Or CS_VREDRAW
wc.lpfnwndproc = GetAddress(AddressOf WndProc)
wc.hInstance = App.hInstance
wc.hIcon = LoadIcon(0&, IDI_APPLICATION)
wc.hCursor = LoadCursor(0&, IDC_ARROW)
wc.hbrBackground = COLOR_WINDOW
wc.lpszClassName = ClassNameStr$

RegisterWindowClass = RegisterClass(wc) <> 0

End Function

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim strTemp As String

    If (hWnd = Null) Then Exit Function 'Existe la ventana?


    Select Case uMsg&

        Case WM_DESTROY:
            Call PostQuitMessage(0&)
        
        Case WM_SETCURSOR
            'SetCursor hCursor
        
        
    End Select

WndProc = DefWindowProc(hWnd&, uMsg&, wParam&, lParam&)

End Function

'Public Function ButtonWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Select Case uMsg&
'Case WM_LBUTTONUP:
''This is executed if the button is clicked
'MsgBox "Estoy apretando un boton en un .frm inexistenteeee XD"
'End Select

'ButtonWndProc = CallWindowProc(ButOldProcVal&, hwnd&, uMsg&, wParam&, lParam&)

'End Function
Public Function GetAddress(ByVal lngAddr As Long) As Long

GetAddress = lngAddr&

End Function

