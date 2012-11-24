Attribute VB_Name = "modSystemTray"
Option Explicit

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4

Private Const STI_CALLBACKEVENT = &H201

Public Const STI_LBUTTONDOWN = &H201
Public Const STI_LBUTTONUP = &H202
Public Const STI_LBUTTONDBCLK = &H203
Public Const STI_RBUTTONDOWN = &H204
Public Const STI_RBUTTONUP = &H205
Public Const STI_RBUTTONDBCLK = &H206

Public Sub CreateSystemTrayIcon(ByRef parentForm As Form, ByVal Tip As String)
  
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_ADD, notIcon
End Sub

Public Sub ModifySystemTrayIcon(ByRef parentForm As Form, ByVal Tip As String)
  
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = STI_CALLBACKEVENT
    .hIcon = parentForm.Icon
    .szTip = Tip & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_MODIFY, notIcon
End Sub

Public Sub DeleteSystemTrayIcon(ByRef parentForm As Form)
  
  Dim notIcon As NOTIFYICONDATA
  
  With notIcon
    .cbSize = Len(notIcon)
    .hwnd = parentForm.hwnd
    .uID = vbNull
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .uCallbackMessage = vbNull
    .hIcon = vbNull
    .szTip = "" & vbNullChar
  End With
  
  Shell_NotifyIconA NIM_DELETE, notIcon
End Sub

