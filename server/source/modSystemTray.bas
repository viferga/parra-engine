Attribute VB_Name = "modSystemTray"
Option Explicit

'****************************************************************************
'    Parra Engine is a MMORPG Isometric Game Engine.
'    Copyright (C) 2009 - 2013 Vicente Eduardo Ferrer Garcia
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as
'    published by the Free Software Foundation, either version 3 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'****************************************************************************

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

