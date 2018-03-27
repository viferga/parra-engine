Attribute VB_Name = "modGeneral"
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

Public bRunning As Boolean 'Switch on/off the game

'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Sub Main()

    'Show the form
    frmMain.Show
        
    'bRunning is true if TileEninge initialize correctly
    bRunning = engineInitializing(130, 230, 950, 650, frmMain, 16, True)
    
    'Work Over The Form
    Call SetWindowPos(frmMain.hwnd, 0, 0, 0, 1000, 700, 0)

    'Starts the game
    gameLoop
    
    'Deinit engine
    engineDeinitializing
    
    Unload frmMain
    
    End
End Sub

Public Sub gameLoop()

    Dim lTimerCount As Long

    Do While (bRunning = True)
    
        gameCheckKeys
       
        'Run Engine
        showNextFrame
    
        If (GetTickCount - lTimerCount >= 1000) Then
            
            FramesPerSec = FramesPerSecCounter
            FramesPerSecCounter = 0
            lTimerCount = GetTickCount
        End If

        DoEvents
    Loop
    
End Sub

Public Sub gameCheckKeys()
    
       If (GetActiveWindow() <> frmMain.hwnd) Then Exit Sub

            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then

                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
            
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
                
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
                 
                 Exit Sub
            End If
            
            'Atack!
            If GetKeyState(vbKeyControl) < 0 Then
                
                Exit Sub
            End If
            
            'Trow An Item
            If GetKeyState(vbKeyT) < 0 Then
                
                Exit Sub
            End If
    
End Sub

