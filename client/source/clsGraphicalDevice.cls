VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Type videoMode
     Width As Long
     Height As Long
     BitsPerPixel As Long
End Type

Public Enum DEVICE_NO_ERROR
    DEVICE_CTX_VALID
    DEVICE_CTX_RESET
    DEVICE_CTX_LOST
    DEVICE_CTX_ERROR
End Enum

Public Enum CLEAR_FLAGS
    CLEAR_COLOR = 1
    CLEAR_DEPTH = 2
    CLEAR_STENCIL = 4
End Enum

Private VideoModeList() As videoMode
Private VideoModeListSize As Long

Public Function Initialize(ByRef Top As Integer, ByRef Left As Integer, ByRef Width As Integer, ByRef Height As Integer, Frm As Form, Optional ByRef BitsPerPixel As Byte = 32, Optional ByRef Windowed As Boolean = True) As Boolean

End Function

Private Function VideoModeGetDesktopMode() As videoMode

End Function

Private Function VideoModeGetAvailableModes() As videoMode()

End Function

Public Sub DeInitialize()

End Sub

Public Function DeviceIsContextValid() As DEVICE_NO_ERROR

End Function

Public Sub BeginScene(ClearRect As RECT, flags As CLEAR_FLAGS, Optional Blur As Boolean = False)

End Sub

Public Sub EndScene(DestRect As RECT, hWnd As Long)

End Sub

Public Function initializeMotionBlur() As Boolean

End Function

Public Sub fontRender(ByRef Text As String, ByRef index As Byte, _
                            ByRef X As Integer, ByRef Y As Integer, _
                            ByRef Width As Integer, ByRef Height As Integer, _
                            format As Long) 'GDK: revisar parametros, no se si seran igual en ogl.
                            

End Sub
Public Function fontInitializing(ByRef Size As Byte) As Boolean

End Function

Public Sub guiRender(ByVal Instance As Byte)

End Sub

Public Sub resetRenderStates()

End Sub

Public Sub resetMotionStates()

End Sub

Public Sub renderBox(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                            Optional ByVal v1 As Long = -1, Optional ByVal v2 As Long = -1, _
                            Optional ByVal v3 As Long = -1, Optional ByVal v4 As Long = -1)

End Sub

Public Sub renderParticleGroup(grIndex As Integer)

End Sub

Public Sub renderTexture(ByRef GrhIndex As Long, ByRef cx As Single, ByRef cy As Single, ByRef Color() As Long, ByRef Iso As IsometricType, Optional ByRef Angle As Single = 0)

End Sub

Public Function FloatToDWord(flo As Single) As Long

End Function
