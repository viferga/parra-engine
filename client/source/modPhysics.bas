Attribute VB_Name = "modPhysics"
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

Public Type structPositionLng
    X   As Long
    Y   As Long
End Type

Public Type structPositionSng
    X   As Single
    Y   As Single
End Type

Public Type structPositionInt
    X   As Integer
    Y   As Integer
End Type

Public Type structPosByte
    X   As Byte
    Y   As Byte
End Type
Public Function isMouseOverQuad(MouseX As Single, MouseY As Single, X1 As Single, X2 As Single, Y1 As Single, Y2 As Single) As Boolean
If MouseX >= X1 And MouseX <= X2 And MouseY >= Y1 And MouseY <= Y2 Then
    isMouseOverQuad = True
Else
    isMouseOverQuad = False
End If
End Function
