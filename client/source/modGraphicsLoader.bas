Attribute VB_Name = "modGraphicsLoader"
#If LoadingMetod = 1 Then

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


#Else

Option Explicit

'Manage Textures
Private Type structGraphic
        FileName   As Long
        D3DTexture As DxVBLibA.Direct3DTexture8
        Used       As Long
        Available  As Boolean
        Width      As Integer
        Height     As Integer
End Type: Public oGraphic() As structGraphic: Public lKeys() As Long

Private lSurfaceSize As Long
Private TexInfo      As D3DXIMAGE_INFO
Private Const lMaxEntrys As Integer = 1000

Private Type tCache
    Number        As Long
    SrcHeight     As Single
    SrcWidth      As Single
End Type: Private Cache As tCache

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal numBytes As Long)

#End If


