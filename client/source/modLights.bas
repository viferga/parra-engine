Attribute VB_Name = "modLights"
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

Public Function CalcVertexLight(ByVal Radio As Byte, LightX As Single, LightY As Single, VertexX As Single, VertexY As Single, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim lXDistance As Single
    Dim lYDistance As Single
    Dim lVertexDistance As Single
    Dim pRadio As Long
    
    pRadio = Radio * 64
        
    ' Calculate distance from vertex
    lXDistance = CLng((Abs(LightX + 64 - VertexX)) / 2)
    lYDistance = CLng(Abs(LightY + 64 - VertexY))
    
    lVertexDistance = CLng(Sqr(lXDistance * lXDistance + lYDistance * lYDistance))
    
    If lVertexDistance <= pRadio Then
    
        Dim CurrentColor As D3DCOLORVALUE
        
        D3DXColorLerp CurrentColor, LightColor, AmbientColor, lVertexDistance / pRadio
        CalcVertexLight = D3DColorXRGB(CurrentColor.r, CurrentColor.g, CurrentColor.b)
    Else
        ' Return lowest value
        CalcVertexLight = D3DColorXRGB(AmbientColor.r, AmbientColor.g, AmbientColor.b)
    End If
End Function
