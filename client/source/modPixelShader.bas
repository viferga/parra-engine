Attribute VB_Name = "modPixelShader"
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

Public ShowShader As Boolean

Public Const dats           As String = "ps.1.1 tex t0 tex t1 add_sat r0, t0, t1 mul r0, t0, v0" 'add_sat r0, t0, t1
Public Const Glowsrc        As String = "vs.1.1 dcl_position v0 dcl_texcoord v3 mov oPos, v0 add oT0, v3, c0 add oT1, v3, c1 add oT2, v3, c2 add oT3, v3, c3"
Public Const Gussian_Blur   As String = "ps.1.4 def c0, 0.2f, 0.2f, 0.2f, 1.0f texld r0, t0 texld r1, t1 texld r2, t2 texld r3, t3 texld r4, t4 add r0, r0, r1 add r2, r2, r3 add r0, r0, r2 add r0, r0, r4 mul r0, r0, c0"
Public Const BrightPasssrc  As String = "ps.1.4 def c0,0.561797752,0.561797752,0.561797752,1 def c1,0.78125,0.78125,0.78125,1 def c2,1,1,1,1 def c3,0.1,0.1,0.1,1 texld r0,t0 mul_x4 r0,r0,c0 mul_x2 r1,r0,c1 add r1,r1,c2 mul r0,r0,r1 mov_x4 r2,c2 add r2,r2,c2 sub r0,r0,r2 mul_sat r0,r0,c3"
'Public Const Glowsrc        As String = "ps.1.1 " & _
                                        "def c0, 0.4, 0.4, 0.4, 0.4 " & _
                                        "tex t0 " & _
                                        "mul r0, t0, c0"

Public Const psDesaturation As String = "ps.1.1 def c0, 0.3, 0.59, 0.11, 1 def c1, 0.5, 0.5, 0.5, 0 tex t0 dp3 r0, t0, c0 add r0,c1,r0"
Public Const psOriginalColor As String = "ps.1.1 tex t0 tex t3 mul r0, t0, t3"


'Pixel shader in use
Public cPixelShader As Long

Dim DXlngShaderArray() As Long
Dim DXlngShaderSize As Long
Dim DXBufferCode As D3DXBuffer
Public Function pixelShaderMake(PSFileName As String)
'Assemble a pixel shader from a file, returning its handle
On Error GoTo PSErr
Set DXBufferCode = D3DX.AssembleShaderFromFile(PSFileName$, 0, "", Nothing)
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
D3DX.BufferGetData DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0)
pixelShaderMake = D3DDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, ""
Set DXBufferCode = Nothing
End Function
Public Function pixelShaderMakeFromMemory(PSContents As String)
'Assemble a pixel shader from a string, return its handle
On Error GoTo PSErr
Set DXBufferCode = D3DX.AssembleShader(PSContents$, 0, Nothing, "")
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
D3DX.BufferGetData DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0)
pixelShaderMakeFromMemory = D3DDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, ""
Set DXBufferCode = Nothing
End Function
Public Sub pixelShaderSet(ByRef lngPixelShaderHandle As Long)
'Enables a pixel shader
D3DDevice.SetPixelShader lngPixelShaderHandle&
End Sub
Public Sub pixelShaderDelete(ByRef lngPixelShaderHandle As Long)
'Deletes a pixel shader
D3DDevice.DeletePixelShader lngPixelShaderHandle&
End Sub
Public Sub pixelShaderSetCRegister(RegIndex As Long, ValR As Single, ValG As Single, ValB As Single, ValA As Single)
Dim SinArr(3)
SinArr(0) = ValR
SinArr(1) = ValG
SinArr(2) = ValB
SinArr(3) = ValA
D3DDevice.SetPixelShaderConstant RegIndex, SinArr(0), 4
End Sub
