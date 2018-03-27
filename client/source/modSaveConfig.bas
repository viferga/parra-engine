Attribute VB_Name = "modSaveConfig"
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

' *********************************************************************
' Autor: Emanuel Matias (Dunkan)
' Nota: Almacena 8 Bool (16 Bytes) en 1 Byte.
'       Se podria ademas hacer esto con un long y almacenar 32 variables _
        (Innecesario por ahora)
'       ���Esto es solo una prueba!!!
' *********************************************************************

' ****************************************
'   Private Type StructConfig
'       OGL As Boolean '�Use OGL?
'   End Type
' ****************************************

Private Const lSaved As Byte = &HFF 'Save(True, True, True, True, True, True, True, True)

Public Sub testConfig()

    Dim Ind     As Long
    Dim tmpByte As Byte
    
    'Const lSaveConf = &H4B
    '4B = Hex(SaveByteConf(True, True, False, True, False, False, True, False))
    '&H4B = tmpByte
    
    tmpByte = SaveByteConf(True, True, False, True, False, False, True, False)
    
    For Ind = 0 To 7
        If (ReadByteConf(tmpByte, Ind) = True) Then
            MsgBox "True, option n�: " & (Ind + 1)
        End If
    Next Ind

End Sub

Private Function ReadByteConf(ByVal lSrc As Byte, ByVal pInd As Byte) As Boolean
    ReadByteConf = (lSrc And 2 ^ pInd)
End Function

Private Function SaveByteConf(ParamArray bParam() As Variant) As Byte
    Dim pInd    As Long
    Dim tmpB()  As Variant
    
    tmpB() = bParam()
    ReDim Preserve tmpB(0 To 7) '8 bits
    
    SaveByteConf = 0
    For pInd = 0 To 7
        If tmpB(pInd) Then SaveByteConf = (SaveByteConf Or 2 ^ pInd)
    Next pInd
End Function

