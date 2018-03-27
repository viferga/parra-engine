Attribute VB_Name = "modHelper"
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

'Processing Data
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Public Sub WriteVar(ByRef file As String, ByRef Main As String, ByRef Var As String, ByRef Value As String): writeprivateprofilestring Main, Var, Value, file: End Sub
Public Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Public Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Integer) As String

Dim I As Integer
Dim UPos As Integer
Dim Separador As String * 1
Dim max As Integer
Dim Lent As Integer
Dim tControl As Integer

Lent = Len(Text)

If Lent <= 1 Then Exit Function

Separador = Chr$(SepASCII)

UPos = 0

For I = 1 To Pos - 1
    UPos = InStr(UPos + 1, Text, Separador, vbTextCompare)
Next

tControl = InStr(UPos + 1, Text, Separador, vbTextCompare) - 1

If tControl = 0 And Pos = 1 Then Exit Function

max = Abs(UPos - tControl)
ReadField = Mid$(Text, UPos + 1, IIf(tControl <= 0, Abs(UPos - Lent), max))

End Function
Public Function FileExist(ByRef file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    If Dir$(file, FileType) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Public Sub IntializeRandom()
    Randomize Timer
End Sub
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
