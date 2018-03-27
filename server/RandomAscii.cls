VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Version 0.1.0                                                                                                          ''''
''''                                                                                                                        ''''
'''' Esta clase, genera un n�mero aleatorio desde una frase...                                                              ''''
''''                                                                                                                        ''''
'''' Agradezco a ElYind por su idea original, codeada por mi.                                                               ''''
''''                                                                                                                        ''''
'''' Lo unico malo de esta clase, es que la longitud y el tama�o de la cadena de                                            ''''
'''' texto se triplica al original, pero si tenemos un buen host, vale la pena                                              ''''
'''' la implementaci�n.                                                                                                     ''''
''''                                                                                                                        ''''
'''' Aunque, como ver�n en la iniciaci�n de la clase se setea un Array (sugerencia de Parra),                               ''''
'''' yo quer�a hacerlo con un Select Case o If/ElseIf :P. Que no est� terminado, y en las funciones                         ''''
'''' de randomize y read est�n comentadas... pero ya est� casi todo echo, el que se pone la termina                         ''''
'''' en menos de 2 horas ;P.                                                                                                ''''
'''' La funci�n es b�sica, tomamos de a dos caracteres de los n�meros generados (XX) y lo reemplaza por                     ''''
'''' un caracter, hay que tener en cuenta que podemos tener 01 02 10 90 52 etc, todas las combinaciones posibles            ''''
'''' y a la hora de leer la frase, es cuestion de usar un Replace() o mismo con un for, como gusten.                        ''''
''''                                                                                                                        ''''
'''' Esto fu� generado, con la idea de armar una dll para distribuir en todos lo servidores y usar                          ''''
'''' alg�n .ini o .dat externo para que no sea en todos los servers iguales :P.                                             ''''
''''                                                                                                                        ''''
'''' Justo con los chicos sali� la idea de un server y bueno... no se logr�, cualquier cosa,                                ''''
'''' les dejo mi mail para que me contacten si tienen ganas. thomy.-@hotmail.com                                            ''''
''''                                                                                                                        ''''
''''                                                                                              Buenos Aires, Argentina   ''''
''''                                                                                                Septiembre del 2009     ''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

Private CaracterDividido(0 To 255) As String


Private Sub Class_Initialize()
'*********************************
'** Autor: DarkThom **************
'** Fecha: 5/09/2009 *************
'*********************************

Dim a As Integer, b As Integer

CaracterDividido(10) = "@"
CaracterDividido(11) = "�"
CaracterDividido(12) = "�"
CaracterDividido(13) = "k"
CaracterDividido(14) = "�"
CaracterDividido(15) = "�"
CaracterDividido(16) = "4"
CaracterDividido(17) = "."
CaracterDividido(18) = ";"
CaracterDividido(19) = ":"
CaracterDividido(20) = ","
CaracterDividido(21) = "-"
CaracterDividido(22) = "_"
CaracterDividido(23) = "'"
CaracterDividido(24) = "�"
CaracterDividido(25) = "A"
CaracterDividido(26) = "�"
CaracterDividido(27) = "$"
CaracterDividido(28) = "�"
CaracterDividido(29) = "�"
CaracterDividido(30) = "x"
CaracterDividido(31) = "+"
CaracterDividido(32) = "^"
CaracterDividido(33) = "*"
CaracterDividido(34) = "�"
CaracterDividido(35) = "�"
CaracterDividido(36) = "�"
CaracterDividido(37) = ""
CaracterDividido(38) = "�"
CaracterDividido(39) = "�"
CaracterDividido(40) = "�"
CaracterDividido(41) = "{"
CaracterDividido(42) = "}"
CaracterDividido(43) = "["
CaracterDividido(44) = "]"
CaracterDividido(45) = "#"
CaracterDividido(46) = "~"
CaracterDividido(47) = "�"
CaracterDividido(48) = "\"
CaracterDividido(49) = "/"
CaracterDividido(50) = "�"
CaracterDividido(51) = "?"
CaracterDividido(52) = "="
CaracterDividido(53) = "�"
CaracterDividido(54) = "�"
CaracterDividido(55) = "!"
CaracterDividido(56) = "�"
CaracterDividido(57) = "�"
CaracterDividido(58) = "�"
CaracterDividido(59) = "�"
CaracterDividido(60) = "�"
CaracterDividido(61) = "�"
CaracterDividido(62) = "|"
CaracterDividido(63) = "<"
CaracterDividido(64) = ">"
CaracterDividido(65) = "A"
CaracterDividido(66) = "B"
CaracterDividido(67) = "C"
CaracterDividido(68) = "D"
CaracterDividido(69) = "E"
CaracterDividido(70) = "F"
CaracterDividido(71) = "G"
CaracterDividido(72) = "H"
CaracterDividido(73) = "I"
CaracterDividido(74) = "J"
CaracterDividido(75) = "K"
CaracterDividido(76) = "L"
CaracterDividido(77) = "E"
CaracterDividido(78) = "M"
CaracterDividido(79) = "N"
CaracterDividido(80) = "O"
CaracterDividido(81) = "P"
CaracterDividido(82) = "Q"
CaracterDividido(83) = "R"
CaracterDividido(84) = "S"
CaracterDividido(85) = "T"
CaracterDividido(86) = "W"
CaracterDividido(87) = "X"
CaracterDividido(88) = "Y"
CaracterDividido(89) = "Z"
CaracterDividido(90) = "V"
CaracterDividido(91) = "0"
CaracterDividido(92) = "9"
CaracterDividido(93) = "8"
CaracterDividido(94) = "7"
CaracterDividido(95) = "6"
CaracterDividido(96) = "5"
CaracterDividido(97) = "4"
CaracterDividido(98) = "3"
CaracterDividido(99) = "2"

'For a = 0 To 255
'   For b = 10 To 99
'     If a <> b Then
'      If CaracterDividido(a) = CaracterDividido(b) Then
'       Text5.Text = Text5.Text & "Se repiten... " & a & " con " & b & "."
'      End If
'     End If
'   Next b
'Next a

End Sub

Function RandomizePhrase(phrase As String) As String

Randomize ' lo peor seria que sean iguales, ya que enviar dos iguales el server lo detecta como chiter -.-

Dim frase As String, aleatoriedad As Integer
Dim i As Integer, a As Integer
Dim backup As String, back As String
Dim what As String, h As Integer
Dim count As Long

count = Len(phrase)
aleatoriedad = Int(Rnd * 200)
aleatoriedad = aleatoriedad + 100

frase = aleatoriedad + Asc(Right$(phrase, 1)) + aleatoriedad
For i = 1 To count
    frase = frase & (Asc(Mid$(phrase, i, 1)) + aleatoriedad)
Next i

'back = frase
'
RandomizePhrase = frase
'
'For a = 1 To Len(frase) / 2
' 'MsgBox Mid$(back, 1, 2)
' 'Text4.Text = Text4.Text & Mid$(back, 1, 2)
' backup = backup & Chr(Mid$(back, 1, 2) + 151)
'' Text5.Text = backup
'  back = Mid$(back, 3, Len(back) - 2)
'Next a
'
'RandomizePhrase = backup

End Function
Function ReadPhrase(phrase As String) As String
On Error Resume Next

Dim i As Integer, frase As String
Dim aleatoriedad As Integer, a As Integer
Dim backup As String, back As String
Dim what As String, h As Integer
Dim count As Long

'backup = phrase
'back = phrase

'For a = 1 To Len(back)
' 'MsgBox Mid$(back, 1, 1)
' 'Text4.Text = Text4.Text & Mid$(back, 1, 1)
' backup = backup & Asc(Mid$(back, 1, 1)) - 151
'
'  back = Mid$(back, 2, Len(back) - 1)
'Next a

'phrase = backup

count = (Len(phrase) / 3) ' recordemos la operacion matem�tica ...
aleatoriedad = Left$(phrase, 3) - Right$(phrase, 3)
phrase = Right$(phrase, Len(phrase) - 3)

For i = 1 To count
    frase = frase & Chr(Mid$(phrase, (i * 3) - 2, 3) - aleatoriedad)
Next i

ReadPhrase = frase

End Function

