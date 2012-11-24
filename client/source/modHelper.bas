Attribute VB_Name = "modHelper"
Option Explicit

'Processing Data
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Public Sub WriteVar(ByRef File As String, ByRef main As String, ByRef Var As String, ByRef value As String): writeprivateprofilestring main, Var, value, File: End Sub
Public Function GetVar(ByVal File As String, ByVal main As String, ByVal Var As String) As String
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Public Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Integer) As String

Dim i As Integer
Dim UPos As Integer
Dim Separador As String * 1
Dim max As Integer
Dim Lent As Integer
Dim tControl As Integer

Lent = Len(Text)

If Lent <= 1 Then Exit Function

Separador = Chr$(SepASCII)

UPos = 0

For i = 1 To Pos - 1
    UPos = InStr(UPos + 1, Text, Separador, vbTextCompare)
Next

tControl = InStr(UPos + 1, Text, Separador, vbTextCompare) - 1

If tControl = 0 And Pos = 1 Then Exit Function

max = Abs(UPos - tControl)
ReadField = Mid$(Text, UPos + 1, IIf(tControl <= 0, Abs(UPos - Lent), max))

End Function
Public Function FileExist(ByRef File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    If Dir$(File, FileType) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    Randomize Timer
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
