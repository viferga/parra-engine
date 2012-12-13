Attribute VB_Name = "modGeneral"
Option Explicit

'Win32 Declarations
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Sub InitializeRandom(): Randomize Timer: End Sub
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long: RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound: End Function
Public Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean: FileExist = LenB(Dir$(file, FileType)) <> 0: End Function
Public Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String): writeprivateprofilestring Main, Var, value, file: End Sub
Public Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
    Dim sSpaces As String  ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100)
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Public Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim Delimiter As String * 1
    
    Delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, Delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos) Else _
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
End Function
Public Sub Consola(ByRef Text As String)

    frmMain.Consola.AddItem Text

End Sub
