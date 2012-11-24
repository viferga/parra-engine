Attribute VB_Name = "modPhysics"
Option Explicit

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
