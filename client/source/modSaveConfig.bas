Attribute VB_Name = "modSaveConfig"
Option Explicit

' *********************************************************************
' Autor: Emanuel Matias (Dunkan)
' Nota: Almacena 8 Bool (16 Bytes) en 1 Byte.
'       Se podria ademas hacer esto con un long y almacenar 32 variables _
        (Innecesario por ahora)
'       ¡¡¡Esto es solo una prueba!!!
' *********************************************************************

' ****************************************
'   Private Type StructConfig
'       OGL As Boolean '¿Use OGL?
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
            MsgBox "True, option nº: " & (Ind + 1)
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

