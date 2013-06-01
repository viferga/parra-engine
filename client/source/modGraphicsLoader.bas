Attribute VB_Name = "modGraphicsLoader"
#If LoadingMetod = 1 Then

Option Explicit


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


