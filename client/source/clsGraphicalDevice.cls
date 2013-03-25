VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Type Color4Byte
 Red As Byte
 Green As Byte
 Blue As Byte
 Alpha As Byte
End Type

Public Type Texture
  ID As Long
  Width As Long
  Height As Long
  TexWidth As Long
  TexHeight As Long
  Pixels() As Color4Byte
End Type

Private Type videoMode
     Width As Long
     Height As Long
     BitsPerPixel As Long
End Type

Public Enum DEVICE_NO_ERROR
    DEVICE_CTX_VALID
    DEVICE_CTX_RESET
    DEVICE_CTX_LOST
    DEVICE_CTX_ERROR
End Enum

Public Enum CLEAR_FLAGS
    CLEAR_COLOR = 1
    CLEAR_DEPTH = 2
    CLEAR_STENCIL = 4
End Enum

Private VideoModeList() As videoMode
Private VideoModeListSize As Long

Public Function Initialize(ByRef Top As Integer, ByRef Left As Integer, ByRef Width As Integer, ByRef Height As Integer, Frm As Form, Optional ByRef BitsPerPixel As Byte = 32, Optional ByRef Windowed As Boolean = True) As Boolean

End Function

Private Function VideoModeGetDesktopMode() As videoMode

End Function

Private Function VideoModeGetAvailableModes() As videoMode()

End Function

Public Sub DeInitialize()

End Sub

Public Function DeviceIsContextValid() As DEVICE_NO_ERROR

End Function

Public Sub BeginScene(ClearRect As Rect, flags As CLEAR_FLAGS)

End Sub

Public Sub EndScene(DestRect As Rect, hwnd As Long)

End Sub

Public Function TextureGet(ByVal index As Long, ByVal Width, ByVal Height) As Texture

End Function

Public Function TextureBind(ByVal index As Long) As Boolean

End Function

Public Function TextureBindEx(ByRef Texture As Texture) As Boolean

End Function

