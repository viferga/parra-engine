Attribute VB_Name = "modTileEngine"
Option Explicit

Private Enum IsometricType
    Normal
    NormalRotation
    IsometricBase
    IsometricBaseRotation
    IsometricHeight
    
    '...
    
End Enum

Private Type sngRECT
    bottom As Single
    Left   As Single
    Right  As Single
    Top    As Single
End Type
    
Private Const Pi             As Single = 3.14159265358979
Private Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

  
Public Type structPos
    Map As Integer
    X   As Byte
    Y   As Byte
End Type

Private Type structGrh
    GrhIndex     As Integer
      
    FrameCounter As Single
    SpeedCounter As Single
      
    Started      As Byte
    Loops        As Integer
End Type
    
Private Type structGrhData
    FileNum   As Long       ' Numero Textura
    
    sX        As Integer    ' Left
    sY        As Integer    ' Top
    Width    As Integer    ' Right
    Height   As Integer    ' Bottom
    
    offsetx   As Integer
    offsety   As Integer
    
    NumFrames As Integer
    Frames()  As Long
    
    Speed     As Single
End Type: Private Grh() As structGrhData

Private Type mapBlock

    Graphic(1 To 4) As structGrh
    
    CharIndex As Integer
    NpcIndex  As Integer
    
    Blocked   As Byte
    Trigger   As Byte
    
    TileExit  As structPos
    
End Type: Private MapData(1 To 100, 1 To 100) As mapBlock

'Quad Draw
Private DirectxRect As D3DRECT
Public RenderRect   As RECT

'FPS Count
Public FramesPerSec As Integer
Public FramesPerSecCounter  As Long

' DirectX8 & Extras
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim DispMode As D3DDISPLAYMODE
Dim DevCaps As D3DCAPS8

Private Dx As DirectX8 'Root object
Private D3D As Direct3D8 ' Direct3D interface
Public D3DX As D3DX8 ' Helper library
Public D3DDevice As Direct3DDevice8 'Represents the hardware doing the rendering

Private Type tCache
    Number        As Long
    SrcHeight     As Single
    SrcWidth      As Single
End Type: Private Cache As tCache

' Directx8 Fonts
Private Type FontInfo
    MainFont As D3DXFont
    MainFontDesc As IFont
    MainFontFormat As New StdFont
    color As Long
End Type: Private Font() As FontInfo

'Manage Textures
Private Type structGraphic
        FileName   As Long
        D3DTexture As Direct3DTexture8
        Used       As Long
        Available  As Boolean
        Width      As Integer
        Height     As Integer
End Type: Private oGraphic()   As structGraphic

Private lKeys()      As Long
Private lSurfaceSize As Long
Private lMaxEntrys   As Long
Private Loader       As D3DX8
Private TexInfo      As D3DXIMAGE_INFO
Private nBuffer(0) As Byte

'***************************
'External Functions
'***************************

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal numBytes As Long)

'Gets number of ticks since windows started
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Very percise counter 64bit system counter
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Sub initializeIndex()
    Dim i As Long
    
      ReDim Preserve Grh(1 To GetVar(App.Path & "\Init\grh.ini", "INIT", "numGrh")) As structGrhData
    
        For i = 1 To UBound(Grh)
        
            With Grh(i)
                .FileNum = ReadField(1, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .sX = ReadField(2, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .sY = ReadField(3, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .Width = ReadField(4, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .Height = ReadField(5, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .offsetx = ReadField(6, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .offsety = ReadField(7, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                
                .NumFrames = ReadField(8, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                
                ReDim .Frames(1 To .NumFrames)
                
                If (.NumFrames < 1) Then
                    
                    Dim frameCount As Long
                    
                    For frameCount = 1 To .NumFrames
                         .Frames(frameCount) = ReadField(frameCount + 8, _
                                                GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                    Next frameCount

                Else
                
                    .Frames(1) = i
                End If
                
                .Speed = ReadField(.NumFrames + 8, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                
            End With
            
            'Add GrhList
            frmMain.grhList.AddItem "Grh" & CStr(i)
            
        Next i
    
End Sub
Private Sub initializeGrhAnim(ByRef cGrh As structGrh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)

    cGrh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If Grh(cGrh.GrhIndex).NumFrames > 1 Then
            cGrh.Started = 1
        Else
            cGrh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If Grh(cGrh.GrhIndex).NumFrames = 1 Then Started = 0
        cGrh.Started = Started
    End If
    
    
    If cGrh.Started Then
        cGrh.Loops = -1
    Else
        cGrh.Loops = 0
    End If
    
    cGrh.FrameCounter = 1
    cGrh.SpeedCounter = Grh(cGrh.GrhIndex).Speed
End Sub

Public Sub showNextFrame()

    If (testCooperative = False) Then Exit Sub

    With D3DDevice
        
        .Clear 1, DirectxRect, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 1#, 0
            .BeginScene
            
               'Render Map
                mapRender
                
               'Render grhSelected
                If frmMain.grhList.ListIndex + 1 > 0 Then
                    deviceRender frmMain.grhList.ListIndex + 1, frmMain.MouseX + 200, frmMain.MouseY + 115, frmMain.IsoType
                End If
                
               'Render Font
                fontRender CStr(FramesPerSec), 1, 3, 5, 14, 30
                                
            .EndScene
        .Present RenderRect, RenderRect, frmMain.hwnd, ByVal 0&

    End With
    
    FramesPerSecCounter = FramesPerSecCounter + 1
End Sub
Public Function testCooperative() As Boolean

    testCooperative = True

    With D3DDevice

        If .TestCooperativeLevel = D3D_OK Then Exit Function
        
                Dim h As Long
                testCooperative = False

                Select Case .TestCooperativeLevel
                
                    Case D3DERR_DEVICELOST
                            'Do a loop while device is lost
                             Do
                                For h = 1 To UBound(Font)
                                    Font(h).MainFont.OnLostDevice
                                Next h
                                    
                                'Clear All Textures
                                texDestroyAll
                                    
                                DoEvents
                            Loop While (.TestCooperativeLevel = D3DERR_DEVICELOST)
                            
                            testCooperative = False
                            Exit Function
                    
                    Case D3DERR_DEVICENOTRESET
                             Do
                                fontDeInitializing
                                
                                'Clear All Textures
                                texDestroyAll
                                
                                'Make Sure The Scene Its over, And Reset The Device
                                .Reset D3DWindow
                                    
                                'Reset Render States
                                deviceResetRenderStates
                                
                                fontInitializing (GetVar(App.Path & "\Init\Fonts.ini", "Info", "Size"))
                                
                                DoEvents
                            Loop While (.TestCooperativeLevel = D3DERR_DEVICELOST)
                            
                            testCooperative = False
                            Exit Function
                End Select
            
    End With

End Function
Public Function engineInitializing(ByRef Top As Integer, ByRef Left As Integer, ByRef Width As Integer, ByRef Height As Integer, Frm As Form, Optional ByRef BitsPerPixel As Byte = 32, Optional ByRef Windowed As Boolean = True) As Boolean

On Error GoTo ErrHandle
          
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate
    Set D3DX = New D3DX8
        
    With RenderRect
        .Top = Top
        .Left = Left
        .Right = Width
        .bottom = Height
    End With
    
    With DirectxRect
        .X1 = Left
        .X2 = Width
        .Y1 = Top
        .Y2 = Height
    End With
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    Dim D3DCreate As CONST_D3DCREATEFLAGS
    
    DevType = D3DDEVTYPE_HAL

    D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DevType, DevCaps
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
    
        If (Windowed = False) Then
        
            .Windowed = 0
        
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            
            'Fullscreen mode stuff here
            Select Case BitsPerPixel
                Case 32
                    .BackBufferFormat = D3DFMT_A8R8G8B8
                Case 24
                    .BackBufferFormat = D3DFMT_R8G8B8
                Case 16
                    .BackBufferFormat = D3DFMT_R5G6B5
            End Select
            
            .BackBufferWidth = 800
            .BackBufferHeight = 600
                        
        Else
            .Windowed = 1
            
            .BackBufferFormat = DispMode.Format
            
            .BackBufferWidth = DispMode.Width
            .BackBufferHeight = DispMode.Height
            
        End If
        
        .SwapEffect = D3DSWAPEFFECT_COPY
        
        .hDeviceWindow = Frm.hwnd
        
        .MultiSampleType = D3DMULTISAMPLE_NONE
        
        'Auto depth stencil format.. para motion blur _
                para mas fps.. usar en el select case y asignar a cada tipo el mismo que backbufferformat _
                    y desactivar el enableatudeph.. (poniendolo en 0)
                    
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = 1
        .BackBufferCount = 1
        
    End With
    
    If (DevCaps.MaxTextureHeight < 512) Or (DevCaps.MaxTextureWidth < 512) Then
        MsgBox "This sample requires a device capable of using 512x512 textures. No compatable device found.", vbCritical, "Failed Init"
    End If

    'To check the rest we use:
    If Not D3D.CheckDeviceType(D3DADAPTER_DEFAULT, DevType, DispMode.Format, DispMode.Format, 1) = D3D_OK Then
        DevType = D3DDEVTYPE_REF
    ElseIf Not D3D.CheckDeviceType(D3DADAPTER_DEFAULT, DevType, DispMode.Format, DispMode.Format, 1) = D3D_OK Then
        DevType = D3DDEVTYPE_SW
    End If
    
    'For Hardware vertex processing:
    If Not (DevCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
        D3DCreate = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    Else
        D3DCreate = D3DCREATE_HARDWARE_VERTEXPROCESSING
    End If
    
    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, Frm.hwnd, D3DCreate, D3DWindow)
    
    'Reset the device's rendering state
    deviceResetRenderStates

    fontInitializing (GetVar(App.Path & "\Init\Fonts.ini", "Info", "Size"))
    
    'Clear the back buffer
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
    
    'Initialize DB
    If (texInitialize() = False) Then GoTo ErrHandle
        
    'Initialize Index
    initializeIndex
    
    engineInitializing = True
    Exit Function
    
ErrHandle:
    MsgBox "Error al iniciar el motor grafico"
    engineInitializing = False
End Function
Private Sub deviceResetRenderStates()

    With D3DDevice
    
        'Set the shader to be used
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
        'Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_AMBIENT, D3DColorXRGB(0.5, 0.5, 0.5)
        
        'Alphas
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW 'NONE
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        
        'Para mostrar los quads
        '.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False

        'Particle engine settings
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        'Set the texture stage stats (filters)
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT
        
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_CLAMP
        .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_CLAMP
        
    End With
    
End Sub
Public Sub engineDeinitializing()

    'Destroy all textures
    texDestroyAll

    'Set no texture in the device to avoid memory leaks
    If Not D3DDevice Is Nothing Then
        D3DDevice.SetTexture 0, Nothing
    End If
    
    fontDeInitializing

    Set Dx = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    
End Sub
Private Function fontInitializing(ByRef Size As Byte) As Boolean
    Dim i As Byte
    
    ReDim Preserve Font(1 To Size) As FontInfo
    
    ' Set configuration
    For i = 1 To UBound(Font)
        With Font(i)
            .MainFontFormat.Name = GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "Name")
            .MainFontFormat.Size = GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "Size")
            .color = D3DColorARGB(ReadField(1, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(2, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(3, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(4, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")))
    
            Set .MainFontDesc = .MainFontFormat
            Set .MainFont = D3DX.CreateFont(D3DDevice, .MainFontDesc.hFont)
        End With
    Next i
    
End Function
Private Sub fontDeInitializing()
    Dim i As Byte
    
    For i = 1 To UBound(Font)
        Set Font(i).MainFont = Nothing
        Set Font(i).MainFontDesc = Nothing
        Set Font(i).MainFontFormat = Nothing
    Next i
    
End Sub
Private Sub fontRender(ByRef Text As String, ByRef Index As Byte, _
                            ByRef Top As Integer, ByRef Left As Integer, ByRef Width As Integer, ByRef Height As Integer)
                            
    Static fontRect As RECT 'This defines where it will be
    
    With fontRect
        .Top = Top + RenderRect.Top
        .Left = Left + RenderRect.Left
        .bottom = Top + Width + RenderRect.Top
        .Right = Left + Height + RenderRect.Left
    End With
    
    D3DX.DrawText Font(Index).MainFont, Font(Index).color, Text, fontRect, DT_LEFT
End Sub
Private Sub mapRender()

  '  deviceResetRenderStates
  '
  '  deviceRender 1, 200, 200, IsometricBase
  '  deviceRender 2, 64 + 200, 32 + 200, IsometricBase
  '  deviceRender 2, 200, 64 + 200, IsometricBase
  '  deviceRender 2, 200 - 64, 32 + 200, IsometricBase
  '
  '  deviceRender 4, 400, 400, IsometricType.Normal
  '
  '  deviceRender 5, 200 - 20, 20 + 200, IsometricHeight
    
End Sub
Private Sub deviceRender(ByRef GrhIndex As Long, ByRef cX As Single, ByRef cY As Single, ByRef Iso As IsometricType)
    
    If (Cache.Number <> GrhIndex) Then

        If oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture Is Nothing Then
        
            With Cache
                Set oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture = texLoad(Grh(GrhIndex).FileNum, nBuffer)
        
                .SrcHeight = oGraphic(lKeys(Grh(GrhIndex).FileNum)).Height + 1
                .SrcWidth = oGraphic(lKeys(Grh(GrhIndex).FileNum)).Width + 1
                
                Cache.Number = GrhIndex
            End With
        
        End If
    
    End If

    Static vector(3) As D3DTLVERTEX
    
        GeometryBoxType Grh(GrhIndex), cX, cY, vector, Iso
               
        D3DDevice.SetTexture 0, oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vector(0), Len(vector(0))

End Sub
Private Function GeometryBoxType(ByRef Grh As structGrhData, ByRef cX As Single, ByRef cY As Single, vertex() As D3DTLVERTEX, ByRef Iso As IsometricType)

        Select Case Iso
            Case IsometricType.Normal
                
                    With Grh
                    
                        vertex(0) = setVertex(cX, cY + .Height, 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        vertex(1) = setVertex(cX, cY, 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        vertex(2) = setVertex(cX + .Width, cY + .Height, 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        vertex(3) = setVertex(cX + .Width, cY, 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                    
                    End With
                    
                    Exit Function
                    
            Case IsometricType.NormalRotation

                    Exit Function
            Case IsometricType.IsometricBase
            
                    With Grh
                        
                        vertex(0) = setVertex(cX + .Width, cY, 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        vertex(1) = setVertex(cX + (.Width * 2), cY + (.Height * 0.5), 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        vertex(2) = setVertex(cX, cY + (.Height * 0.5), 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        vertex(3) = setVertex(cX + .Width, cY + .Height, 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        
                    End With
                    
                    Exit Function
                    
            Case IsometricType.IsometricBaseRotation
                    
                    Exit Function
                    
            Case IsometricType.IsometricHeight
            
                    With Grh
                        
                        vertex(0) = setVertex(cX, cY - (.Height * 0.5), 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        vertex(1) = setVertex(cX + .Width, cY, 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        vertex(2) = setVertex(cX, cY + (.Height * 0.5), 0, 1, D3DColorXRGB(255, 255, 255), 0, .sX / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        vertex(3) = setVertex(cX + .Width, cY + .Height, 0, 1, D3DColorXRGB(255, 255, 255), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        
                    End With
                    
                    Exit Function
        End Select
        
        
End Function
Public Function setVertex(ByRef X As Single, ByRef Y As Single, ByRef z As Single, ByRef rhw As Single, ByRef color As Long, ByRef Specular As Long, ByRef tu As Single, ByRef tv As Single) As D3DTLVERTEX
    
    With setVertex
        .sX = X
        .sY = Y
        .sz = z
        .rhw = rhw
        .color = color
        .Specular = Specular
        .tu = tu
        .tv = tv
    End With
    
End Function
Private Function texInitialize() As Boolean
On Error GoTo ErrHandle
    
    lMaxEntrys = 1000
        
    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
    
    nBuffer(0) = 0
    
    Set Loader = New D3DX8
    
    texInitialize = True
    Exit Function
ErrHandle:
    texInitialize = False
End Function
Private Function texLoad(ByRef Filenumber As Long, ByRef Buffer() As Byte) As Direct3DTexture8

    oGraphic(lKeys(Filenumber)).Used = oGraphic(lKeys(Filenumber)).Used + 1

    If (oGraphic(lKeys(Filenumber)).Available = False) Then
        If (texCreateFrom(Filenumber, Buffer) = False) Then
            Set texLoad = Nothing: Exit Function
        Else
            lSurfaceSize = lSurfaceSize - 1
        End If
    End If

    Set texLoad = oGraphic(lKeys(Filenumber)).D3DTexture

End Function
Private Function texDelete(ByRef Filenumber As Long) As Boolean
    ZeroMemory oGraphic(Filenumber), Len(oGraphic(Filenumber))
    lSurfaceSize = lSurfaceSize + 1
End Function
Private Function texCreateFrom(ByRef Filenumber As Long, ByRef Buffer() As Byte) As Boolean
    Dim i As Long
    Dim TexNum As Long
    Dim DelTex As Long

    TexNum = 0

    For i = 1 To lMaxEntrys
        If (oGraphic(i).Available = False) Then
            TexNum = i
            oGraphic(i).Available = True
            Exit For
        Else
            If (oGraphic(i).Used < 0) Then oGraphic(i).Used = 0: DelTex = i
        End If
    Next i

    If (TexNum = 0) Then
        If (texDelete(DelTex) = False) Then
            texCreateFrom = False: Exit Function
        Else
            lKeys(Filenumber) = DelTex
        End If
    Else
        lKeys(Filenumber) = TexNum
    End If

'    If Buffer(0) <> 0 Then 'Load From Memory
'        If (texFromMemory(Filenumber, Buffer) = False) Then texCreateFrom = False: Exit Function
'    Else 'Load From File
        If (texFromFile(Filenumber) = False) Then texCreateFrom = False: Exit Function
'    End If

    texCreateFrom = True
End Function
Private Function texFromFile(ByRef Filenumber As Long) As Boolean
    
    With oGraphic(lKeys(Filenumber))
        Set .D3DTexture = Loader.CreateTextureFromFileEx(D3DDevice, App.Path & "\Graphics\" & CStr(Filenumber) & ".bmp", D3DX_DEFAULT, _
                                             D3DX_DEFAULT, 0, 0, 0, D3DPOOL_MANAGED, _
                                             D3DX_FILTER_POINT, D3DX_FILTER_NONE, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)

        .Width = TexInfo.Width
        .Height = TexInfo.Height
    End With
    
    texFromFile = True

End Function

Private Function texFromMemory(ByRef Filenumber As Long, ByRef Buffer() As Byte) As Boolean
    
    With oGraphic(lKeys(Filenumber))
        Set .D3DTexture = Loader.CreateTextureFromFileInMemoryEx(D3DDevice, Buffer(0), UBound(Buffer), D3DX_DEFAULT, _
                                             D3DX_DEFAULT, 0, 0, 0, D3DPOOL_MANAGED, _
                                             D3DX_FILTER_POINT, D3DX_FILTER_NONE, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)

        .Width = TexInfo.Width
        .Height = TexInfo.Height
    End With
    
    texFromMemory = True

End Function

Public Sub texDestroyAll()
    Dim i As Long

    For i = 1 To lSurfaceSize
        With oGraphic(i)
            If (.Available = True) Then
                Set .D3DTexture = Nothing
                    .Available = False
            End If
        End With
    Next i
    
    lSurfaceSize = lMaxEntrys

    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
    
End Sub
Public Function gameGetElapsedTime() As Single
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    gameGetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
