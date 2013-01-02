Attribute VB_Name = "modIsoTileEngine"
Option Explicit

Public WireFrame As Boolean

Public Color(1) As D3DCOLORVALUE

Public Const EngineWidth As Integer = 800
Public Const EngineHeight As Integer = 600

Private Const TileBufferSize As Integer = 2

Private Enum IsometricType
    Normal
    NormalRotation
    IsometricBase
    IsometricBaseRotation
    IsometricHeight
    
    '...
    
End Enum
    
Private Const Pi             As Single = 3.14159265358979
Private Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

Public Type structGrh
    GrhIndex     As Long
      
    FrameCounter As Single
    SpeedCounter As Single
      
    Started      As Byte
    Loops        As Integer
End Type
    
Private Type structGrhData
    FileNum   As Long      ' Numero Textura
    
    sX        As Integer   ' Left
    sY        As Integer   ' Top
    Width    As Integer    ' Right
    Height   As Integer    ' Bottom
    
    offsetX   As Integer
    offsetY   As Integer
    
    NumFrames As Integer
    Frames()  As Long
    
    Speed     As Single
End Type: Public Grh() As structGrhData

Public Enum eDirection
    NorthEast = 1
    NorthWest = 2
    SouthEast = 3
    SouthWest = 4
    North = 5
    South = 6
    East = 7
    West = 8
End Enum

Public Type Character
    Active As Byte
    Pos As structPositionInt
    
    Body As Integer
    Head As Integer
    Heading As eDirection

    'Body As BodyData
    'Head As HeadData
    'Casco As HeadData
    'Arma As WeaponAnimData
    'Escudo As ShieldAnimData
    
    FX As structGrh
    FXIndex As Integer
        
    Name As String
    
    Moving As Byte
    scrollDirection As structPositionInt
    MoveOffset As structPositionSng
End Type

Public characterList() As Character
Public charLast As Integer

Public ScrollPixelsPerFrame As structPositionInt

Public UserMoving As Byte
Public UserPos As structPositionInt
Public AddtoUserPos As structPositionInt
Public playerCharIndex As Integer

'Quad Draw
Public RenderRect As RECT

'FPS Count
Public FramesPerSec As Integer
Public FramesPerSecCounter  As Long

' DirectX8 & Extras
Public D3DWindow As DxVBLibA.D3DPRESENT_PARAMETERS
Public DispMode As DxVBLibA.D3DDISPLAYMODE
Public DevCaps As DxVBLibA.D3DCAPS8

Public dX As DxVBLibA.DirectX8 'Root object
Public D3D As DxVBLibA.Direct3D8 ' Direct3D interface
Public D3DX As DxVBLibA.D3DX8 ' Helper library
Public D3DDevice As DxVBLibA.Direct3DDevice8 'Represents the hardware doing the rendering

Private g_Adapters() As D3DUTIL_ADAPTERINFO      ' Array of Adapter infos
Private g_lCurrentAdapter As Long                ' current adapter (index into infos)
Private g_lNumAdapters As Long                   ' size of the g_Adapters array
Private g_EnumCallback As Object                 ' object that defines VerifyDevice function

Private g_behaviorflags As Long                  ' Current VertexProcessing (hardware or software)
Private g_focushwnd As Long                      ' Current focus window handle
Private g_lWindowWidth As Long                   ' backbuffer width of windowed state
Private g_lWindowHeight As Long                   ' backbuffer height of windowed state

Private D3DDeviceType As CONST_D3DDEVTYPE

' Directx8 Fonts
Private Type FontInfo
    MainFont As DxVBLibA.D3DXFont
    MainFontDesc As IFont
    MainFontFormat As New StdFont
    Color As Long
End Type: Private Font() As FontInfo

' Vector Usado para los Quads
Public Vector(3) As D3DTLVERTEX

' INDEX BUFFERS
Public vbQuadIdx As DxVBLibA.Direct3DVertexBuffer8
Public ibQuad As DxVBLibA.Direct3DIndexBuffer8
Public indexList(0 To 5) As Integer 'the 6 indices required (note that the number is the
                              'same as the vertex count in the previous version).
'for motion blurring
Public m_pDisplayTexture As DxVBLibA.Direct3DTexture8
Public m_pDisplayTextureSurface As DxVBLibA.Direct3DSurface8
Public m_pDisplayZSurface As DxVBLibA.Direct3DSurface8
Public m_pBackBuffer As DxVBLibA.Direct3DSurface8
Public m_pZBuffer As DxVBLibA.Direct3DSurface8

Public VertList(0 To 3) As D3DTLVERTEX

Public errMotion     As Boolean
Public MotionBlur    As Boolean
Public lBlurFactor   As Byte

Public BasicColor(3) As Long

' ElapsedTime
Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single

'***************************
'External Functions
'***************************

'Gets number of ticks since windows started
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Very percise counter 64bit system counter
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef RECT As RECT) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef RECT As RECT) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hwndafter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal options As Long) As Long
'Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal newVal As Long) As Long
'Private Declare Function GetWindowLongA Lib "user32.dll" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Sub initializeIndex()
    Dim i As Long
    
      ReDim Preserve Grh(1 To GetVar(App.Path & "\Init\grh.ini", "INIT", "numGrh")) As structGrhData
    
        For i = 1 To UBound(Grh)
        
            With Grh(i)
                .FileNum = ReadField(1, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .sX = ReadField(2, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .sY = ReadField(3, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .Width = ReadField(4, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .Height = ReadField(5, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .offsetX = ReadField(6, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                .offsetY = ReadField(7, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(i)), Asc("-"))
                                
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
            
            #If WorldEditor = 1 Then
                'Add GrhList
                frmMain.grhList.AddItem "Grh" & CStr(i)
            #End If
        Next i
    
End Sub
Private Sub initializeGrhAnim(ByRef cGrh As structGrh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'GDK: sin uso

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

#If ParticleEditor = 1 Then
    If EditParticle = False Then
#End If

    Static OffsetCounter As structPositionSng

    If UserMoving Then
    
    '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounter.x = OffsetCounter.x - ScrollPixelsPerFrame.x * AddtoUserPos.x * timerTicksPerFrame
            If Abs(OffsetCounter.x) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                OffsetCounter.x = 0
                AddtoUserPos.x = 0
                UserMoving = False
            End If
        End If
                    
    '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounter.y = OffsetCounter.y - ScrollPixelsPerFrame.y * AddtoUserPos.y * timerTicksPerFrame
            If Abs(OffsetCounter.y) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                OffsetCounter.y = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
        
    End If
    
#If ParticleEditor = 1 Then
    End If
#End If

    If (testCooperative = False) Then Exit Sub

    With D3DDevice
        
        If MotionBlur = True And errMotion = False Then
            .SetRenderTarget m_pDisplayTextureSurface, m_pDisplayZSurface, 0
            .Clear 1, RenderRect, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
        Else
            .Clear 1, RenderRect, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
        End If
        
            .BeginScene
            
            #If ParticleEditor = 1 Then
                If EditParticle = False Then
            #End If
        
               'Render Map
                mapRender OffsetCounter

                fontRender CStr("X: " & CStr(UserPos.x) & " Y: " & CStr(UserPos.y)), 2, 300, 0, 130, 20, DT_LEFT
            
            #If ParticleEditor = 1 Then
                End If
            #End If
                
                fontRender CStr("FPS:" & CStr(FramesPerSec)), 3, 1, 0, 90, 12, DT_LEFT
                
                #If WorldEditor = 1 Then
                    If EditMap = True Then
                        fontRender CStr("MouseX: " & CStr(MouseTilesPos.x) & " MouseY: " & CStr(MouseTilesPos.y)), 2, 560, 1, 240, 20, DT_LEFT
                        
                        'Render grhSelected
                        If frmMain.grhList.ListIndex + 1 > 0 Then
                            deviceRenderTexture frmMain.grhList.ListIndex + 1, Mouse.x, Mouse.y, BasicColor(), frmMain.cmbMode.ListIndex
                        End If
                    End If
                #End If
                
                'Render Gui
                If RenderGUI = True Then
                    guiRender
                End If
                
                If MotionBlur = True And errMotion = False Then ResetMotionStates
            
            .EndScene
        .Present RenderRect, ByVal 0&, 0, ByVal 0&

    End With
    
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = gameGetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * 0.018   ' Engine Speed
    
End Sub
Public Function testCooperative() As Boolean

    testCooperative = True

    With D3DDevice

        If .TestCooperativeLevel = D3D_OK Then Exit Function
        
                Dim H As Long
                testCooperative = False

                Select Case .TestCooperativeLevel
                
                    Case D3DERR_DEVICELOST
                            'Do a loop while device is lost
                             Do
                                For H = 1 To UBound(Font)
                                    Font(H).MainFont.OnLostDevice
                                Next H
                                    
                                'Clear All Textures
                                #If LoadingMetod = 0 Then
                                    texReloadAll
                                #Else
                                    surfaceTerminate
                                #End If
                                    
                                DoEvents
                            Loop While (.TestCooperativeLevel = D3DERR_DEVICELOST)
                            
                            testCooperative = False
                            Exit Function
                    
                    Case D3DERR_DEVICENOTRESET
                             Do
                                fontDeInitializing
                                                                
                                'Make Sure The Scene Its over, And Reset The Device
                                .Reset D3DWindow
                                
                                'Clear All Textures
                                #If LoadingMetod = 0 Then
                                    texReloadAll
                                #Else
                                    surfaceTerminate
                                #End If
                                    
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

On Error GoTo errHandle

    ' Initialize the DirectX8 and d3dx8 objects
    If dX Is Nothing Then Set dX = New DirectX8
    If D3DX Is Nothing Then Set D3DX = New D3DX8
    
    ' Create the Direct3D object
    Set D3D = dX.Direct3DCreate
    
    ' Call the sub that builds a list of available adapters,
    ' adapter device types, and display modes
    Call D3DEnum_BuildAdapterList(Frm)
    
    If Windowed = False Then
        SetWindowPos frmConnect.hwnd, 0, 0, 0, 800, 600, 0
    End If
        
    'With RenderRect
    '    .Y1 = Top
    '    .X1 = Left
    '    .X2 = Width
    '    .Y2 = Height
    'End With

    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    
    If Windowed Then
        engineInitializing = engineInitializingInWindow(Frm.hwnd, 0, DevType, True)
    Else
        engineInitializing = engineInitializingInFullscreen(Frm.hwnd, 0, 0, DevType, True, BitsPerPixel)
    End If
    

    'Dim D3DCreate As CONST_D3DCREATEFLAGS
    
    '

    'D3D.GetDeviceCaps D3DADAPTER_DEFAULT, DevType, DevCaps
    
   ' With D3DWindow
    
       ' If (Windowed = False) Then
        
           ' .Windowed = 0
        
           ' DispMode.Width = 800
           ' DispMode.Height = 600
        
           ' .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
           ' .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            
            'Fullscreen mode stuff here
           ' Select Case BitsPerPixel
            '    Case 32
            '        DispMode.Format = D3DFMT_A8R8G8B8
            '    Case 24
           '         DispMode.Format = D3DFMT_R8G8B8
           '     Case 16
           '         DispMode.Format = D3DFMT_R5G6B5
           ' End Select
                        
       ' Else
        '    .Windowed = 1
            
         '   D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
            
        'End If
            
       ' .BackBufferWidth = DispMode.Width
       ' .BackBufferHeight = DispMode.Height
       ' .BackBufferFormat = DispMode.Format
       ' .BackBufferCount = 1
        
       ' .hDeviceWindow = Frm.hwnd

       ' .SwapEffect = D3DSWAPEFFECT_COPY
       ' .MultiSampleType = D3DMULTISAMPLE_NONE
        
        'To check the rest we use:
       ' If Not D3D.CheckDeviceType(D3DADAPTER_DEFAULT, DevType, DispMode.Format, DispMode.Format, 1) = D3D_OK Then
       '     DevType = D3DDEVTYPE_REF
       ' ElseIf Not D3D.CheckDeviceType(D3DADAPTER_DEFAULT, DevType, DispMode.Format, DispMode.Format, 1) = D3D_OK Then
       '     DevType = D3DDEVTYPE_SW
       ' End If
                            
       ' If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, DevType, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
            'We can use a 16 bit Z-Buffer
        '    .AutoDepthStencilFormat = D3DFMT_D16 '//16 bit Z-Buffer
       ' Else
       '     MsgBox "Error: 16 bit Z-Buffer not suported"
       ' End If

       '' .EnableAutoDepthStencil = 1
        
   ' End With
    
    'For Hardware vertex processing:
    'If Not (DevCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
   '     D3DCreate = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    'Else
    '    D3DCreate = D3DCREATE_HARDWARE_VERTEXPROCESSING
    'End If
    
    'Set the D3DDevices
   ' If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
   ' Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, Frm.hwnd, D3DCreate, D3DWindow)
    
    'Ocultamos el form
    Frm.Visible = False
    
    'Reset the device's rendering state
    deviceResetRenderStates

    fontInitializing (GetVar(App.Path & "\Init\Fonts.ini", "Info", "Size"))
    
    'Clear the back buffer
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
    
    'Initialize DB
    #If LoadingMetod = 0 Then
        If (texInitialize() = False) Then GoTo errHandle
    #Else
        surfaceInitialize App.Path & "\Graphics\", D3DX, D3DDevice, 30
    #End If
    
    'Initialize Index
    initializeIndex
    
   ' Initialize Motion Blur
    
    If (initializeMotionBlur() = False) Then MsgBox "Error al iniciar el Motion Blur. " & _
                                                                    "La aplicacion se iniciara sin el."
    ' Particle Load
    loadParticleGroup
        
    'Load Index List
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5

    Set ibQuad = D3DDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
        
    ' Index Quad
    Set vbQuadIdx = D3DDevice.CreateVertexBuffer(Len(Vector(0)) * 4, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    
    'Basic color for graphics
    BasicColor(0) = -1: BasicColor(1) = -1: BasicColor(2) = -1: BasicColor(3) = -1
    
    'Engine light color
    Color(0).r = 80: Color(0).g = 80: Color(0).b = 220
    Color(1).r = 250: Color(1).g = 250: Color(1).b = 200
    
    'Pixelsperframe for tile engine
    ScrollPixelsPerFrame.x = 8
    ScrollPixelsPerFrame.y = 8
    
    'Create a pixel shader
    cPixelShader = pixelShaderMakeFromMemory(psOriginalColor)
    If cPixelShader > 0 Then D3DDevice.SetPixelShader cPixelShader
    
    'Initialize gui
    If guiInitialize = False Then GoTo errHandle
    
    engineInitializing = True
    Exit Function
    
errHandle:

    If Err.Number = 429 Then
      MsgBox "No se puede iniciar el motor grafico, ya que no ha sido detectado DirectX 8. Reinstalalo", vbCritical
    Else
        MsgBox "Error al iniciar el motor grafico" & vbNewLine _
        & "ErrNumber: " & Err.Number & vbNewLine _
        & "ErrDescription: " & Err.Description, vbOKOnly, "Error"
    End If

    engineInitializing = False

End Function
Private Function engineInitializingInWindow(hwnd As Long, AdapterIndex As Long, DevType As CONST_D3DDEVTYPE, bTryFallbacks As Boolean) As Boolean
   
    On Error GoTo errOut
    
    'save the current adapter
    g_lCurrentAdapter = AdapterIndex
    
    Dim d3ddm As D3DDISPLAYMODE

    ' Initialize the present parameters structure
    ' to use 1 back buffer and a 16 bit depth buffer
    ' change the autoDepthStencilFormat if you need stencil bits
    With D3DWindow
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = hwnd
    End With
        
    g_focushwnd = hwnd
    
    Dim rc As RECT
    Call GetClientRect(g_focushwnd, rc)
    RenderRect.Right = rc.Right - rc.Left
    RenderRect.Bottom = rc.Bottom - rc.Top
    Call GetWindowRect(g_focushwnd, RenderRect)
    
    
    With g_Adapters(g_lCurrentAdapter)
    
        ' If running windowed, set the current desktop format
        ' as the format the device will use.
        ' note the the current mode and backbuffer width and height
        ' information is ignored by d3d
        Call D3D.GetAdapterDisplayMode(g_lCurrentAdapter, d3ddm)
        
        ' figure out if this format supports hardware acceleration
        ' by looking it up in our format list
        g_behaviorflags = D3DEnum_FindInFormatList(g_lCurrentAdapter, DevType, d3ddm.Format)
        If g_behaviorflags <= 0 Then g_behaviorflags = D3DEnum_CheckFormatCompatibility(AdapterIndex, DevType, d3ddm.Format, False, False)

        
        
        D3DWindow.BackBufferFormat = d3ddm.Format
        D3DWindow.BackBufferWidth = 0
        D3DWindow.BackBufferHeight = 0
        
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD
        
                
        .bWindowed = True
        .DeviceType = DevType
        
        D3DDeviceType = DevType
       
    End With
            
    'Try to create the device now that we have everything set.
    On Local Error Resume Next
    Set D3DDevice = D3D.CreateDevice(g_lCurrentAdapter, DevType, g_focushwnd, g_behaviorflags, D3DWindow)

    If Err.Number Then

        If bTryFallbacks = False Then Exit Function
        
        'If a HAL device was being attempted, try again with a REF device instead.
        If D3DDeviceType = D3DDEVTYPE_HAL Then
            Err.Clear

            'Make sure the user knows that this is less than an optimal 3D environment.
            MsgBox "No hardware support found. Switching to reference rasterizer.", vbInformation
            
            'reset our variable to use ref
            g_Adapters(g_lCurrentAdapter).DeviceType = D3DDEVTYPE_REF
            g_Adapters(g_lCurrentAdapter).bReference = True
            D3DDeviceType = D3DDEVTYPE_REF
            Set D3DDevice = D3D.CreateDevice(g_lCurrentAdapter, D3DDEVTYPE_REF, g_focushwnd, g_behaviorflags, D3DWindow)
            
        End If
            
    End If

    If Err.Number Then
        
        'The app still hit an error. Both HAL and REF devices weren't created. The app will have to exit at this point.
        MsgBox "No suitable device was found to initialize D3D. Application will now exit.", vbCritical
        engineInitializingInWindow = False
        End
        Exit Function

    End If

    'update our device caps data
    D3DDevice.GetDeviceCaps DevCaps
    
    'set any state we need to initialize
    'D3DXMatrixIdentity g_identityMatrix 'GDK: sirve?
    
    'set the reference flag if we choose a ref device
    With g_Adapters(g_lCurrentAdapter)
        If .DeviceType = D3DDEVTYPE_REF Then
            .bReference = True
        Else
            .bReference = False
        End If
    End With
    
    engineInitializingInWindow = True
    Exit Function
    
errOut:
    Debug.Print "Failed Engine Initiation"

End Function
Private Function engineInitializingInFullscreen(hwnd As Long, AdapterIndex As Long, modeIndex As Long, DevType As CONST_D3DDEVTYPE, bTryFallbacks As Boolean, BitsPerPixel As Byte) As Boolean
    
    On Error GoTo errOut
    
    Dim ModeInfo As D3DUTIL_MODEINFO

    Dim rc As RECT
    
    'save the current adapter
    g_lCurrentAdapter = AdapterIndex
        
    
    g_focushwnd = hwnd
    
    
    'switching from windowed to fullscreen so save height and width
    If D3DWindow.Windowed = 1 Then
        Call GetClientRect(g_focushwnd, rc)
        g_lWindowWidth = rc.Right - rc.Left
        g_lWindowHeight = rc.Bottom - rc.Top
        Call GetWindowRect(g_focushwnd, RenderRect)
    End If

    
    ' Initialize the present parameters structure
    ' to use 1 back buffer and a 16 bit depth buffer
    ' change the autoDepthStencilFormat if you need stencil bits
    With D3DWindow
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = g_focushwnd
    End With
    

    
    'Fullscreen mode stuff here
    Select Case BitsPerPixel
        Case 32
            ModeInfo.Format = D3DFMT_A8R8G8B8
        Case 24
            ModeInfo.Format = D3DFMT_R8G8B8
        Case 16
            ModeInfo.Format = D3DFMT_R5G6B5
    End Select

    
    With g_Adapters(g_lCurrentAdapter)
            
        With .DevTypeInfo(DevType)
            ModeInfo = .Modes(modeIndex)
            g_behaviorflags = .Modes(modeIndex).VertexBehavior
            
            D3DWindow.BackBufferWidth = ModeInfo.lWidth
            D3DWindow.BackBufferHeight = ModeInfo.lHeight
            D3DWindow.BackBufferFormat = ModeInfo.Format
            D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
            D3DWindow.Windowed = 0
            D3DWindow.FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            
            .lCurrentMode = modeIndex
        End With
             

        .bWindowed = False
        .DeviceType = DevType
        
        If g_behaviorflags <= 0 Then g_behaviorflags = D3DEnum_CheckFormatCompatibility(AdapterIndex, DevType, ModeInfo.Format, False, False)
        D3DDeviceType = DevType
        
    End With
            
    'Try to create the device now that we have everything set.
    On Local Error Resume Next
    Set D3DDevice = D3D.CreateDevice(g_lCurrentAdapter, DevType, g_focushwnd, g_behaviorflags, D3DWindow)

    If Err.Number Then
        If bTryFallbacks = False Then Exit Function

        'If a HAL device was being attempted, try again with a REF device instead.
        If D3DDeviceType = D3DDEVTYPE_HAL Then
            Err.Clear

            'Make sure the user knows that this is less than an optimal 3D environment.
            MsgBox "No hardware support found. Switching to reference rasterizer.", vbInformation
            
            'reset our variable to use ref
            g_Adapters(g_lCurrentAdapter).DeviceType = D3DDEVTYPE_REF
            D3DDeviceType = D3DDEVTYPE_REF
            Set D3DDevice = D3D.CreateDevice(g_lCurrentAdapter, D3DDEVTYPE_REF, g_focushwnd, g_behaviorflags, D3DWindow)
            
        End If
            
    End If


    If Err.Number Then
        
        'The app still hit an error. Both HAL and REF devices weren't created. The app will have to exit at this point.
        MsgBox "No suitable device was found to initialize D3D. Application will now exit.", vbCritical
        engineInitializingInFullscreen = False
        End
        Exit Function

    End If

    'update our device caps data
    D3DDevice.GetDeviceCaps DevCaps
    
    'set any state we need to initialize
    'D3DXMatrixIdentity g_identityMatrix
    
    'set the reference flag if we choose a ref device
    With g_Adapters(g_lCurrentAdapter)
        If .DeviceType = D3DDEVTYPE_REF Then
           .bReference = True
        Else
           .bReference = False
        End If
    End With
    engineInitializingInFullscreen = True
    Exit Function
    
errOut:
    Debug.Print "Failed Engine Init at Fullscreen"
End Function

Public Sub engineDeinitializing()
    Dim emptycaps As D3DCAPS8
    Dim emptyrect As RECT
    Dim emptypresent As D3DPRESENT_PARAMETERS
    
    guiDestroy

    If cPixelShader > 0 Then D3DDevice.SetPixelShader 0
    pixelShaderDelete cPixelShader

    Erase indexList
    Erase Vector

    'Index Buffers
    Set vbQuadIdx = Nothing
    Set ibQuad = Nothing


    MotionBlur = False

    Set m_pDisplayTexture = Nothing
    Set m_pDisplayZSurface = Nothing
    Set m_pBackBuffer = Nothing
    Set m_pZBuffer = Nothing
    Set m_pDisplayTextureSurface = Nothing

 
    'Destroy all textures
    #If LoadingMetod = 0 Then
        texDestroyAll
    #Else
        surfaceTerminate
    #End If

    'Set no texture in the device to avoid memory leaks

    If Not D3DDevice Is Nothing Then
        D3DDevice.SetTexture 0, Nothing
    End If

    fontDeInitializing

    Set dX = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set g_EnumCallback = Nothing
    
    g_focushwnd = 0
    g_behaviorflags = 0
    g_lWindowWidth = 0
    g_lWindowHeight = 0
   
    DevCaps = emptycaps
    D3DWindow = emptypresent
    RenderRect = emptyrect
    
    ReDim g_Adapters(0)
    g_lNumAdapters = 0

End Sub

Public Function initializeMotionBlur() As Boolean

On Error GoTo errHandle

    Dim TexSizeW As Integer, TexSizeH As Integer
    
    TexSizeW = 800: TexSizeH = 600

    'Configure MotionBlur code
    Set m_pDisplayTexture = D3DX.CreateTexture(D3DDevice, TexSizeW, TexSizeH, 1, D3DUSAGE_RENDERTARGET, DispMode.Format, D3DPOOL_DEFAULT)
    Set m_pDisplayZSurface = D3DDevice.CreateDepthStencilSurface(TexSizeW, TexSizeH, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set m_pBackBuffer = D3DDevice.GetRenderTarget()
    Set m_pZBuffer = D3DDevice.GetDepthStencilSurface()
    Set m_pDisplayTextureSurface = m_pDisplayTexture.GetSurfaceLevel(0)
    
    VertList(0).sX = -1: VertList(0).sY = -1
    VertList(1).sX = RenderRect.Right: VertList(1).sY = -1
    VertList(2).sX = -1: VertList(2).sY = RenderRect.Bottom
    VertList(3).sX = RenderRect.Right: VertList(3).sY = RenderRect.Bottom
    
    VertList(0).rhw = 1: VertList(1).rhw = 1: VertList(2).rhw = 1: VertList(3).rhw = 1
    
    'Chose colors of Motion Blur
    VertList(0).Color = D3DColorXRGB(255, 255, 255)
    VertList(1).Color = D3DColorXRGB(255, 255, 255)
    VertList(2).Color = D3DColorXRGB(255, 255, 255)
    VertList(3).Color = D3DColorXRGB(255, 255, 255)
    
    'we need to adjust texcoords to factor in that we're not using ALL of the texture
    VertList(0).tu = 0#: VertList(0).tv = 0#
    VertList(1).tu = RenderRect.Right / TexSizeW: VertList(1).tv = 0#
    VertList(2).tu = 0#: VertList(2).tv = RenderRect.Bottom / TexSizeH
    VertList(3).tu = RenderRect.Right / TexSizeW: VertList(3).tv = RenderRect.Bottom / TexSizeH
    
    lBlurFactor = 10
    
    initializeMotionBlur = True
    errMotion = False
    
    Exit Function

errHandle:

    initializeMotionBlur = False
    errMotion = True

End Function
Public Sub deviceResetRenderStates()

    With D3DDevice
    
        'Set the shader to be used
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX2 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
        'Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        '.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(0.5, 0.5, 0.5)
        
        .SetRenderState D3DRS_INDEXVERTEXBLENDENABLE, 1
        
        'Alphas
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1            'For The Particle Engine
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0             'Also For The Particle Engine
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        'Sets up properties for transparency.
        '.SetRenderState D3DRS_ALPHAREF, 255
        '.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
        
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW 'NONE
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        
        'Para mostrar los quads
        If WireFrame Then
            .SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        End If
        
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

        If ShowShader Then
            .SetPixelShader cPixelShader
        Else
            .SetPixelShader 0
        End If

    End With
    
End Sub
Private Sub ResetMotionStates()
        
        With D3DDevice
        
            If cPixelShader > 0 Then D3DDevice.SetPixelShader 0
            
            .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
            .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                
            .SetRenderTarget m_pBackBuffer, m_pZBuffer, 0
            .SetTexture 0, m_pDisplayTexture
            .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
            .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(lBlurFactor, 255, 255, 255)
            
            .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            .SetRenderState D3DRS_ALPHABLENDENABLE, True
            
            .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
            .SetRenderState D3DRS_ZENABLE, 0
                
            .SetVertexShader (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
            .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertList(0), Len(VertList(0))
                
            .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
            .SetRenderState D3DRS_ZENABLE, 1
            .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            .SetRenderState D3DRS_ALPHABLENDENABLE, True
            .SetTexture 0, Nothing
            .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
            
        End With
End Sub

Public Function fontInitializing(ByRef Size As Byte) As Boolean
    Dim i As Byte
    
    ReDim Preserve Font(1 To Size) As FontInfo
    
    ' Set configuration
    For i = 1 To UBound(Font)
        With Font(i)
            .MainFontFormat.Name = GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "Name")
            .MainFontFormat.Size = GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "Size")
            .Color = D3DColorARGB(ReadField(1, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(2, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(3, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")), _
                                                  ReadField(4, GetVar(App.Path & "\Init\Fonts.ini", "Font" & CStr(i), "ARGB"), Asc("-")))
    
            Set .MainFontDesc = .MainFontFormat
            Set .MainFont = D3DX.CreateFont(D3DDevice, .MainFontDesc.hFont)
        End With
    Next i
    
End Function
Public Sub fontDeInitializing()
    Dim i As Byte
    
    For i = 1 To UBound(Font)
        Set Font(i).MainFont = Nothing
        Set Font(i).MainFontDesc = Nothing
        Set Font(i).MainFontFormat = Nothing
    Next i
    
End Sub
Private Sub fontRender(ByRef Text As String, ByRef Index As Byte, _
                            ByRef x As Integer, ByRef y As Integer, _
                            ByRef Width As Integer, ByRef Height As Integer, _
                            Format As Long)
                            
    Static fontRect As RECT 'This defines where it will be
    
    With fontRect
        .Top = y + RenderRect.Top
        .Left = x + RenderRect.Left
        .Bottom = y + Height + RenderRect.Top
        .Right = x + Width + RenderRect.Left
    End With
    
    D3DX.DrawText Font(Index).MainFont, Font(Index).Color, Text, fontRect, Format
End Sub
Public Sub Move(ByVal Direction As eDirection)
    
    Dim x As Integer
    Dim y As Integer
    
    'Figure out which way to move
    Select Case Direction
        Case eDirection.NorthEast: y = -1: x = 1
        
        Case eDirection.NorthWest: y = -1: x = -1
        
        Case eDirection.SouthEast: y = 1: x = 1
        
        Case eDirection.SouthWest: y = 1: x = -1
        
        Case eDirection.North: y = -1: x = 0
        
        Case eDirection.South: y = 1: x = 0
        
        Case eDirection.East: y = 0: x = 1
        
        Case eDirection.West: y = 0: x = -1
        
    End Select
    
    Dim PositionOk As Boolean
    
    PositionOk = mapLegalPos(UserPos.x + x, UserPos.y + y)
    
    If PositionOk Then 'and usernot paralizate, etc..
        MoveChar playerCharIndex, Direction
        MoveScreen Direction
    Else
        If characterList(playerCharIndex).Heading <> Direction Then
            'writechangeheading direction
        End If
    End If
    
End Sub
Private Sub MoveChar(ByRef characterIndex As Integer, Direction As eDirection)

    Dim addX As Integer, addY As Integer
    Dim x As Integer, y As Integer
    
    With characterList(characterIndex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case Direction
        
            Case eDirection.NorthEast
                addY = -1
                addX = 1

            Case eDirection.NorthWest
                addX = -1
                addY = -1
                
            Case eDirection.SouthEast
                addY = 1
                addX = 1
            
            Case eDirection.SouthWest
                addX = -1
                addY = 1
                
            Case eDirection.North
                addY = -1
                addX = 0

            Case eDirection.South
                addY = 1
                addX = 0
                
            Case eDirection.East
                addY = 0
                addX = 1
            
            Case eDirection.SouthWest
                addY = 0
                addX = -1
                
        End Select
        
        mapData(x + addX, y + addY).CharIndex = characterIndex
        .Pos.x = x + addX
        .Pos.y = y + addY
        mapData(x, y).CharIndex = 0
        
        .MoveOffset.x = -1 * (TilePixelWidth * addX)
        .MoveOffset.y = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = Direction
        
        .scrollDirection.x = addX
        .scrollDirection.y = addY
    End With
    
    'If uStats.Estado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    
    If mapInBounds(x + addX, y + addY) = False Then
        'charactererase characterIndex
    End If

End Sub
Private Sub MoveScreen(Direction As eDirection)
    
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
        Select Case Direction
        
            Case eDirection.NorthEast: y = -1: x = 1
            
            Case eDirection.NorthWest: y = -1: x = -1
            
            Case eDirection.SouthEast: y = 1: x = 1
            
            Case eDirection.SouthWest: y = 1: x = -1
            
            Case eDirection.North: y = -1: x = 0
            
            Case eDirection.South: y = 1: x = 0
            
            Case eDirection.East: y = 0: x = 1
            
            Case eDirection.West: y = 0: x = -1
    
                
        End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If mapInBounds(tX, tY) = False Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        AddtoUserPos.y = y
        UserPos.x = tX
        UserPos.y = tY
        UserMoving = 1
    End If
End Sub
Private Sub mapRender(ByRef PixelOffset As structPositionSng)
        
        Dim lX As Long, lY As Long ' position tiles
        Dim tX As Integer, tY As Integer ' inicio tiles
        Dim tX2 As Integer, tY2 As Integer ' fin tiles
        Dim tempX As Single, tempY As Single ' temp position
        Dim offX As Single, offY As Single ' temp offset
        
        If PixelOffset.x <> 0 Then
            If PixelOffset.x < 0 Then
                offX = 64 + PixelOffset.x
            Else
                offX = -64 + PixelOffset.x
            End If
        End If
        
        If PixelOffset.y <> 0 Then
            If PixelOffset.y < 0 Then
                offY = 64 + PixelOffset.y
            Else
                offY = -64 + PixelOffset.y
            End If
        End If
        
        
        ' Controla el tamaño de tX & tY
        
        tX = UserPos.x
        tY = UserPos.y
        
        If tY - TileBufferSize < 1 Then
            tY = 1 + TileBufferSize
        End If
        
        If tX - TileBufferSize < 1 Then
            tX = 1 + TileBufferSize
        End If
        
        If tY + TileBufferSize > MaxTilesY Then
            tY = MaxTilesY - TileBufferSize
        End If
        
        If tX + TileBufferSize > MaxTilesY Then
            tX = MaxTilesY - TileBufferSize
        End If
        
        ' Controla el tamaño de tX2 & tY2
        
        tX2 = tX + (EngineWidth \ TilePixelWidth)
        tY2 = tY + (EngineHeight \ TilePixelHeight)
        
        If tY2 - TileBufferSize < 1 Then
            tY2 = 1 + TileBufferSize
        End If
        
        If tX2 - TileBufferSize < 1 Then
            tX2 = 1 + TileBufferSize
        End If
        
        
        If tX2 + TileBufferSize > MaxTilesX Then
            tX2 = MaxTilesX - TileBufferSize
        End If
        
        If tY2 + TileBufferSize > MaxTilesY Then
            tY2 = MaxTilesY - TileBufferSize
        End If
        
        ' Controla el mouseTilePos
        
        MouseTilesPos.x = (UserPos.x * 64 + Mouse.x) \ TilePixelWidth
        MouseTilesPos.y = (UserPos.y * 64 + Mouse.y) \ TilePixelHeight
        
        If MouseTilesPos.x < 1 Then
            MouseTilesPos.x = 1
        End If
        
        If MouseTilesPos.y < 1 Then
            MouseTilesPos.y = 1
        End If
        
        If MouseTilesPos.x > MaxTilesX Then
            MouseTilesPos.x = MaxTilesX
        End If
        
        If MouseTilesPos.y > MaxTilesY Then
            MouseTilesPos.y = MaxTilesY
        End If
        
        MousePosOnMap.x = mapPreCalcPos(MouseTilesPos.x, MouseTilesPos.y).x
        MousePosOnMap.y = mapPreCalcPos(MouseTilesPos.x, MouseTilesPos.y).y
        
    Static LastCount As Long
    
    If GetTickCount - LastCount > 47 Then
    
        For lY = tY - TileBufferSize To tY2 + TileBufferSize
            For lX = tX - TileBufferSize To tX2 + TileBufferSize
    
                tempX = mapPreCalcPos(lX, lY).x - UserPos.x * TilePixelWidth + offX
                tempY = mapPreCalcPos(lX, lY).y - UserPos.y * TilePixelHeight + offY
    
                mapData(lX, lY).LightColor(0) = CalcVertexLight(3, Mouse.x, Mouse.y, tempX + 64, tempY, Color(1), Color(0))
                mapData(lX, lY).LightColor(1) = CalcVertexLight(3, Mouse.x, Mouse.y, tempX + 128, tempY + 32, Color(1), Color(0))
                mapData(lX, lY).LightColor(2) = CalcVertexLight(3, Mouse.x, Mouse.y, tempX, tempY + 32, Color(1), Color(0))
                mapData(lX, lY).LightColor(3) = CalcVertexLight(3, Mouse.x, Mouse.y, tempX + 64, tempY + 64, Color(1), Color(0))
    
                'vertex(0) = setVertex(cX + .Width, cY, 0, 1, Color(0), 0, 0, 0)
                'vertex(1) = setVertex(cX + (.Width * 2), cY + (.Height * 0.5), 0, 1, Color(1), 0, 1, 0)
                'vertex(2) = setVertex(cX, cY + (.Height * 0.5), 0, 1, Color(2), 0, 0, 1)
                'vertex(3) = setVertex(cX + .Width, cY + .Height, 0, 1, Color(3), 0, 1, 1)
                        
                '         1
                ' 2               4
                '         3
    
            Next lX
        Next lY
        
        LastCount = GetTickCount
        
    End If
        
        For lY = tY To tY2 + 2
            For lX = tX - 1 To tX2 + TileBufferSize
                
                tempX = mapPreCalcPos(lX, lY).x - UserPos.x * TilePixelWidth + offX
                tempY = mapPreCalcPos(lX, lY).y - UserPos.y * TilePixelHeight + offY
                
                'Layer 1 **********************************
                deviceRenderTexture mapData(lX, lY).Layer(1).GrhIndex, _
                                    tempX, tempY, _
                                    mapData(lX, lY).LightColor(), _
                                    IsometricType.IsometricBase
                '******************************************
                
                'Layer 2 **********************************
                If mapData(lX, lY).Layer(2).GrhIndex > 0 Then
                    deviceRenderTexture mapData(lX, lY).Layer(2).GrhIndex, _
                                    tempX, tempY, _
                                    mapData(lX, lY).LightColor(), _
                                    IsometricType.IsometricBase
                End If
                '******************************************
                
            Next lX
        Next lY
        
        
        For lY = tY - TileBufferSize To tY2 + TileBufferSize
            For lX = tX - TileBufferSize To tX2 + TileBufferSize
        
                tempX = mapPreCalcPos(lX, lY).x - UserPos.x * TilePixelWidth + offX
                tempY = mapPreCalcPos(lX, lY).y - UserPos.y * TilePixelHeight + offY
        
                'Layer 3 **********************************
                If mapData(lX, lY).Layer(3).GrhIndex > 0 Then
                    deviceRenderTexture mapData(lX, lY).Layer(3).GrhIndex, _
                                    tempX, tempY, _
                                    mapData(lX, lY).LightColor(), _
                                    IsometricType.Normal
                End If
                '******************************************
        
            Next lX
        Next lY
        
        For lY = tY - TileBufferSize To tY2 + TileBufferSize
            For lX = tX - TileBufferSize To tX2 + TileBufferSize
        
                tempX = mapPreCalcPos(lX, lY).x - UserPos.x * TilePixelWidth + offX
                tempY = mapPreCalcPos(lX, lY).y - UserPos.y * TilePixelHeight + offY
        
                'ParticleLayer ****************************
                If mapData(lX, lY).particleIndex > 0 Then
                    UpdateParticleGroup mapData(lX, lY).particleIndex, tempX, tempY
                    RenderParticleGroup mapData(lX, lY).particleIndex
                End If
                '******************************************
        
            Next lX
        Next lY
       
       
    'Set DeviceStates
    deviceResetRenderStates
        
    'Render HUD
    deviceRenderTexture 10, 0, 484, BasicColor(), IsometricType.Normal
    
End Sub
Private Sub deviceRenderBox(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                            Optional ByVal v1 As Long = -1, Optional ByVal v2 As Long = -1, _
                            Optional ByVal v3 As Long = -1, Optional ByVal v4 As Long = -1)

    Dim myVertex(3) As D3DTLVERTEX

        myVertex(0) = setVertex(X1, Y1 + X2, 0, 1, v1, 0, 0, 0)
        myVertex(1) = setVertex(X1, Y1, 0, 1, v2, 0, 1, 0)
        myVertex(2) = setVertex(X1 + Y2, Y1 + X2, 0, 1, v3, 0, 0, 1)
        myVertex(3) = setVertex(X1 + Y2, Y1, 0, 1, v4, 0, 1, 1)

    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, myVertex(0), Len(myVertex(0))

End Sub
Private Sub RenderParticleGroup(grIndex As Integer)

    With D3DDevice
    
        ' Set the render states for using point sprites
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetRenderState D3DRS_POINTSIZE, ParticleGroup(grIndex).lngFloatSize
        .SetRenderState D3DRS_POINTSIZE_MIN, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_A, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_B, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_C, ParticleGroup(grIndex).lngFloat1
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
         
        ' Set our texture
        #If LoadingMetod = 0 Then
            .SetTexture 0, oGraphic(lKeys(Grh(ParticleGroup(grIndex).myTextureGrh).FileNum)).D3DTexture
        #Else
        
        Static tSurface As tempTexture
        
        Set tSurface.Surface = getSurface(Grh(ParticleGroup(grIndex).myTextureGrh).FileNum, tSurface.Width, tSurface.Height)
        
        .SetTexture 0, tSurface.Surface
        #End If
        
        ' And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleGroup(grIndex).ParticleCounts, _
            ParticleGroup(grIndex).vertsPoints(0), Len(ParticleGroup(grIndex).vertsPoints(0))
            
    End With
    
End Sub
Private Sub deviceRenderTexture(ByRef GrhIndex As Long, ByRef cx As Single, ByRef cy As Single, ByRef Color() As Long, ByRef Iso As IsometricType, Optional ByRef Angle As Single = 0)
        
    #If LoadingMetod = 0 Then
    
    If textureLoad(GrhIndex) = False Then Exit Sub
    
    If Not oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture Is Nothing Then
        GeometryBoxType Grh(GrhIndex), cx, cy, Vector, Color(), Iso, Angle
    End If
               
    D3DDevice.SetTexture 0, oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture
        
    #Else
    
    Static d3dSurface As tempTexture
    
    Set d3dSurface.Surface = getSurface(Grh(GrhIndex).FileNum, d3dSurface.Width, d3dSurface.Height)
    
    If Not d3dSurface.Surface Is Nothing Then
        GeometryBoxType Grh(GrhIndex), cx, cy, Vector, Color(), Iso, Angle
    End If
    
    D3DDevice.SetTexture 0, d3dSurface.Surface
    
    #End If
        
    '##RENDERING METHOD 1## - Medium Faster
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vector(0), Len(Vector(0))
                
    '##RENDERING METHOD 2## - Faster
    D3DDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
            indexList(0), D3DFMT_INDEX16, _
            Vector(0), Len(Vector(0))
                
End Sub
Private Function GeometryBoxType(ByRef Grh As structGrhData, ByRef cx As Single, ByRef cy As Single, vertex() As D3DTLVERTEX, ByRef Color() As Long, ByRef Iso As IsometricType, Optional ByRef Angle As Single = 0)

        Select Case Iso
            Case IsometricType.Normal
                
                    With Grh
                    
                        vertex(0) = setVertex(cx, cy + .Height, 0, 1, Color(0), 0, 0, 1)
                        vertex(1) = setVertex(cx, cy, 0, 1, Color(1), 0, 0, 0)
                        vertex(2) = setVertex(cx + .Width, cy + .Height, 0, 1, Color(2), 0, 1, 1)
                        vertex(3) = setVertex(cx + .Width, cy, 0, 1, Color(3), 0, 1, 0)
                    
                    End With
                
            Case IsometricType.NormalRotation

            Dim X_Center As Single
            Dim Y_Center As Single
            Dim Radio    As Single
            Dim L_Point  As Single
            Dim R_Point  As Single
            Dim temp     As Single
            
                With Grh
             
                 X_Center = .sX + (.Width - .sX - 1) / 2
                 Y_Center = .sY + (.Height - .sY - 1) / 2
            
                 Radio = Sqr((.Width - X_Center) ^ 2 + (.Height - Y_Center) ^ 2)
            
                 temp = (.Width - X_Center) / Radio
                 R_Point = Atn(temp / Sqr(-temp * temp + 1))
                 L_Point = Pi - R_Point
                
                    'vertex(0) = setVertex(cX + X_Center + Cos(-L_Point - Angle) * Radio, cY + Y_Center - Sin(-L_Point - Angle) * Radio, 0, 1, Color(0), 0, .sX / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                    'vertex(1) = setVertex(cX + X_Center + Cos(L_Point - Angle) * Radio, cY + Y_Center - Sin(L_Point - Angle) * Radio, 0, 1, Color(1), 0, .sX / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                    'vertex(2) = setVertex(cX + X_Center + Cos(-R_Point - Angle) * Radio, cY + Y_Center - Sin(-R_Point - Angle) * Radio, 0, 1, Color(2), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                    'vertex(3) = setVertex(cX + X_Center + Cos(R_Point - Angle) * Radio, cY + Y_Center - Sin(R_Point - Angle) * Radio, 0, 1, Color(3), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                
                End With
            Case IsometricType.IsometricBase
            
                    With Grh
                        
                        vertex(0) = setVertex(cx + .Width, cy, 0, 1, Color(0), 0, 0, 0)
                        vertex(1) = setVertex(cx + (.Width * 2), cy + (.Height * 0.5), 0, 1, Color(1), 0, 1, 0)
                        vertex(2) = setVertex(cx, cy + (.Height * 0.5), 0, 1, Color(2), 0, 0, 1)
                        vertex(3) = setVertex(cx + .Width, cy + .Height, 0, 1, Color(3), 0, 1, 1)
                        
                    End With
                    
            Case IsometricType.IsometricHeight
            
                    With Grh
                        
                        'vertex(0) = setVertex(cX, cY - (.Height * 0.5), 0, 1, Color(0), 0, .sX / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        'vertex(1) = setVertex(cX + .Width, cY, 0, 1, Color(1), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, .sY / oGraphic(lKeys(.FileNum)).Height)
                        'vertex(2) = setVertex(cX, cY + (.Height * 0.5), 0, 1, Color(2), 0, .sX / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        'vertex(3) = setVertex(cX + .Width, cY + .Height, 0, 1, Color(3), 0, (.sX + .Width + 1) / oGraphic(lKeys(.FileNum)).Width, (.sY + .Height + 1) / oGraphic(lKeys(.FileNum)).Height)
                        
                    End With
        End Select
        
End Function
Public Function setVertex(ByRef x As Single, ByRef y As Single, ByRef z As Single, ByRef rhw As Single, ByRef Color As Long, ByRef Specular As Long, ByRef tu As Single, ByRef tv As Single) As D3DTLVERTEX
    
    With setVertex
        .sX = x
        .sY = y
        .sz = z
        .rhw = rhw
        .Color = Color
        .Specular = Specular
        .tu = tu
        .tv = tv
    End With
    
End Function
Public Function FloatToDWord(flo As Single) As Long
'A helper function, converts from C++ Float to C++ DWord
    Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, flo
    D3DX.BufferGetData buf, 0, 4, 1, FloatToDWord
End Function

'-----------------------------------------------------------------------------
'DOC: D3DEnum_Cleanup
'DOC:   Used to release any reference to the callback object passed in
'DOC:   and deallocate memory
'DOC: Params:
'DOC:   none
'DOC: Remarks:
'DOC:   none
'-----------------------------------------------------------------------------
Public Sub D3DEnum_Cleanup()
    Set g_EnumCallback = Nothing
    ReDim g_Adapters(0)
End Sub

'-----------------------------------------------------------------------------
'DOC: D3DEnum_BuildAdapterList
'DOC:   Used to intialzed a list of valid adapters and display modes
'DOC:
'DOC: Params:
'DOC:   EnumCallback    - can be Nothing or an object that has implemented
'DOC:                     VerifyDevice(usageflags as long ,format as CONST_D3DFORMAT)
'DOC:                     ussgeflags can be
'DOC:                           D3DCREATE_SOFTWARE_VERTEXPROCESSING
'DOC:                           D3DCREATE_HARDWARE_VERTEXPROCESSING
'DOC: Remarks:
'DOC:   caps for the device are passed to VerifyDevice in the DevCaps global
'DOC:
'-----------------------------------------------------------------------------
Private Function D3DEnum_BuildAdapterList(EnumCallback As Object) As Boolean
    
    On Local Error GoTo errOut
    
    Dim lAdapter As Long
        
    ' empty the list
    Call D3DEnum_Cleanup
            
    ' create d3d and dx objects if not already created
    If dX Is Nothing Then Set dX = New DirectX8
    If D3D Is Nothing Then Set D3D = dX.Direct3DCreate
    If D3DX Is Nothing Then Set D3DX = New D3DX8
    
    ' save callback
    Set g_EnumCallback = EnumCallback
    
    ' Make space for new adapter
    g_lNumAdapters = D3D.GetAdapterCount
    ReDim g_Adapters(g_lNumAdapters)
    
    ' Loop through all the adapters on the system
    For lAdapter = 0 To g_lNumAdapters - 1
    
        ' build a list of valid backbuffer formats
        D3DEnum_BuildValidFormatList lAdapter, D3DDEVTYPE_HAL
        D3DEnum_BuildValidFormatList lAdapter, D3DDEVTYPE_REF
        
                
        ' build a list of valid display modes for those formats
        D3DEnum_BuildDisplayModeList lAdapter, D3DDEVTYPE_HAL
        D3DEnum_BuildDisplayModeList lAdapter, D3DDEVTYPE_REF
        
        ' get the adapter identifier
        D3D.GetAdapterIdentifier lAdapter, 0, g_Adapters(lAdapter).d3dai
        
    Next
    
    D3DEnum_BuildAdapterList = True
    Exit Function
    
errOut:
    Debug.Print "Failed D3DEnum_BuildAdapterList"
End Function

'-----------------------------------------------------------------------------
' D3DEnum_BuildValidFormatList
'-----------------------------------------------------------------------------
Private Sub D3DEnum_BuildValidFormatList(lAdapter As Long, DevType As CONST_D3DDEVTYPE)
                        
        Dim lMode As Long
        Dim lUsage As Long
        Dim NumModes As Long
        Dim DisplayMode As D3DDISPLAYMODE
        Dim bCanDoWindowed As Boolean
        Dim bCanDoFullscreen As Boolean
                
        
        With g_Adapters(lAdapter).DevTypeInfo(DevType)
        
            ' Reset the number of available formats to
            .lNumFormats = 0
        
            ' Get the number of display modes
            ' a display mode is a size and format (ie 640x480 X8R8G8B8 60hz)
            NumModes = D3D.GetAdapterModeCount(lAdapter)
            ReDim .FormatInfo(NumModes)
                                
            ' Loop through all the display modes
            For lMode = 0 To NumModes - 1
                    
                ' Get information about this adapter in all the modes it supports
                Call D3D.EnumAdapterModes(lAdapter, lMode, DisplayMode)
                                
                ' See if the format is already in our format list
                If -1 <> D3DEnum_FindInFormatList(lAdapter, DevType, DisplayMode.Format) Then GoTo Continue
                                    
                ' Check the compatiblity of the format
                
                lUsage = D3DEnum_CheckFormatCompatibility(lAdapter, DevType, DisplayMode.Format, bCanDoWindowed, bCanDoFullscreen)
                                                                            
                ' Usage will come back -1 if VerifyDevice reject format
                If -1 = lUsage Then GoTo Continue
                
                ' Add the valid format and ussage
                .FormatInfo(.lNumFormats).Format = DisplayMode.Format
                .FormatInfo(.lNumFormats).usage = lUsage
                .FormatInfo(.lNumFormats).bCanDoWindowed = bCanDoWindowed
                .FormatInfo(.lNumFormats).bCanDoFullscreen = bCanDoFullscreen
                .lNumFormats = .lNumFormats + 1

                                
Continue:
            Next
            
        End With

End Sub

'-----------------------------------------------------------------------------
' D3DEnum_BuildDisplayModeList
'-----------------------------------------------------------------------------
Private Sub D3DEnum_BuildDisplayModeList(lAdapter As Long, DevType As CONST_D3DDEVTYPE)
                        
        Dim lMode As Long
        Dim NumModes As Long
        Dim DisplayMode As D3DDISPLAYMODE

        With g_Adapters(lAdapter).DevTypeInfo(DevType)
        
            ' Reset the number of validated display modes to 0
            .lNumModes = 0
            
            ' Get the number of display modes
            ' Note this list includes refresh rates
            ' a display mode is a size and format (ie 640x480 X8R8G8B8 60hz)
            NumModes = D3D.GetAdapterModeCount(lAdapter)

            ' Allocate space for our mode list
            ReDim .Modes(NumModes)

            ' Save the format of the desktop for windowed operation
            Call D3D.GetAdapterDisplayMode(lAdapter, g_Adapters(lAdapter).DesktopMode)
                                
            ' Loop through all the display modes
            For lMode = 0 To NumModes - 1
                    
                ' Get information about this adapter in all the modes it supports
                Call D3D.EnumAdapterModes(lAdapter, lMode, DisplayMode)
                
                ' filter out low resolution modes
                If DisplayMode.Width < 640 Or DisplayMode.Height < 400 Then GoTo Continue
                
                ' filter out modes allready in the list
                ' that might differ only in refresh rate
                If -1 <> D3DEnum_FindInDisplayModeList(lAdapter, DevType, DisplayMode) Then GoTo Continue
                
                
                ' filter out modes with formats that arent confirmed to work
                ' see BuildFormatList and ConfirmFormatList
                If -1 = D3DEnum_FindInFormatList(lAdapter, DevType, DisplayMode.Format) Then GoTo Continue
                                                
                ' At this point, the modes format has been validated,
                ' is not a duplicate refresh rate, and not a low res mode
                ' Add the mode to the list of working modes for the adapter
                .Modes(.lNumModes).lHeight = DisplayMode.Height
                .Modes(.lNumModes).lWidth = DisplayMode.Width
                .Modes(.lNumModes).Format = DisplayMode.Format
                .lNumModes = .lNumModes + 1
                                            
Continue:
            Next
            
        End With

End Sub

'-----------------------------------------------------------------------------
' D3DEnum_FindInDisplayModeList
'-----------------------------------------------------------------------------
Public Function D3DEnum_FindInDisplayModeList(lAdapter As Long, DevType As CONST_D3DDEVTYPE, DisplayMode As D3DDISPLAYMODE) As Long
    
    Dim lMode As Long
    Dim NumModes As Long
    
    NumModes = g_Adapters(lAdapter).DevTypeInfo(DevType).lNumModes
    D3DEnum_FindInDisplayModeList = -1
    
    For lMode = 0 To NumModes - 1
      With g_Adapters(lAdapter).DevTypeInfo(DevType).Modes(lMode)
          If .lWidth = DisplayMode.Width And _
              .lHeight = DisplayMode.Height And _
              .Format = DisplayMode.Format Then
              D3DEnum_FindInDisplayModeList = lMode
              Exit Function
          End If
      End With
    Next
    
End Function


'-----------------------------------------------------------------------------
' D3DEnum_FindInFormatList
'-----------------------------------------------------------------------------
Public Function D3DEnum_FindInFormatList(lAdapter As Long, DevType As CONST_D3DDEVTYPE, Format As CONST_D3DFORMAT) As Long
    
    Dim lFormat As Long
    Dim NumFormats As Long
    
    NumFormats = g_Adapters(lAdapter).DevTypeInfo(DevType).lNumFormats
    D3DEnum_FindInFormatList = -1
    
    For lFormat = 0 To NumFormats - 1
      With g_Adapters(lAdapter).DevTypeInfo(DevType).FormatInfo(lFormat)
          If .Format = Format Then
             D3DEnum_FindInFormatList = .usage
             Exit Function
          End If
      End With
    Next
    
    D3DEnum_FindInFormatList = -1
    
End Function

'-----------------------------------------------------------------------------
' D3DEnum_CheckFormatCompatibility
'-----------------------------------------------------------------------------
Private Function D3DEnum_CheckFormatCompatibility(lAdapter As Long, DeviceType As CONST_D3DDEVTYPE, Format As CONST_D3DFORMAT, ByRef OutCanDoWindowed As Boolean, ByRef OutCanDoFullscreen As Boolean) As Long
        On Local Error GoTo errOut

        D3DEnum_CheckFormatCompatibility = -1
        
        'Dim d3dcaps As D3DCAPS8
        Dim flags As Long

        ' Filter out incompatible backbuffers
        ' Note: framework always has the backbuffer and the frontbuffer (screen) format matching
        OutCanDoWindowed = True: OutCanDoFullscreen = True
        If 0 <> D3D.CheckDeviceType(lAdapter, DeviceType, Format, Format, 0) Then OutCanDoWindowed = False
        If 0 <> D3D.CheckDeviceType(lAdapter, DeviceType, Format, Format, 1) Then OutCanDoFullscreen = False
        If (OutCanDoWindowed = False) And (OutCanDoFullscreen = False) Then Exit Function

        ' If no form was passed in to use as a callback
        ' then default to sofware vertex processing

        ' Get the device capablities
        D3D.GetDeviceCaps lAdapter, DeviceType, DevCaps
        g_Adapters(lAdapter).d3dcaps = DevCaps
        
        ' If user doesnt want to verify the device (didnt provide a callback)
        ' fall back to software
        D3DEnum_CheckFormatCompatibility = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        If g_EnumCallback Is Nothing Then Exit Function
        
        ' Confirm the device for HW vertex processing
        flags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = flags
        If DevCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then
           If g_EnumCallback.VerifyDevice(flags, Format) Then Exit Function
        End If
        
        ' Try Software VertexProcesing
        flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = flags
        If g_EnumCallback.VerifyDevice(flags, Format) Then Exit Function
                                
        ' Fail
        D3DEnum_CheckFormatCompatibility = -1
        
        Exit Function
errOut:

End Function

