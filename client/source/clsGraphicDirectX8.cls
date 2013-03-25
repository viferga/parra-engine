VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicDirectX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DOC: Type that hold info about a display mode
Private Type D3DUTIL_MODEINFO
    lWidth As Long                                  'Screen width in this mode
    lHeight As Long                                 'Screen height in this mode
    format As CONST_D3DFORMAT                       'Pixel format in this mode
    VertexBehavior As CONST_D3DCREATEFLAGS          'Whether this mode does SW or HW vertex processing
End Type

' DOC: Type that hold info about a particular back buffer format
Private Type D3DUTIL_FORMATINFO
    format As CONST_D3DFORMAT
    usage As Long
    bCanDoWindowed As Boolean
    bCanDoFullscreen As Boolean
End Type

Private Type D3DUTIL_DEVTYPEINFO
    'List of compatible Modes for this device
    lNumModes As Long
    Modes() As D3DUTIL_MODEINFO
        
    'List of compatible formats for this device
    lNumFormats As Long
    FormatInfo() As D3DUTIL_FORMATINFO

    lCurrentMode As Long        'index into modes array of the current mode
    
End Type

Private Type D3DUTIL_ADAPTERINFO
    'Device data
    DeviceType As CONST_D3DDEVTYPE                  'Reference, HAL
    d3dcaps As D3DCAPS8                             'Caps of this device
    sDesc As String                                 'Name of this device
    lCanDoWindowed As Long                          'Whether or not this device can work windowed
                
    DevTypeInfo(2) As D3DUTIL_DEVTYPEINFO           'format and display mode list for hal=1 for ref=2
    
    d3dai As D3DADAPTER_IDENTIFIER8
            
    DesktopMode As D3DDISPLAYMODE
    
    'CurrentState
    bWindowed As Boolean        'currently in windowed mode
    bReference As Boolean       'currently using reference rasterizer

    
    
End Type

Private g_Adapters() As D3DUTIL_ADAPTERINFO      ' Array of Adapter infos
Private g_lCurrentAdapter As Long                ' current adapter (index into infos)
Private g_lNumAdapters As Long                   ' size of the g_Adapters array
Private g_EnumCallback As Object

' DirectX8 & Extras
Private D3DWindow As DxVBLibA.D3DPRESENT_PARAMETERS
Private DispMode As DxVBLibA.D3DDISPLAYMODE
Private DevCaps As DxVBLibA.D3DCAPS8

Private dX As DxVBLibA.DirectX8 'Root object
Private D3D As DxVBLibA.Direct3D8 ' Direct3D interface
Private D3DX As DxVBLibA.D3DX8 ' Helper library
Private D3DDevice As DxVBLibA.Direct3DDevice8 'Represents the hardware doing the rendering

Private D3DDeviceType As CONST_D3DDEVTYPE
                                                ' Current state (use as read only)
Private g_behaviorflags As Long                  ' Current VertexProcessing (hardware or software)
Private g_focushwnd As Long                      ' Current focus window handle
  
Private g_lWindowWidth As Long                   ' backbuffer width of windowed state
Private g_lWindowHeight As Long                  ' backbuffer  height of windowed state
Private g_WindowRect As Rect                     ' size of window (including title bar)

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef Rect As Rect) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef Rect As Rect) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hwndafter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal options As Long) As Long
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal newVal As Long) As Long
Private Declare Function GetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Implements clsGraphicalDevice

Private Function clsGraphicalDevice_Initialize(ByRef Top As Integer, ByRef Left As Integer, ByRef Width As Integer, ByRef Height As Integer, Frm As Form, Optional ByRef BitsPerPixel As Byte = 32, Optional ByRef Windowed As Boolean = True) As Boolean
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
        SetWindowPos frmConnect.hWnd, 0, 0, 0, 800, 600, 0
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
        clsGraphicalDevice_Initialize = engineInitializingInWindow(Frm.hWnd, 0, DevType, True)
    Else
        clsGraphicalDevice_Initialize = engineInitializingInFullscreen(Frm.hWnd, 0, 0, DevType, True, BitsPerPixel)
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
    ScrollPixelsPerFrame.X = 8
    ScrollPixelsPerFrame.Y = 8
    
    'Create a pixel shader
    cPixelShader = pixelShaderMakeFromMemory(psOriginalColor)
    If cPixelShader > 0 Then D3DDevice.SetPixelShader cPixelShader
    
    'Initialize gui
    If guiInitialize = False Then GoTo errHandle
    
    clsGraphicalDevice_Initialize = True
    Exit Function
    
errHandle:

    If Err.Number = 429 Then
      MsgBox "No se puede iniciar el motor grafico, ya que no ha sido detectado DirectX 8. Reinstalalo", vbCritical
    Else
        MsgBox "Error al iniciar el motor grafico" & vbNewLine _
        & "ErrNumber: " & Err.Number & vbNewLine _
        & "ErrDescription: " & Err.Description, vbOKOnly, "Error"
    End If

    clsGraphicalDevice_Initialize = False

End Function

Private Function engineInitializingInWindow(hWnd As Long, AdapterIndex As Long, DevType As CONST_D3DDEVTYPE, bTryFallbacks As Boolean) As Boolean
   
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
        .hDeviceWindow = hWnd
    End With
        
    g_focushwnd = hWnd
    
    Dim rc As Rect
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
        g_behaviorflags = D3DEnum_FindInFormatList(g_lCurrentAdapter, DevType, d3ddm.format)
        If g_behaviorflags <= 0 Then g_behaviorflags = D3DEnum_CheckFormatCompatibility(AdapterIndex, DevType, d3ddm.format, False, False)

        
        
        D3DWindow.BackBufferFormat = d3ddm.format
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
Private Function engineInitializingInFullscreen(hWnd As Long, AdapterIndex As Long, modeIndex As Long, DevType As CONST_D3DDEVTYPE, bTryFallbacks As Boolean, BitsPerPixel As Byte) As Boolean
    
    On Error GoTo errOut
    
    Dim ModeInfo As D3DUTIL_MODEINFO

    Dim rc As Rect
    
    'save the current adapter
    g_lCurrentAdapter = AdapterIndex
        
    
    g_focushwnd = hWnd
    
    
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
            ModeInfo.format = D3DFMT_A8R8G8B8
        Case 24
            ModeInfo.format = D3DFMT_R8G8B8
        Case 16
            ModeInfo.format = D3DFMT_R5G6B5
    End Select

    
    With g_Adapters(g_lCurrentAdapter)
            
        With .DevTypeInfo(DevType)
            ModeInfo = .Modes(modeIndex)
            g_behaviorflags = .Modes(modeIndex).VertexBehavior
            
            D3DWindow.BackBufferWidth = ModeInfo.lWidth
            D3DWindow.BackBufferHeight = ModeInfo.lHeight
            D3DWindow.BackBufferFormat = ModeInfo.format
            D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
            D3DWindow.Windowed = 0
            D3DWindow.FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            
            .lCurrentMode = modeIndex
        End With
             

        .bWindowed = False
        .DeviceType = DevType
        
        If g_behaviorflags <= 0 Then g_behaviorflags = D3DEnum_CheckFormatCompatibility(AdapterIndex, DevType, ModeInfo.format, False, False)
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
                If -1 <> D3DEnum_FindInFormatList(lAdapter, DevType, DisplayMode.format) Then GoTo Continue
                                    
                ' Check the compatiblity of the format
                
                lUsage = D3DEnum_CheckFormatCompatibility(lAdapter, DevType, DisplayMode.format, bCanDoWindowed, bCanDoFullscreen)
                                                                            
                ' Usage will come back -1 if VerifyDevice reject format
                If -1 = lUsage Then GoTo Continue
                
                ' Add the valid format and ussage
                .FormatInfo(.lNumFormats).format = DisplayMode.format
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
                If -1 = D3DEnum_FindInFormatList(lAdapter, DevType, DisplayMode.format) Then GoTo Continue
                                                
                ' At this point, the modes format has been validated,
                ' is not a duplicate refresh rate, and not a low res mode
                ' Add the mode to the list of working modes for the adapter
                .Modes(.lNumModes).lHeight = DisplayMode.Height
                .Modes(.lNumModes).lWidth = DisplayMode.Width
                .Modes(.lNumModes).format = DisplayMode.format
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
              .format = DisplayMode.format Then
              D3DEnum_FindInDisplayModeList = lMode
              Exit Function
          End If
      End With
    Next
    
End Function


'-----------------------------------------------------------------------------
' D3DEnum_FindInFormatList
'-----------------------------------------------------------------------------
Public Function D3DEnum_FindInFormatList(lAdapter As Long, DevType As CONST_D3DDEVTYPE, format As CONST_D3DFORMAT) As Long
    
    Dim lFormat As Long
    Dim NumFormats As Long
    
    NumFormats = g_Adapters(lAdapter).DevTypeInfo(DevType).lNumFormats
    D3DEnum_FindInFormatList = -1
    
    For lFormat = 0 To NumFormats - 1
      With g_Adapters(lAdapter).DevTypeInfo(DevType).FormatInfo(lFormat)
          If .format = format Then
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
Private Function D3DEnum_CheckFormatCompatibility(lAdapter As Long, DeviceType As CONST_D3DDEVTYPE, format As CONST_D3DFORMAT, ByRef OutCanDoWindowed As Boolean, ByRef OutCanDoFullscreen As Boolean) As Long
        On Local Error GoTo errOut

        D3DEnum_CheckFormatCompatibility = -1
        
        'Dim d3dcaps As D3DCAPS8
        Dim Flags As Long

        ' Filter out incompatible backbuffers
        ' Note: framework always has the backbuffer and the frontbuffer (screen) format matching
        OutCanDoWindowed = True: OutCanDoFullscreen = True
        If 0 <> D3D.CheckDeviceType(lAdapter, DeviceType, format, format, 0) Then OutCanDoWindowed = False
        If 0 <> D3D.CheckDeviceType(lAdapter, DeviceType, format, format, 1) Then OutCanDoFullscreen = False
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
        Flags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = Flags
        If DevCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then
           If g_EnumCallback.VerifyDevice(Flags, format) Then Exit Function
        End If
        
        ' Try Software VertexProcesing
        Flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = Flags
        If g_EnumCallback.VerifyDevice(Flags, format) Then Exit Function
                                
        ' Fail
        D3DEnum_CheckFormatCompatibility = -1
        
        Exit Function
errOut:

End Function

Private Sub clsGraphicalDevice_DeInitialize()

    Dim emptycaps As D3DCAPS8
    Dim emptyrect As Rect
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

Private Function clsGraphicalDevice_DeviceIsContextValid() As DEVICE_NO_ERROR
    
    clsGraphicalDevice_DeviceIsContextValid = DEVICE_CTX_VALID

    With D3DDevice

        If .TestCooperativeLevel = D3D_OK Then Exit Function
        
                Dim H As Long
                clsGraphicalDevice_DeviceIsContextValid = DEVICE_CTX_ERROR

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
                            
                            clsGraphicalDevice_DeviceIsContextValid = DEVICE_CTX_LOST
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
                            Loop While (.TestCooperativeLevel = D3DERR_DEVICELOST) 'no seria DEVICENORESET en vez de lost?
                            
                            clsGraphicalDevice_DeviceIsContextValid = DEVICE_CTX_RESET
                            Exit Function
                End Select
            
    End With
End Function

Private Sub clsGraphicalDevice_BeginScene(ClearRect As Rect, Flags As CLEAR_FLAGS)
Dim CustomFlags As Long
CustomFlags = 0

    If (Flags And CLEAR_COLOR) Then
          CustomFlags = CustomFlags Or D3DCLEAR_TARGET
    ElseIf (Flags And CLEAR_STENCIL) Then
          CustomFlags = CustomFlags Or D3DCLEAR_STENCIL
    ElseIf (Flags And CLEAR_DEPTH) Then
          CustomFlags = CustomFlags Or D3DCLEAR_ZBUFFER
    End If

    With D3DDevice
                    
        .Clear 1, D3DRECT, CustomFlags, 0, 1#, 0
        .BeginScene
    
    End With
End Sub

Private Sub clsGraphicalDevice_EndScene(DestRect As Rect, hWnd As Long)

    With D3DDevice
                    
        'Flip the backbuffer to the screen
        .Present DestRect, DestRect, hWnd, ByVal 0&
        .EndScene
    
    End With
    
End Sub
