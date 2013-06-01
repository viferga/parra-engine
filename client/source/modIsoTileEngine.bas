Attribute VB_Name = "modIsoTileEngine"
Option Explicit

Public WireFrame As Boolean

Public Color(1) As D3DCOLORVALUE

Public Const EngineWidth As Integer = 800
Public Const EngineHeight As Integer = 600

Private Const TileBufferSize As Integer = 2

Public Enum IsometricType
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
        
    name As String
    
    Moving As Byte
    scrollDirection As structPositionInt
    MoveOffset As structPositionSng
End Type

Public Enum characterType
    player = 0
    Npc = 1
End Enum

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

' Directx8 Fonts
Public Type FontInfo
    MainFont As DxVBLibA.D3DXFont
    MainFontDesc As IFont
    MainFontFormat As New StdFont
    Color As Long
End Type: Public Font() As FontInfo

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
    Dim I As Long
    
      ReDim Preserve Grh(1 To GetVar(App.Path & "\Init\grh.ini", "INIT", "numGrh")) As structGrhData
    
        For I = 1 To UBound(Grh)
        
            With Grh(I)
                .FileNum = ReadField(1, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .sX = ReadField(2, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .sY = ReadField(3, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .Width = ReadField(4, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .Height = ReadField(5, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .offsetX = ReadField(6, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                .offsetY = ReadField(7, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                                
                .NumFrames = ReadField(8, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                
                ReDim .Frames(1 To .NumFrames)
                
                If (.NumFrames < 1) Then
                    
                    Dim frameCount As Long
                    
                    For frameCount = 1 To .NumFrames
                         .Frames(frameCount) = ReadField(frameCount + 8, _
                                                GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                    Next frameCount

                Else

                    .Frames(1) = I
                End If
                
                .Speed = ReadField(.NumFrames + 8, GetVar(App.Path & "\Init\grh.ini", "GRH", "grh" & CStr(I)), Asc("-"))
                
            End With
            
            #If WorldEditor = 1 Then
                'Add GrhList
                frmMain.grhList.AddItem "Grh" & CStr(I)
            #End If
        Next I
    
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

    'If (testCooperative = False) Then Exit Sub
   If Not (GraphicalDevice.DeviceIsContextValid = DEVICE_CTX_VALID) Then Exit Sub
   ' With D3DDevice
        
   '     If MotionBlur = True And errMotion = False Then
   '         .SetRenderTarget m_pDisplayTextureSurface, m_pDisplayZSurface, 0
   '         .Clear 1, RenderRect, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
   '     Else
   '         .Clear 1, RenderRect, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
   '     End If
        
   '         .BeginScene
    GraphicalDevice.BeginScene RenderRect, CLEAR_FLAGS.CLEAR_COLOR
    
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
                            GraphicalDevice.renderTexture frmMain.grhList.ListIndex + 1, Mouse.x, Mouse.y, BasicColor(), frmMain.cmbMode.ListIndex
                        End If
                    End If
                #End If
                
                'Render Gui
                If RenderGUI = True Then
                    GraphicalDevice.guiRender (Instance)
                End If
                
                If MotionBlur = True And errMotion = False Then GraphicalDevice.resetMotionStates
            
            '.EndScene
        '.Present RenderRect, ByVal 0&, 0, ByVal 0&
    GraphicalDevice.EndScene RenderRect, HwndVal& ' frmMain.hwnd
    
    'End With
    
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = gameGetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * 0.018   ' Engine Speed
    
End Sub

Public Sub fontDeInitializing()
    Dim I As Byte
    
    For I = 1 To UBound(Font)
        Set Font(I).MainFont = Nothing
        Set Font(I).MainFontDesc = Nothing
        Set Font(I).MainFontFormat = Nothing
    Next I
    
End Sub
Private Sub fontRender(ByRef Text As String, ByRef index As Byte, _
                            ByRef x As Integer, ByRef y As Integer, _
                            ByRef Width As Integer, ByRef Height As Integer, _
                            format As Long)
                            

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
        WriteCharEvents 1, playerCharIndex, player
        'faltaria lo de npc
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
        
        mapData(x + addX, y + addY).charindex = characterIndex
        .Pos.x = x + addX
        .Pos.y = y + addY
        mapData(x, y).charindex = 0
        
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
                GraphicalDevice.renderTexture mapData(lX, lY).Layer(1).GrhIndex, _
                                    tempX, tempY, _
                                    mapData(lX, lY).LightColor(), _
                                    IsometricType.IsometricBase
                '******************************************
                
                'Layer 2 **********************************
                If mapData(lX, lY).Layer(2).GrhIndex > 0 Then
                    GraphicalDevice.renderTexture mapData(lX, lY).Layer(2).GrhIndex, _
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
                    GraphicalDevice.renderTexture mapData(lX, lY).Layer(3).GrhIndex, _
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
                    GraphicalDevice.renderParticleGroup mapData(lX, lY).particleIndex
                End If
                '******************************************
        
            Next lX
        Next lY
       
       
    'Set DeviceStates
    GraphicalDevice.resetRenderStates 'GDK: Necesario¿?
        
    'Render HUD
    GraphicalDevice.renderTexture 10, 0, 484, BasicColor(), IsometricType.Normal
    
End Sub

Public Function GeometryBoxType(ByRef Grh As structGrhData, ByRef cx As Single, ByRef cy As Single, vertex() As D3DTLVERTEX, ByRef Color() As Long, ByRef Iso As IsometricType, Optional ByRef Angle As Single = 0)

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


