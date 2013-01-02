Attribute VB_Name = "modGeneral"
Option Explicit

Public Mouse As structPositionSng

Public generalIP As String
Public generalPort As Integer

Public bRunning As Boolean 'Switch on/off the game
Public gamePaused As Boolean 'Pauses/continuous the game

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    x As Long
    y As Long
End Type

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Sub main()

    IntializeRandom

    'Show the connect form
    frmConnect.Show

    'Show the main form
    frmMain.Show
    frmMain.Visible = False
    
    loadConnectionInfo generalIP, generalPort
    
    MotionBlur = False
    
    gamePaused = True
    
    'bRunning is true if TileEninge initialize correctly
    bRunning = engineInitializing(0, 0, 800, 600, frmMain, 16, True)   '553
    
    'Load MapData
    mapLoadAll
    
    gamePaused = False
    
    'Start Socket
    Set frmMain.Winsock = New clsSocket

    'Starts the game
    gameLoop
    
    'Close winsock
    frmMain.Winsock.CloseSck
    
    'Start Socket
    Set frmMain.Winsock = Nothing
    
    ' UnloadMapdata
    mapUnloadAll
    
    'Deinit engine
    engineDeinitializing
    
    'Stop Audio
    Sound_Destroy
    
    Unload frmMain
    
    End
End Sub

Public Sub gameLoop()

    Dim lTimerCount As Long

    Do While (bRunning = True)
    
        If gamePaused = False Then
        
            If (frmMain.Visible = True And frmMain.WindowState <> vbMinimized) Then
              
                If (gamePaused = False) Then
                    If (GetActiveWindow() = frmMain.hwnd) Then
                        gameCheckKeys
                    End If
                End If
               
                'Run Engine
                showNextFrame
            
                If (GetTickCount - lTimerCount >= 1000) Then
                    
                    FramesPerSec = FramesPerSecCounter
                    FramesPerSecCounter = 1
                    lTimerCount = GetTickCount
                End If
                
            Else
                Sleep 10&
            End If
        
        End If
        
        'Outgoing Protocol Sending
        SendBuffer

        DoEvents
    Loop
    
End Sub
Private Sub loadConnectionInfo(ByRef ip As String, ByRef port As Integer)

    ip = GetVar(App.Path & "\cliente.ini", "MAIN", "ip")
    port = CInt(GetVar(App.Path & "\cliente.ini", "MAIN", "port"))

End Sub
Private Sub gameCheckKeys()

    Static lastMovement(1) As Long
  
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    
#If ParticleEditor = 1 Then
    If EditParticle = False Then
#End If

    If GetTickCount - lastMovement(0) > 35 Then
            gameCheckMovmentKeys
        lastMovement(0) = GetTickCount
    End If
    
#If ParticleEditor = 1 Then
    End If
#End If
    
    If GetTickCount - lastMovement(1) > 65 Then
        gameCheckRoutineKeys
        lastMovement(1) = GetTickCount
    End If

End Sub
Private Sub gameCheckMovmentKeys()
                            
    If UserMoving = 0 Then
        
        If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyUp) < 0 Then
            Move NorthEast
            Exit Sub
        End If
                
        If GetKeyState(vbKeyRight) < 0 And GetKeyState(vbKeyDown) < 0 Then
            Move SouthEast
            Exit Sub
        End If
                
        If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyUp) < 0 Then
            Move NorthWest
            Exit Sub
        End If
                
        If GetKeyState(vbKeyLeft) < 0 And GetKeyState(vbKeyDown) < 0 Then
            Move SouthWest
            Exit Sub
        End If
        
        If GetKeyState(vbKeyUp) < 0 Then
            Move North
            Exit Sub
        End If
        
        If GetKeyState(vbKeyDown) < 0 Then
            Move South
            Exit Sub
        End If
        
        If GetKeyState(vbKeyRight) < 0 Then
            Move East
            Exit Sub
        End If
        
        If GetKeyState(vbKeyLeft) < 0 Then
            Move West
            Exit Sub
        End If
        
    End If

End Sub
Private Sub gameCheckRoutineKeys()
            
        'Atack!
        If GetKeyState(vbKeyControl) < 0 Then
                
            Exit Sub
        End If
            
        'Trow An Item
        If GetKeyState(vbKeyT) < 0 Then
                
            Exit Sub
        End If
        
        If GetKeyState(vbKeyF12) < 0 Then
            playerStatus = playerState.plyExit
            WriteOutgoingData ClientPacketID.UserEvents
            Exit Sub
        End If
        
        If GetKeyState(vbKey0) < 0 Then
            MotionBlur = Not MotionBlur
            Exit Sub
        End If
        
        If GetKeyState(vbKey1) < 0 Then
            If Not MotionBlur Or lBlurFactor = 1 Then Exit Sub
            lBlurFactor = lBlurFactor - 1
            Exit Sub
        End If
        
        If GetKeyState(vbKey2) < 0 Then
            If Not MotionBlur Or lBlurFactor = 255 Then Exit Sub
            lBlurFactor = lBlurFactor + 1
            Exit Sub
        End If
        
        If GetKeyState(vbKeyE) < 0 Then
            #If WorldEditor = 1 Then
                ' ** Edicion de Mapas **
                EditMap = Not EditMap
                frmMain.picEditor.Visible = Not frmMain.picEditor.Visible
                frmMain.grhList.ListIndex = -1
            #End If
            Exit Sub
        End If
        
        If GetKeyState(vbKeyP) < 0 Then
            #If ParticleEditor = 1 Then
                ' ** Edicion de Particulas **
                EditParticle = Not EditParticle
            #End If
            Exit Sub
        End If
        
        If GetKeyState(vbKeyV) < 0 Then
            'guiReLocateBox boxindex, mouse.x, mouse.y
            Exit Sub
        End If
        
        If GetKeyState(vbKeyG) < 0 Then
            RenderGUI = Not RenderGUI
            Exit Sub
        End If
        
        If GetKeyState(vbKeyS) < 0 Then
            ShowShader = Not ShowShader
            Exit Sub
        End If
        
        If GetKeyState(vbKeyW) < 0 Then
            WireFrame = Not WireFrame
            Exit Sub
        End If
        
End Sub

Public Function gameGetElapsedTime() As Single
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    QueryPerformanceCounter start_time
    gameGetElapsedTime = (start_time - end_time) / timer_freq * 1000
    QueryPerformanceCounter end_time
End Function
