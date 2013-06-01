Attribute VB_Name = "modParticle"
Option Explicit

Public EditParticle As Boolean

' ONLY PARTICLE
 
Private Type structParticle

    sngA            As Single
    sngR            As Single
    sngG            As Single
    sngB            As Single
    sngAlphaDecay   As Single
    
    sngSize         As Single
    
    sngX            As Single
    sngY            As Single
    
    sngXAccel       As Single
    sngYAccel       As Single
    
    sngXSpeed       As Single
    sngYSpeed       As Single
    
End Type

' ONLY GROUP

Private Type structGroupParticle

    ParticleCounts  As Long
    Particles()     As structParticle
    vertsPoints()   As D3DTLVERTEX
    
    ' POSITION GROUP & INFO
    myTextureGrh        As Long ' Contiene el indice del grh del grafico
    
    sngX                As Single
    sngY                As Single
    sngProgression      As Single
 
    lngFloat0           As Long
    lngFloat1           As Long
    lngFloatSize        As Long
    
    lngPreviousFrame    As Long
    
End Type: Public ParticleGroup() As structGroupParticle

Public Enum meteoState
    Lluvia = 0
    Niebla = 1
    Nublado = 2
    Viento = 3
    Tormenta = 4
    Normal = 5
End Enum
Public Sub meteoChangeStatus(ByVal fxStat As meteoState)
    
    Select Case fxStat
        Case meteoState.Lluvia
        'Particlegrupcreate(lluvia)
        Exit Sub
        
        Case meteoState.Niebla
        
        Exit Sub
        
        Case meteoState.Nublado
        
        Exit Sub
        
        Case meteoState.Viento
        
        Exit Sub
        
        Case meteoState.Tormenta
        
        Exit Sub
        
        Case meteoState.Normal
        
        Exit Sub
        
        Case Else: 'Particlegrupcreate(normal)
        
    End Select
    
End Sub
Private Function GetParticleInfo(index As Integer, ByVal Data As String) As Single

    GetParticleInfo = CSng(GetVar(App.Path & "\Init\particle.ini", CStr(index), Data))

End Function
Public Sub loadParticleGroup()
    Dim GroupCount As Integer
    Dim I As Integer
    
    GroupCount = CInt(GetVar(App.Path & "\Init\particle.ini", "MAIN", "GroupCount"))
        
    ReDim Preserve ParticleGroup(1 To GroupCount) As structGroupParticle
    
    For I = 1 To GroupCount

        ParticleGroup(I).ParticleCounts = GetParticleInfo(CStr(I), "pCount")
        ParticleGroup(I).myTextureGrh = GetParticleInfo(CStr(I), "TextureGrh")
        
        '#If LoadingMetod = 0 Then
        '    If textureLoad(ParticleGroup(I).myTextureGrh) = False Then MsgBox "Error: Cargando las texturas de las particulas", vbCritical
        '#End If
        
        Begin I
        
    Next I

End Sub

Private Sub Begin(grIndex As Integer)
    ' We initialize our stuff here
    Dim I As Long
     
    With ParticleGroup(grIndex)
    
        
        .lngFloat0 = GraphicalDevice.FloatToDWord(GetParticleInfo(CStr(grIndex), "FloatA"))
        .lngFloat1 = GraphicalDevice.FloatToDWord(GetParticleInfo(CStr(grIndex), "FloatB"))
        .lngFloatSize = GraphicalDevice.FloatToDWord(GetParticleInfo(CStr(grIndex), "Size")) ' Size of our flame particles..
         
        ' Redim our particles to the particlecount
        ReDim .Particles(0 To .ParticleCounts)
         
        ' Redim vertices to the particle count
        ' Point sprites, so 1 per particle
        ReDim .vertsPoints(0 To .ParticleCounts)
        
        ' Now generate all particles
        For I = 0 To .ParticleCounts
            Reset grIndex, I
        Next I

        ' Set initial time
        .lngPreviousFrame = GetTickCount()
    
    End With
    
End Sub

Private Sub Reset(grIndex As Integer, I As Long) ' Reset GROUP
Dim X As Single, Y As Single, Radio As Single
Dim Progression As Integer, Direction As Single
    
With ParticleGroup(grIndex)

    Select Case grIndex
        
            Case 1 ' Blue Ball
                
                X = .sngX: Y = .sngY
                ResetIt grIndex, I, X, Y, -20, (-1 * Rnd), 0.01, Rnd, 16
                ResetColor grIndex, I, 0.25, 0.25, 1, 1, 0.1 + (0.1 * Rnd)
            
            Case 2 ' Fire
                
                X = .sngX + (Rnd * 10)
                Y = .sngY
                 
                ' This is were we will reset individual particles.
                ResetIt grIndex, I, X, Y, -0.4 + (Rnd * 0.8), -0.5 - (Rnd * 0.4), 0, -(Rnd * 0.3), 2
                ResetColor grIndex, I, 1, 0.5, 0.2, 0.6 + (0.2 * Rnd), 0.01 + Rnd * 0.05
            
            Case 3 ' Smoke
                
                X = .sngX + 1 * Rnd + .sngX
                Y = .sngY * Rnd + .sngY
                ResetIt grIndex, I, X, Y, -(Rnd / 3 + 0.1), ((Rnd / 2) - 0.7) * 3, (Rnd - 0.5) / 200, (Rnd - 0.5) / 200, 20
                ResetColor grIndex, I, 0.8, 0.8, 0.8, 0.3, (Rnd * 0.005) + 0.005
            
            Case 4 ' Snow
            
                X = .sngX * Rnd
                Y = .sngY * Rnd

                ResetIt grIndex, I, X, Y, Rnd - 0.5, (Rnd + 0.3) * 4, 0, 0, ((Rnd + 0.3) * 4) * 3
                ResetColor grIndex, I, 1, 1, 1, 0.5, 0.02 * Rnd
            
            Case 5 ' MagicFire
            
                X = .sngX + (Rnd * 2): Y = .sngY
                ResetIt grIndex, I, X, Y, -0.4 + (Rnd * 0.8), -0.5 - (Rnd * 0.4), 0, -(Rnd * 0.3), 32
                ResetColor grIndex, I, 1, 0.5, 0.1, 0.7 + (0.2 * Rnd), 0.01 + Rnd * 0.05
            
            Case 6 ' LevelUp
            
                X = .sngX: Y = .sngY
                ResetIt grIndex, I, X, Y, Rnd * 1.5 - 0.75, Rnd * 1.5 - 0.75, Rnd * 4 - 2, Rnd * -4 + 2, 16
                ResetColor grIndex, I, 1, 0.5, 0.1, 1, 0.07 + Rnd * 0.01
            
            Case 7 ' LevelUp2
            
                X = .sngX: Y = .sngY
                ResetIt grIndex, I, X + (Rnd * 32 - 16), Y + (Rnd * 64 - 32), Rnd * 1 - 0.5, Rnd * 1 - 0.5, Rnd - 0.5, Rnd * -0.9 + 0.45, 16
                ResetColor grIndex, I, 0.1 + (Rnd * 0.1), 0.1 + (Rnd * 0.1), 0.8 + (Rnd * 0.3), 1, 0.07 + Rnd * 0.01
            
            Case 8 ' Heal
            
                X = .sngX: Y = .sngY
                ResetIt grIndex, I, X, Y, Rnd * 1.4 - 0.7, Rnd * -0.4 - 1.5, Rnd - 0.5, Rnd * -0.2 + 0.1, 16
                ResetColor grIndex, I, 0.2, 0.3, 0.9, 0.4, 0.01 + Rnd * 0.01
            
            Case 9 ' WormHole
                
                Dim lo As Integer
                Dim la As Integer
                Dim VarB(3) As Single
                For lo = 0 To 3
                    VarB(lo) = Rnd * 5
                    la = Int(Rnd * 8)
                    If la * 0.5 <> Int(la * 0.5) Then VarB(lo) = -(VarB(lo))
                Next lo
            
                Progression = Int(Rnd * 10)
                Radio = (I * 0.0125) * Progression
                X = .sngX + (Radio * Cos((I)))
                Y = .sngY + (Radio * Sin((I)))
            
                ResetIt grIndex, I, X, Y, VarB(0), VarB(1), VarB(2), VarB(3), 32
                ResetColor grIndex, I, 1, 0.6, 0.3, 1, 0.02 + Rnd * 0.3
            
            Case 10 ' Twirl
            

                Progression = Progression + Direction
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1

                Y = .sngY - 10 + Progression * Cos((I * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((I * 0.01) + Progression * 2)
                
                ResetIt grIndex, I, X, Y, 1, 1, 0, 0, 16
                ResetColor grIndex, I, 1, 0.25, 0.25, 1, 0.6 + Rnd * 0.3
            
            Case 11 ' Flower
            
                Radio = Cos(2 * (I * 0.1)) * 50
                X = .sngX + Radio * Cos(I * 0.1)
                Y = .sngY + Radio * Sin(I * 0.1)
                ResetIt grIndex, I, X, Y, 1, 1, 0, 0, 16
                ResetColor grIndex, I, 1, 0.25, 0.1, 1, 0.3 + (0.2 * Rnd) + Rnd * 0.3
            
            Case 12 ' Galaxy
                Radio = Sin(20 / (I + 1)) * 60
                X = .sngX + (Radio * Cos((I)))
                Y = .sngY + (Radio * Sin((I)))
                ResetIt grIndex, I, X, Y, 0, 0, 0, 0, 16
                ResetColor grIndex, I, 0.2, 0.2, 0.6 + 0.4 * Rnd, 1, 0 + Rnd * 0.3
            
            Case 13 ' Heart
            
                Y = .sngY - 50 * Cos(I * 0.01 * 2) * Sqr(Abs(Sin(I * 0.01)))
                X = .sngX + 50 * Sin(I * 0.01 * 2) * Sqr(Abs(Cos(I * 0.01)))
                ResetIt grIndex, I, X, Y, 0, 0, 0, -(Rnd * 0.2), 16
                ResetColor grIndex, I, 1, 0.5, 0.2, 0.6 + (0.2 * Rnd), 0.01 + Rnd * 0.08
            
            Case 14 ' BlueExplotion
            
                X = .sngX + Cos(I) * 30
                Y = .sngY + Sin(I) * 20
                ResetIt grIndex, I, X, Y, 0, -5 * (Rnd * 0.5), 0, -5 * (Rnd * 0.5), 8
                ResetColor grIndex, I, 0.3, 0.6, 1, 1, 0.05 + (Rnd * 0.1)
            
            Case 15 ' GP
            
                Radio = 50 + Rnd * 15 * Cos(I * 3.5)
                X = .sngX + (Radio * Cos((I * 0.01428571428)))
                Y = .sngY + (Radio * Sin((I * 0.01428571428)))
                ResetIt grIndex, I, X, Y, 0, 0, 0, 0, 16
                ResetColor grIndex, I, 0.2, 0.8, 0.4, 0.5, 0# + Rnd * 0.1
            
            Case 16 ' BTwirl
            

                Progression = (Progression + Direction) * Rnd
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1

                Y = .sngY - 10 + Progression * Cos((I * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((I * 0.01) + Progression * 2)
                ResetIt grIndex, I, X, Y, 1, 1, 0, 0, 8
                ResetColor grIndex, I, 0.25, 0.25, 1, 1, 0.1 + Rnd * 0.3 + Rnd * 0.3
            
            Case 17 ' BT

                Progression = Progression + Direction
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1

                Y = .sngY - 10 + Progression * Cos((I * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((I * 0.01) + Progression * 2)
                ResetIt grIndex, I, X, Y, -10, -1 * Rnd, 0, Rnd, 8
                ResetColor grIndex, I, 0.25, 0.25, 1, 1, 0.1 + (0.1 * Rnd) + Rnd * 0.3
            
            Case 18 ' Atomic
            
                Radio = 10 + Sin(2 * (I * 0.1)) * 50
                X = .sngX + Radio * Cos(I * 0.033333)
                Y = .sngY + Radio * Sin(I * 0.033333)
                ResetIt grIndex, I, X, Y, 1, 1, 0, 0, 8
                ResetColor grIndex, I, 0.4, 0.25, 1, 1, 0.3 + (0.2 * Rnd) + Rnd * 0.3
                
            Case 19 ' Medit
            
                X = .sngX + Cos(I * Rnd) * 45 * Sin(I * Rnd)
                Y = .sngY
                ResetIt grIndex, I, X, Y, 1, 1, 0, -10, 30
                ResetColor grIndex, I, 0, 0.5, 0.1, 1, 0.01 + (0.2 * Rnd)
            
    End Select
        
End With
    
End Sub
Public Sub UpdateParticleGroup(grIndex As Integer, sngNewX As Single, sngNewY As Single)
    Dim I As Long
    Dim sngElapsedTime As Single
    
    With ParticleGroup(grIndex)
     
        ' We calculate the time difference here
        sngElapsedTime = (GetTickCount() - .lngPreviousFrame) / 100
        .lngPreviousFrame = GetTickCount()
        
        .sngX = sngNewX
        .sngY = sngNewY
         
        For I = 0 To .ParticleCounts
            With .Particles(I)
                UpdateParticle grIndex, I, sngElapsedTime
                 
                ' If the particle is invisible, reset it again.
                If .sngA <= 0 Then
                    Reset grIndex, I
                End If
                
                ParticleGroup(grIndex).vertsPoints(I).rhw = 1
                ParticleGroup(grIndex).vertsPoints(I).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                ParticleGroup(grIndex).vertsPoints(I).sX = .sngX
                ParticleGroup(grIndex).vertsPoints(I).sY = .sngY
                
            End With
            
        Next I
        
    End With
    
End Sub

' FUNCTIONS FOR ONLY PARTICLE

Private Sub ResetColor(grIndex As Integer, Particle As Long, sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
    ' Reset color to the new values
    With ParticleGroup(grIndex).Particles(Particle)
        .sngR = sngRed
        .sngG = sngGreen
        .sngB = sngBlue
        .sngA = sngAlpha
        .sngAlphaDecay = sngDecay
    End With
End Sub
 
Private Sub ResetIt(grIndex As Integer, Particle As Long, X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, sngResetSize As Single)
    
    With ParticleGroup(grIndex).Particles(Particle)
        .sngX = X
        .sngY = Y
        .sngXSpeed = XSpeed
        .sngYSpeed = YSpeed
        .sngXAccel = XAcc
        .sngYAccel = YAcc
        .sngSize = sngResetSize
    End With
    
End Sub
 
Private Sub UpdateParticle(grIndex As Integer, Particle As Long, sngTime As Single)
    
    With ParticleGroup(grIndex).Particles(Particle)
        .sngX = .sngX + .sngXSpeed * sngTime
        .sngY = .sngY + .sngYSpeed * sngTime
     
        .sngXSpeed = .sngXSpeed + .sngXAccel * sngTime
        .sngYSpeed = .sngYSpeed + .sngYAccel * sngTime
     
        .sngA = .sngA - .sngAlphaDecay * sngTime
    End With
    
End Sub
