Attribute VB_Name = "modParticle"
Option Explicit

' Pi...
Public Const ConstPI As Single = 3.14159
     
' To convert to Radians
Public Const RAD As Single = ConstPI / 180
     
' To convert to Degrees
Public Const DEG As Single = 180 / ConstPI

' Public StartTime
Public lngStartTime As Long
     
Public myTexture As Direct3DTexture8
     
' TL Vertex
Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR)
    
'  V E R T E X   B U F F E R

Private m_vertBuff As Direct3DVertexBuffer8  'we assume this has been created
Private m_vertCount As Long                 'we assume this has been set

Private addr As Long                        'will holds the address the D3D
                                                   'managed memory
'Private verts() As D3DTLVERTEX                'array that we want to point to
                                                   'D3D managed memory

 
 'ONLY FOR PARTICLE
 
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

'ONLY ON GROUP

Private Type structGroupParticle

    ParticleCounts  As Long
    Particles()     As structParticle
    vertsPoints()   As D3DTLVERTEX
    
    ' POSITION GROUP & INFO
    sngX                As Single
    sngY                As Single
    sngProgression      As Single
 
    lngFloat0           As Long
    lngFloat1           As Long
    lngFloatSize        As Long
    
    lngPreviousFrame    As Long
 
    
End Type

Public ParticleGroup() As structGroupParticle

Public Sub loadGroupParticle()

    ReDim Preserve ParticleGroup(1 To 1) As structGroupParticle
    
    ParticleGroup(1).ParticleCounts = 100
    ReLocate 1, 150, 150
    Begin 1

End Sub
Public Sub Begin(grIndex As Integer)
    '//We initialize our stuff here
    Dim i As Long
     
    With ParticleGroup(grIndex)
    
        .lngFloat0 = FtoDW(0)
        .lngFloat1 = FtoDW(1)
        .lngFloatSize = FtoDW(20) '//Size of our flame particles..
         
        ' Redim our particles to the particlecount
        ReDim .Particles(0 To .ParticleCounts)
         
        ' Redim vertices to the particle count
        ' Point sprites, so 1 per particle
        ReDim .vertsPoints(0 To .ParticleCounts)
             
        Set m_vertBuff = g_dev.CreateVertexBuffer(Len(.vertsPoints(0)) * .ParticleCounts, 0, D3DFVF_TLVERTEX, D3DPOOL_MANAGED)
        
        m_vertBuff.Lock 0, Len(.vertsPoints(0)) * .ParticleCounts, addr, 0
        
        DXLockArray8 m_vertBuff, addr, .vertsPoints
        
        ' Now generate all particles
        For i = 0 To .ParticleCounts - 1
            Reset grIndex, i
        Next i

        DXUnlockArray8 m_vertBuff, .vertsPoints
        
        m_vertBuff.Unlock
        
        ' Set initial time
        .lngPreviousFrame = GetTickCount()
    
    End With
    
End Sub
 
Public Sub Reset(grIndex As Integer, i As Long) ' Reset GROUP
     
    With ParticleGroup(grIndex)
             
        ' This is were we will reset individual particles.
        ResetIt grIndex, i, .sngX, .sngY, frmMain.XSpeed + (frmMain.XSpeed * Rnd), frmMain.YSpeed + (frmMain.YSpeed * Rnd), frmMain.XAcc + (frmMain.XAcc * Rnd), frmMain.YAcc + (frmMain.YAcc * Rnd), 2
        ResetColor grIndex, i, frmMain.Red / 100, frmMain.Green / 100, frmMain.Blue / 100, frmMain.Alpha / 100, frmMain.AlphaDecay / 100 + (0.2 * Rnd)
        
    End With
End Sub
 
Public Sub Update(grIndex As Integer)
    Dim i As Long
    Dim sngElapsedTime As Single
    
    With ParticleGroup(grIndex)
     
        '//We calculate the time difference here
        sngElapsedTime = (GetTickCount() - .lngPreviousFrame) / 100
        .lngPreviousFrame = GetTickCount()
         
        For i = 0 To .ParticleCounts
            With .Particles(i)
                UpdateParticle grIndex, i, sngElapsedTime
                 
                '//If the particle is invisible, reset it again.
                If .sngA <= 0 Then
                    Reset grIndex, i
                End If
                
                ParticleGroup(grIndex).vertsPoints(i).rhw = 1
                ParticleGroup(grIndex).vertsPoints(i).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                ParticleGroup(grIndex).vertsPoints(i).sx = .sngX
                ParticleGroup(grIndex).vertsPoints(i).sy = .sngY
                
            End With
            
        Next i
        
        D3DVertexBuffer8SetData m_vertBuff, 0, Len(.vertsPoints(0)) * .ParticleCounts, 0, .vertsPoints(0)

    End With
End Sub
 
Public Sub Render(grIndex As Integer)
    With g_dev
        '//Set the render states for using point sprites
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
         
        '//Set up the vertex shader
        .SetVertexShader D3DFVF_TLVERTEX
         
        '//Set our texture
        .SetTexture 0, myTexture
        
        .SetStreamSource 0, m_vertBuff, Len(ParticleGroup(grIndex).vertsPoints(0))
         
        '//And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleGroup(grIndex).ParticleCounts, _
            ParticleGroup(grIndex).vertsPoints(0), Len(ParticleGroup(grIndex).vertsPoints(0))
         
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 0 'False
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'False
        
    End With
    
   '        m_vertBuff.Lock 0, Len(verts(0)) * m_vertCount, addr, 0

   '        DXLockArray8 m_vertBuff, addr, verts
   '
   '        Dim i As Long

   '        For i = 0 To m_vertCount - 1
   '            verts(i).X = i 'or what ever you want to dow with the data
   '        Next

   '        DXUnlockArray8 m_vertBuff, verts

   '        m_vertBuff.Unlock
    
End Sub

' FUNCTIONS FOR ONLY GROUPS

Public Sub ReLocate(grIndex As Integer, sngNewX As Single, sngNewY As Single) ' RELOCATE GROUP
    ParticleGroup(grIndex).sngX = sngNewX
    ParticleGroup(grIndex).sngY = sngNewY
End Sub

' FUNCTIONS FOR ONLY PARTICLE

Public Sub ResetColor(grIndex As Integer, Particle As Long, sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
    ' Reset color to the new values
    With ParticleGroup(grIndex).Particles(Particle)
        .sngR = sngRed
        .sngG = sngGreen
        .sngB = sngBlue
        .sngA = sngAlpha
        .sngAlphaDecay = sngDecay
    End With
End Sub
 
Public Sub ResetIt(grIndex As Integer, Particle As Long, X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, sngResetSize As Single)
    
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
 
Public Sub UpdateParticle(grIndex As Integer, Particle As Long, sngTime As Single)
    
    With ParticleGroup(grIndex).Particles(Particle)
        .sngX = .sngX + .sngXSpeed * sngTime
        .sngY = .sngY + .sngYSpeed * sngTime
     
        .sngXSpeed = .sngXSpeed + .sngXAccel * sngTime
        .sngYSpeed = .sngYSpeed + .sngYAccel * sngTime
     
        .sngA = .sngA - .sngAlphaDecay * sngTime
    End With
    
End Sub

