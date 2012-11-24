VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFireFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Particle"
Attribute VB_Ext_KEY = "Member0" ,"Particle"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
 
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

Private Particles() As structParticle

'ONLY ON GROUP

Private vertsPoints() As typeTRANSLITVERTEX
  
Private sngX As Single, sngY As Single ' POSITION GROUP
Public ParticleCounts As Long
 
Private lngFloat0 As Long
Private lngFloat1 As Long
Private lngFloatSize As Long
 
Private lngPreviousFrame As Long
 
Public Sub Begin()
    '//We initialize our stuff here
    Dim i As Long
     
    lngFloat0 = FtoDW(0)
    lngFloat1 = FtoDW(1)
    lngFloatSize = FtoDW(25) '//Size of our flame particles..
     
    '//Redim our particles to the particlecount
    ReDim Particles(0 To ParticleCounts)
     
    '//Redim vertices to the particle count
    '//Point sprites, so 1 per particle
    ReDim vertsPoints(0 To ParticleCounts)
         
    '//Now generate all particles
    For i = 0 To ParticleCounts
        vertsPoints(i).rhw = 1
        Call Reset(i)
    Next i
     
    '//Set initial time
    lngPreviousFrame = timeGetTime
End Sub
 
Public Sub Reset(i As Long)
    Dim X As Single, Y As Single
     
    X = sngX + (Rnd * 10)
    Y = sngY
     
    ' This is were we will reset individual particles.
    ResetIt i, X, Y, -0.4 + (Rnd * 0.8), -0.5 - (Rnd * 0.4), 0, -(Rnd * 0.3), 2
    ResetColor i, 1, 0.5, 0.2, 0.6 + (0.2 * Rnd), 0.01 + Rnd * 0.05
End Sub
 
Public Sub Update()
    Dim i As Long
    Dim sngElapsedTime As Single
     
    '//We calculate the time difference here
    sngElapsedTime = (timeGetTime - lngPreviousFrame) / 100
    lngPreviousFrame = timeGetTime
     
    For i = 0 To ParticleCounts
        With Particles(i)
            UpdateParticle i, sngElapsedTime
             
            '//If the particle is invisible, reset it again.
            If .sngA <= 0 Then
                Call Reset(i)
            End If
         
            vertsPoints(i).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            vertsPoints(i).X = .sngX
            vertsPoints(i).Y = .sngY
         
        End With
    Next i
End Sub
 
Public Sub Render()
    With g_dev
        '//Set the render states for using point sprites
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetRenderState D3DRS_POINTSIZE, lngFloatSize
        .SetRenderState D3DRS_POINTSIZE_MIN, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_A, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_B, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_C, lngFloat1
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
         
        '//Set up the vertex shader
        .SetVertexShader D3DFVF_TLVERTEX
         
        '//Set our texture
        .SetTexture 0, myTexture
         
        '//And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleCounts / 2, vertsPoints(0), Len(vertsPoints(0))
         
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 0 'False
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'False
    End With
End Sub

' FUNCTIONS FOR ONLY GROUPS

Public Sub ReLocate(sngNewX As Single, sngNewY As Single) ' RELOCATE GROUP
    sngX = sngNewX
    sngY = sngNewY
End Sub

' FUNCTIONS FOR ONLY PARTICLE

Public Sub ResetColor(Particle As Long, sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
    ' Reset color to the new values
    With Particles(Particle)
        .sngR = sngRed
        .sngG = sngGreen
        .sngB = sngBlue
        .sngA = sngAlpha
        .sngAlphaDecay = sngDecay
    End With
End Sub
 
Public Sub ResetIt(Particle As Long, X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, sngResetSize As Single)
    
    With Particles(Particle)
        .sngX = X
        .sngY = Y
        .sngXSpeed = XSpeed
        .sngYSpeed = YSpeed
        .sngXAccel = XAcc
        .sngYAccel = YAcc
        .sngSize = sngResetSize
    End With
End Sub
 
Public Sub UpdateParticle(Particle As Long, sngTime As Single)
    
    With Particles(Particle)
        .sngX = .sngX + .sngXSpeed * sngTime
        .sngY = .sngY + .sngYSpeed * sngTime
     
        .sngXSpeed = .sngXSpeed + .sngXAccel * sngTime
        .sngYSpeed = .sngYSpeed + .sngYAccel * sngTime
     
        .sngA = .sngA - .sngAlphaDecay * sngTime
    End With
    
End Sub
