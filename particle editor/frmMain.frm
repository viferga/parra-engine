VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Particles"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit

Dim ccount As Long
Dim cclast As Long
Private Sub Form_KeyPress(KeyAscii As Integer)
    Form_Unload 0
End Sub
 
Private Sub Form_Load()
    Dim m_bInit As Boolean
     
    Randomize
     
    Me.Show
    DoEvents
     
    m_bInit = D3DUtil_Init(Me.hwnd, True, 0, 0, D3DDEVTYPE_HAL, Me)
    If Not (m_bInit) Then End
 
    '//Reset states
    ResetStates

    loadGroupParticle
 
    Set myTexture = g_d3dx.CreateTextureFromFileEx(g_dev, App.Path & "\particle.bmp", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
     
    Do
 
         
        g_dev.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1, 0
        g_dev.BeginScene
         
            Update 1
            Render 1
         
        g_dev.EndScene
        g_dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
                
        DoEvents
        
        If GetTickCount() - cclast >= 1000 Then
            cclast = GetTickCount()
            Debug.Print ccount
            ccount = 0
        End If
        
        ccount = ccount + 1
    Loop
     
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

