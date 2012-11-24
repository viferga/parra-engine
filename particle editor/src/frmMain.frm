VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "Particle Editor"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Color Editor"
      Height          =   3375
      Left            =   9840
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
      Begin VB.HScrollBar AlphaDecay 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.HScrollBar Alpha 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.HScrollBar Blue 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.HScrollBar Green 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.HScrollBar Red 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alpha Decay"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alpha"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movement Editor"
      Height          =   2775
      Left            =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.HScrollBar XAcc 
         Height          =   255
         Left            =   360
         Max             =   100
         Min             =   -100
         TabIndex        =   6
         Top             =   1680
         Width           =   2415
      End
      Begin VB.HScrollBar YAcc 
         Height          =   255
         Left            =   360
         Max             =   100
         Min             =   -100
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.HScrollBar YSpeed 
         Height          =   255
         Left            =   360
         Max             =   100
         Min             =   -100
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.HScrollBar XSpeed 
         Height          =   255
         Left            =   360
         Max             =   100
         Min             =   -100
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X Acceleration"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Y Acceleration"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Y Speed"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X Speed"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStreamNew 
         Caption         =   "&New Stream"
      End
      Begin VB.Menu mnuStreamOpen 
         Caption         =   "&Open Stream File"
      End
      Begin VB.Menu mnuStreamAllSave 
         Caption         =   "&Save All Streams"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Run As Boolean

Private Sub Form_Load()
    
    If Not D3DUtil_Init(Me.hwnd, True, 0, 0, D3DDEVTYPE_HAL, Me) Then End
    
    D3D_ResetRenderStates
    
    loadGroupParticle
    
    Set myTexture = g_d3dx.CreateTextureFromFileEx(g_dev, App.Path & "\Graphics\Particle.bmp", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)

    Me.Show

    DoEvents
    
    Run = True
    
    Dim RenderRect As D3DRECT: With RenderRect: .X1 = 0: .X2 = 500: .Y1 = 0: .Y2 = 400: End With
    
    Dim LastCount As Long
    Dim fpsCount  As Long

    Do
        If Run = False Then Exit Do
        
        g_dev.Clear 1, RenderRect, D3DCLEAR_TARGET, 0, 1, 0
        g_dev.BeginScene

            Update 1
            Render 1
         
        g_dev.EndScene
        g_dev.Present RenderRect, RenderRect, 0, ByVal 0&
                
        DoEvents
        
        'update the scene stats once per second
        If GetTickCount() - LastCount >= 1000& Then
            Me.Caption = "Particle Editor - FPS: " & CStr(fpsCount)
            LastCount = GetTickCount()
            fpsCount = 1&
        End If
        
        'show frame rate and device statistics
        fpsCount = fpsCount + 1&
    
    Loop
    
    D3DUtil_Destroy
    
    End
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Run = False
End Sub

