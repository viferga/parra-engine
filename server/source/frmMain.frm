VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Servidor"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsEsc 
      Left            =   420
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   666
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   666
   End
   Begin VB.TextBox Command 
      BackColor       =   &H000F0F0F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox Logo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1200
      ScaleHeight     =   1335
      ScaleWidth      =   5895
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin VB.ListBox Consola 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label lblWorldTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblOnline 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Conexiones: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   6120
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   3570
      Left            =   345
      Top             =   1545
      Width           =   7485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Commands:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5265
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If frmMain.Command.Visible = False Then
            frmMain.Command.Visible = True
            frmMain.Command.SetFocus
        Else
            Select Case LCase$(frmMain.Command.Text)
                Case "exit"
                    'Close Server
                    prgRun = False
                    #If Testing = 1 Then
                     modGame.gameUnLoad
                    #End If
                    Exit Sub
                Case "msg"
                    ' HandleSendMessage sckIndex, MessageType.MessageBox, frmMain.Command.Text
                    Exit Sub
            End Select
            
            frmMain.Command.Text = vbNullString
            frmMain.Command.Visible = False
            frmMain.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    frmMain.Logo.Picture = LoadPicture(App.Path & "\Logo.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    prgRun = False
    
    #If Testing = 1 Then
        modGame.gameUnLoad
    #End If
End Sub
Private Sub Winsock_Close(Index As Integer)
    sckClose Index
End Sub
Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    'Send info to sckDataArrival Event
    sckDataArrival Index, bytesTotal
    
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckClose Index
End Sub

Private Sub wsEsc_ConnectionRequest(ByVal RID As Long)

    Dim sckIndex As Integer

    For sckIndex = 1 To sckMax
        
        If Winsock(sckIndex).State = sckClosed Then
            
            If (sckIndex < sckMax) Then
            
                'Open next connexion
                sckOpen sckIndex, RID
            
            Else
                HandleSendMessage sckIndex, MessageType.MessageBox, "El server se encuentra lleno en este momento."
                DoEvents
                sckClose sckIndex
            End If

            
            Exit For
        End If
    
    Next sckIndex
    
End Sub

