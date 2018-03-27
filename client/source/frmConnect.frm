VERSION 5.00
Begin VB.Form frmConnect 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Connecting"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPlayers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   390
      TabIndex        =   8
      Top             =   2235
      Width           =   2820
   End
   Begin VB.TextBox txtKillMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000000F&
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
      Height          =   270
      Left            =   4785
      TabIndex        =   7
      Top             =   5235
      Width           =   2400
   End
   Begin VB.TextBox txtKillPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000119&
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
      Height          =   270
      Left            =   4800
      TabIndex        =   6
      Top             =   4410
      Width           =   2400
   End
   Begin VB.TextBox txtKillName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000119&
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
      Height          =   270
      Left            =   4785
      TabIndex        =   5
      Top             =   3555
      Width           =   2400
   End
   Begin VB.TextBox txtCreateMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000000F&
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
      Height          =   270
      Left            =   4770
      TabIndex        =   4
      Top             =   5235
      Width           =   2400
   End
   Begin VB.TextBox txtCreatePass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000119&
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
      Height          =   270
      Left            =   4785
      TabIndex        =   3
      Top             =   4425
      Width           =   2400
   End
   Begin VB.TextBox txtCreateName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000119&
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
      Height          =   270
      Left            =   4785
      TabIndex        =   2
      Top             =   3570
      Width           =   2400
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000115&
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
      Height          =   270
      Left            =   4860
      TabIndex        =   1
      Top             =   4830
      Width           =   2340
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000119&
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
      Height          =   270
      Left            =   4830
      TabIndex        =   0
      Top             =   3990
      Width           =   2400
   End
   Begin VB.Image Kill 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4365
      Top             =   5655
      Width           =   1800
   End
   Begin VB.Image KillCancel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6465
      Top             =   5670
      Width           =   1185
   End
   Begin VB.Image main3 
      Height          =   3735
      Left            =   3720
      Top             =   2685
      Width           =   4665
   End
   Begin VB.Image CreateCancel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6420
      Top             =   5670
      Width           =   1185
   End
   Begin VB.Image create 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4470
      Top             =   5640
      Width           =   1635
   End
   Begin VB.Image main2 
      Height          =   3735
      Left            =   3720
      Top             =   2685
      Width           =   4665
   End
   Begin VB.Image Connect 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4725
      Top             =   5250
      Width           =   1230
   End
   Begin VB.Image Cancel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6105
      Top             =   5235
      Width           =   1230
   End
   Begin VB.Image mainPic 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   3
      Left            =   4710
      Top             =   4935
      Width           =   2625
   End
   Begin VB.Image mainPic 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   2
      Left            =   4710
      Top             =   4500
      Width           =   2625
   End
   Begin VB.Image mainPic 
      Appearance      =   0  'Flat
      Height          =   465
      Index           =   1
      Left            =   4710
      Top             =   4050
      Width           =   2625
   End
   Begin VB.Image mainPic 
      Appearance      =   0  'Flat
      Height          =   465
      Index           =   0
      Left            =   4710
      Top             =   3600
      Width           =   2625
   End
   Begin VB.Image main 
      Height          =   3735
      Left            =   3720
      Top             =   2685
      Width           =   4665
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************
'    Parra Engine is a MMORPG Isometric Game Engine.
'    Copyright (C) 2009 - 2013 Vicente Eduardo Ferrer Garcia
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as
'    published by the Free Software Foundation, either version 3 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'****************************************************************************

Public UserName As String

Dim i As Byte
Public Sub ConnectShow()

    'Show mainpic
    For i = 0 To 3
        mainPic(i).Visible = True
    Next i
    
    'Change AccountStatus
    accountStatus = accNone
    
    'cancelclick
    main.Visible = False
    Cancel.Visible = False
    Connect.Visible = False
    
    txtName.Visible = False
    txtPassword.Visible = False
    
    txtName.Text = ""
    txtPassword.Text = ""

    'cancelconnect
    main2.Visible = False
    create.Visible = False
    CreateCancel.Visible = False
    
    txtCreateName.Visible = False
    txtCreatePass.Visible = False
    txtCreateMail.Visible = False
    
    txtCreateName.Text = ""
    txtCreatePass.Text = ""
    txtCreateMail.Text = ""
    
    'cancelkill
    main3.Visible = False
    Kill.Visible = False
    KillCancel.Visible = False
            
    txtKillName.Visible = False
    txtKillPass.Visible = False
    txtKillMail.Visible = False
            
    txtKillName.Text = ""
    txtKillPass.Text = ""
    txtKillMail.Text = ""

End Sub
Private Sub Cancel_Click()
    ConnectShow
End Sub

Private Sub Connect_Click()
    If txtName.Text <> "" And _
            txtPassword.Text <> "" Then WriteOutgoingData ClientPacketID.AccountEvents
End Sub

Private Sub Create_Click()
    If txtCreateName.Text <> "" And _
            txtCreatePass.Text <> "" And txtCreateMail.Text <> "" Then WriteOutgoingData ClientPacketID.AccountEvents
End Sub

Private Sub CreateCancel_Click()
    ConnectShow
End Sub

Private Sub Form_Load()
        
    Me.Picture = LoadPicture(App.Path & "\ui\frmconnect.jpg")
    main.Picture = LoadPicture(App.Path & "\ui\connectmain2.jpg")
    main2.Picture = LoadPicture(App.Path & "\ui\connectmain3.jpg")
    main3.Picture = LoadPicture(App.Path & "\ui\connectmain4.jpg")
    
    txtName.Text = GetVar(App.Path & "\cliente.ini", "CHARACTER", "Acc")
    txtPassword.Text = GetVar(App.Path & "\cliente.ini", "CHARACTER", "Pass")
    
    Cancel.Visible = False
    Connect.Visible = False
    
    txtName.Visible = False
    txtPassword.Visible = False
    
    main.Visible = False
    main2.Visible = False
    main3.Visible = False
    
    create.Visible = False
    CreateCancel.Visible = False
    txtCreateName.Visible = False
    txtCreatePass.Visible = False
    txtCreateMail.Visible = False
    
    Kill.Visible = False
    KillCancel.Visible = False
    txtKillName.Visible = False
    txtKillPass.Visible = False
    txtKillMail.Visible = False
    
    'listplayer
    lstPlayers.Clear
    lstPlayers.Visible = False
    
    'Change AccountStatus
    accountStatus = accState.accNone
    
    
    
End Sub
Private Sub Kill_Click()
    If txtKillName.Text <> "" And _
            txtKillPass.Text <> "" And txtKillMail.Text <> "" Then WriteOutgoingData ClientPacketID.AccountEvents
End Sub

Private Sub KillCancel_Click()
    ConnectShow
End Sub
Private Sub lstPlayers_DblClick()
    playerStatus = playerState.plyLogin
    
    WriteOutgoingData ClientPacketID.UserEvents
    
End Sub

Private Sub mainPic_Click(Index As Integer)
    
    For i = 0 To 3
        mainPic(i).Visible = False
    Next i
    
    If Index <> 3 Then
        If frmMain.Winsock.State <> (sckConnected Or sckConnecting) Then
        
            'Connect to server
            frmMain.Winsock.Connect generalIP, generalPort
            
        End If
    End If
    
    Select Case Index
        Case 0
            Cancel.Visible = True
            Connect.Visible = True
            txtName.Visible = True
            txtPassword.Visible = True
            main.Visible = True
            
            'Change AccountStatus
            accountStatus = accLogin
        Case 1
            main2.Visible = True
            create.Visible = True
            CreateCancel.Visible = True
            
            txtCreateName.Visible = True
            txtCreatePass.Visible = True
            txtCreateMail.Visible = True
            
            'Change AccountStatus
            accountStatus = accCreate
        Case 2
            main3.Visible = True
            Kill.Visible = True
            KillCancel.Visible = True
            
            txtKillName.Visible = True
            txtKillPass.Visible = True
            txtKillMail.Visible = True
            
            'Change AccountStatus
            accountStatus = accKill
        Case 3
            bRunning = False
            
            'Change AccountStatus
            accountStatus = accExit
    End Select

End Sub
