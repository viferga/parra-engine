VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Parra Engine"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEditor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3090
      Begin VB.ComboBox cmbMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00040404&
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   30
         List            =   "frmMain.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3315
         Width           =   3015
      End
      Begin VB.ListBox grhList 
         Appearance      =   0  'Flat
         BackColor       =   &H000A0A0A&
         ForeColor       =   &H00C0C0C0&
         Height          =   2175
         ItemData        =   "frmMain.frx":006B
         Left            =   30
         List            =   "frmMain.frx":006D
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   855
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MousePicX As Single: Private MousePicY As Single
Private cMouse As Boolean

Public WithEvents Winsock As clsSocket
Attribute Winsock.VB_VarHelpID = -1
Private Sub Form_Load()
    Me.picEditor.Picture = LoadPicture(App.Path & "\ui\menu.jpg")
    Sound_Play "mambo.mp3", 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    #If WorldEditor = 1 Then
        If EditMap = True Then
            modMap.mapAddGrh 3 'layer..
        End If
    #End If
    
    cMouse = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If cMouse = True Then
        guiEvents X, Y
    End If

    Mouse.X = X + RenderRect.X1
    Mouse.Y = Y + RenderRect.Y1
    
    'Trim to fit screen
    If Mouse.X < 0 Then
        Mouse.X = 0
    ElseIf Mouse.X > RenderRect.X2 Then
        Mouse.X = RenderRect.X2
    End If
    
    'Trim to fit screen
    If Mouse.Y < 0 Then
        Mouse.Y = 0
    ElseIf Mouse.Y > RenderRect.Y2 Then
        Mouse.Y = RenderRect.Y2
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cMouse = False
    
End Sub

Private Sub picEditor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePicX = X
    MousePicY = Y
End Sub

Private Sub picEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        picEditor.Top = picEditor.Top + Y - MousePicY
        picEditor.Left = picEditor.Left + X - MousePicX
    End If
End Sub

' ########################## W I N S O C K  ##########################

Private Sub Winsock_Connect()
    'Clean input and output buffers
    incomingData.ReadASCIIStringFixed (incomingData.Length)
    outgoingData.ReadASCIIStringFixed (outgoingData.Length)
    
    Select Case LoginStatus
        Case LogStatus.None

        Case LogStatus.OnAcc

        Case LogStatus.OnPj
        
    End Select
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    'Dim RD As String
    Dim Data() As Byte
    
    Winsock.GetData Data, vbByte, bytesTotal
        
    'Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    HandleIncomingData
End Sub
Private Sub Winsock_CloseSck()
    
    If Winsock.State <> sckClosed Then _
        Winsock.CloseSck
        
    If frmConnect.Visible = True Then frmConnect.ConnectShow
    
End Sub
Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' Handle socket errors

    If Number = 10061 Then
        MsgBox "No se ha podido establecer conexión con el servidor.", vbCritical
    Else
        MsgBox Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error"
    End If
    
    If Winsock.State <> sckClosed Then _
        Winsock.CloseSck
End Sub
