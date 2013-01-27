Attribute VB_Name = "modProtocol"
Option Explicit

'   Handle:        The Server sends the Client
'   Write:         The Client sends the Server

'   Binary protocol management
Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

Public Enum LogStatus
    None = 0
    OnAcc = 1
    OnPj = 2
End Enum: Public LoginStatus As LogStatus

Public Enum accState 'Rutinas Account
    accNone = 0
    accLogin = 1
    accCreate = 2
    accKill = 3
    accExit = 4
End Enum: Public accountStatus As accState

Public Enum playerState 'Rutinas del Usuario
    plyNone = 0
    plyLogin
    plyCreate
    plyKill
    plyExit
End Enum: Public playerStatus As playerState

Private Const SEPARATOR As String * 1 = vbNullChar

Public Enum ServerPacketID
    AccountEvents = 1
    UserEvents
    HandleMessage
    CharacterEvent
End Enum

Public Enum ClientPacketID
    AccountEvents = 1
    UserEvents
    CharEvents
End Enum

' Handles incoming data.

Public Sub HandleIncomingData()

    Select Case incomingData.PeekByte()
        Case ServerPacketID.AccountEvents
                HandleIncomingAccount
                Exit Sub
                
        Case ServerPacketID.UserEvents
                HandleIncomingUser
                Exit Sub
        
        Case ServerPacketID.HandleMessage
                HandleIncomingMessage
                Exit Sub
                
        Case ServerPacketID.CharacterEvent
            'Remove packet ID
            incomingData.ReadByte
            
                If incomingData.ReadByte() = 1 Then
                    HandleIncomingCharacterCreate
                Else
                    HandleIncomingCharacterRemove
                End If
        
       ' Case ServerPacketID.UserEvents
         '       HandleCharacterMove
         '       Exit Sub
                
        Case Else: Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        HandleIncomingData
    End If
    
End Sub
Private Sub HandleIncomingAccount()
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim packetByte As Byte
    
        packetByte = incomingData.ReadByte()
    
        Select Case packetByte
        
            Case 1 'Connect
                
                Dim NumPlayers As Byte, I As Long
                
                NumPlayers = incomingData.ReadByte
                
                For I = 1 To NumPlayers
                    frmConnect.lstPlayers.AddItem incomingData.ReadASCIIString()
                Next I
            
                frmConnect.lstPlayers.Visible = True
                
            Case 2 'Create
               MsgBox "Cuenta Creada"
               
            Case 3 'Kill
               MsgBox "Cuenta Borrada"
               
            Case 4
            
                If frmMain.Winsock.State <> sckClosed Then _
                    frmMain.Winsock.CloseSck
                    
                frmMain.Visible = False
                frmConnect.Show
                frmConnect.ConnectShow
            
        End Select
    
    frmConnect.ConnectShow
End Sub
Private Sub HandleIncomingUser()

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim packetByte As Byte
    
        packetByte = incomingData.ReadByte()
    
        Select Case packetByte
                Case 1 'connect
                
                    frmMain.Visible = True
                    frmConnect.Visible = False
                    gamePaused = False
                
                Case 2 'create
                    
                    'add list index
                    
                Case 3 'kill
                    
                    'remove list index
                    
                Case 4 'exit
                   
                    frmConnect.Visible = True
                    frmConnect.lstPlayers.Visible = True
                    playerStatus = playerState.plyNone
                        
        End Select

End Sub
Private Sub HandleIncomingMessage()
    'Remove packet ID
    Call incomingData.ReadByte

    Dim MessageType As Byte, Message As String
        
        MessageType = incomingData.ReadByte()
        Message = incomingData.ReadASCIIString()
        
            Select Case MessageType
                    
                Case 0 'messagebox
                    
                    MsgBox Message, vbOKOnly, "Mensaje Servidor"
                
                Case 1 'console
                
                    
            End Select
        
End Sub
Private Sub HandleIncomingCharacterCreate()

    'Remove packet ID
    
        Dim charindex As Integer
        charindex = incomingData.ReadInteger()
        
        If charLast > 0 Then
            If charindex > charLast Then charLast = charindex
            ReDim Preserve characterList(1 To charLast) As Character
        Else
            ReDim characterList(1) As Character
            charLast = 1
        End If
            
        
        With characterList(charindex)
            
            .Name = incomingData.ReadASCIIString()
            .Body = incomingData.ReadInteger()
            .Head = incomingData.ReadInteger()
            .Heading = incomingData.ReadByte()
            .Pos.X = incomingData.ReadByte()
            .Pos.Y = incomingData.ReadByte()
            
            If .Name = frmConnect.UserName Then
                playerCharIndex = charindex
                
                UserPos.X = .Pos.X
                UserPos.Y = .Pos.Y
            
            End If
        
            .Active = 1
            
        End With
End Sub

Private Sub HandleIncomingCharacterRemove()

    Dim charindex As Integer
    charindex = incomingData.ReadInteger()

    With characterList(charindex)
    
        .Active = 0
        
        'Update lastchar
        If charindex = charLast Then
            Do Until .Active = 1
                charLast = charLast - 1
                If charLast = 0 Then Exit Do
            Loop
        End If
        
        mapData(.Pos.X, .Pos.Y).charindex = 0

        .Body = 0
        .FXIndex = 0
        .Head = 0
        .Heading = 0
        .Pos.X = 0
        .Pos.Y = 0
        .Name = ""
        .Moving = 0
    
    End With
    
    'Redimensionamos el array
    If charLast > 0 Then
        ReDim Preserve characterList(1 To charLast) As Character
    Else
        ReDim characterList(0) As Character
    End If
End Sub

Public Sub WriteOutgoingData(ByRef Packed As ClientPacketID)

    outgoingData.WriteByte Packed

    Select Case Packed
        Case ClientPacketID.AccountEvents
            WriteOutgoingAccount
            Exit Sub
            
        Case ClientPacketID.UserEvents
            WriteOutgoingUser
            Exit Sub
        
        'Case ClientPacketID.CharEvents 'lo iva a poner aca pero habia que agregar algunos parametros, asi q mejor no GDK
        '    WriteCharEvents
        '    Exit Sub
            
        Case Else: Exit Sub
        
    End Select
    
    'Done with this packet, move on to next one
    If outgoingData.Length > 0 And Err.Number <> outgoingData.NotEnoughDataErrCode Then
        Err.Clear
        WriteOutgoingData Packed
    End If
End Sub
Private Sub WriteOutgoingAccount()
    
    'Send packet ID
    With outgoingData
    
        .WriteByte accountStatus
    
        Select Case accountStatus
            Case accState.accLogin
    
                    .WriteASCIIString frmConnect.txtName.Text
                    .WriteASCIIString frmConnect.txtPassword.Text
                
            Case accState.accCreate
            
                    .WriteASCIIString frmConnect.txtCreateName.Text
                    .WriteASCIIString frmConnect.txtCreatePass.Text
                    .WriteASCIIString frmConnect.txtCreateMail.Text
                    
            Case accState.accKill
            
                    .WriteASCIIString frmConnect.txtKillName.Text
                    .WriteASCIIString frmConnect.txtKillPass.Text
                    .WriteASCIIString frmConnect.txtKillMail.Text
            
        End Select
    
    End With
    
End Sub
Private Sub WriteOutgoingUser()
    
    With outgoingData
    
        .WriteByte playerStatus
        
        Select Case playerStatus
            Case playerState.plyLogin
                If frmConnect.lstPlayers.List(frmConnect.lstPlayers.ListIndex) <> vbNullString Then
                    frmConnect.UserName = frmConnect.lstPlayers.List(frmConnect.lstPlayers.ListIndex)
                    .WriteASCIIString frmConnect.UserName
                End If
                
                frmConnect.lstPlayers.Visible = False
                
            Case playerState.plyCreate
            
            Case playerState.plyKill
            
            Case playerState.plyExit
            
            Case playerState.plyNone: Exit Sub
            
        End Select
        
    End With
            
End Sub
Public Sub WriteCharEvents(bytepacket As Byte, charindex As Integer, chartype As characterType)

    With outgoingData 'Enviamos
        .WriteByte bytepacket 'tipo del paquete - accion -
        .WriteInteger charindex 'index del char
        .WriteByte chartype 'tipo del char - npc o player -
    
        Select Case bytepacket
            Case 1
                .WriteByte characterList(charindex).Heading 'no estoy seguro, pero asi deberia funcionar
                
        End Select
    End With
    
End Sub

Public Sub SendBuffer()
    
    With outgoingData
        If .Length = 0 Then Exit Sub
        
        Dim sndData As String
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        'No enviamos nada si no estamos conectados
        If frmMain.Winsock.State = sckConnected Then _
                frmMain.Winsock.SendData sndData
    
    End With

End Sub

