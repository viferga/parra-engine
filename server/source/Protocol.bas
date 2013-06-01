Attribute VB_Name = "modProtocol"
Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 245

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue

Public Enum ServerPacketID
    AccountEvents = 1
    UserEvents
    HandleMessage
    CharEvents
End Enum

Public Enum ClientPacketID
    AccountEvents = 1
    UserEvents
    CharEvents
End Enum

Public Enum MessageType
    MessageBox = 0
    Console
End Enum
Public Sub HandleIncomingData(ByRef sckIndex As Integer)
On Error Resume Next
    Dim packetID As Byte
    
    packetID = socketList(sckIndex).incomingData.PeekByte()
    
    'If packet isn't found in the enum
    If packetID > LAST_CLIENT_PACKET_ID Then
        sckClose sckIndex
    End If
    
    Select Case packetID
        Case ClientPacketID.AccountEvents
              HandleIncomingAccount sckIndex
              Exit Sub
              
        Case ClientPacketID.UserEvents
              HandleIncomingUser sckIndex
              Exit Sub
              
        Case ClientPacketID.CharEvents
              HandleIncomingChar sckIndex
              Exit Sub
            
        Case Else
            sckClose sckIndex
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If socketList(sckIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        HandleIncomingData sckIndex

    ElseIf Err.Number <> 0 And Not Err.Number = socketList(sckIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Debug.Print ("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - User Socket: " & sckIndex & " - producido al manejar el paquete: " & CStr(packetID))
        sckClose sckIndex
        
    End If
End Sub
Private Sub HandleIncomingAccount(ByRef sckIndex As Integer)

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(socketList(sckIndex).incomingData)
    
    'Remove packet ID
    buffer.ReadByte

    'Leemos paquete identificador
    
    Dim bytePacket As Byte
    Dim accountName As String, accountPassword As String, accountMail As String
    
    bytePacket = buffer.ReadByte()
    
        Select Case bytePacket
        
            Case 1
               accountName = buffer.ReadASCIIString()
               accountPassword = buffer.ReadASCIIString()
               
               If accountConnect(accountName, accountPassword) = False Then
                    HandleSendMessage sckIndex, MessageType.MessageBox, "Error: no se ha podido conectar la cuenta."
                    DoEvents
                    sckClose sckIndex
               Else
                    socketList(sckIndex).Status = sOnAcc
               End If
            
            Case 2
               accountName = buffer.ReadASCIIString()
               accountPassword = buffer.ReadASCIIString()
               accountMail = buffer.ReadASCIIString()
               
               If accountCreate(accountName, accountPassword, accountMail) = False Then
                    HandleSendMessage sckIndex, MessageType.MessageBox, "Error: no se ha podido crear la cuenta."
               End If
               
            Case 3
               accountName = buffer.ReadASCIIString()
               accountPassword = buffer.ReadASCIIString()
               accountMail = buffer.ReadASCIIString()
               
               If accountKill(accountName, accountPassword, accountMail) = False Then
                    HandleSendMessage sckIndex, MessageType.MessageBox, "Error: no se ha podido borrar la cuenta."
               End If
               
            Case 4
                'account exit
            
            Case Else: Exit Sub
            
        End Select

    
    'If we got here then packet is complete, copy data back to original queue
    Call socketList(sckIndex).incomingData.CopyBuffer(buffer)
    
    'Write Log in Client
    Call socketList(sckIndex).outgoingData.WriteByte(ServerPacketID.AccountEvents)
    
    'Enviamos paquete identificador
    Call socketList(sckIndex).outgoingData.WriteByte(bytePacket)
    
    If bytePacket = 1 Then
        If accountSendInfo(socketList(sckIndex).outgoingData) = False Then
            HandleSendMessage sckIndex, MessageType.MessageBox, "Error: no se ha podido enviar la información de la cuenta."
        End If
    End If
    
    SendBufferSocket sckIndex, socketList(sckIndex).outgoingData
    
    DoEvents
    
    If bytePacket = 2 Or bytePacket = 3 Or bytePacket = 4 Then
        sckClose sckIndex
    End If
    
End Sub
Private Sub HandleIncomingUser(ByRef sckIndex As Integer)

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(socketList(sckIndex).incomingData)
    
    'Remove packet ID
    buffer.ReadByte
    
    'Leemos paquete identificador
    
    Dim bytePacket As Byte
    
        bytePacket = buffer.ReadByte()
    
        Select Case bytePacket
            Case 1 'login

               Dim playerName As String
    
               playerName = buffer.ReadASCIIString()
    
               If playerConnect(sckIndex, playerName) = False Then
                    HandleSendMessage sckIndex, MessageType.MessageBox, "Error: el personaje no existe."
                    sckClose sckIndex
                    Exit Sub
               End If
            
            Case 2 'create
            
            
            Case 3 'kill
            
            
            Case 4 'exit
               
               playerDisconnect sckIndex
            
        End Select
    
    'If we got here then packet is complete, copy data back to original queue
    Call socketList(sckIndex).incomingData.CopyBuffer(buffer)
    
    'Write Log in Client
    Call socketList(sckIndex).outgoingData.WriteByte(ServerPacketID.UserEvents)
    
    'Aqui enviamos paquete identificador
    Call socketList(sckIndex).outgoingData.WriteByte(bytePacket)
    
    SendBufferSocket sckIndex, socketList(sckIndex).outgoingData
    
    DoEvents
    
End Sub
Private Sub HandleIncomingChar(ByRef sckIndex As Integer)
'GDK: no veo ningun string por aca

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(socketList(sckIndex).incomingData)
    
    'Remove packet ID
    buffer.ReadByte
    
    Dim bytePacket As Byte
    Dim CharIndex As Integer
    Dim charType As characterType
    
    'Leemos paquete identificador
    bytePacket = buffer.ReadByte()
        
    CharIndex = buffer.ReadInteger()
    charType = buffer.ReadByte()
        
    
   '     CharIndex = buffer.ReadInteger() 2 veces xD GDK
    
        Select Case bytePacket
            Case 1 'charactermove
                Dim Direction As characterDirection
                
                Direction = buffer.ReadByte()
                characterMove CharIndex, charType, Direction
                
            Case 2
            
            
            Case 3
            
            
            Case 4

            
        End Select
    
    'If we got here then packet is complete, copy data back to original queue
    Call socketList(sckIndex).incomingData.CopyBuffer(buffer)
        
End Sub

Public Sub HandleSendMessage(ByRef sckIndex As Integer, ByRef MsgType As MessageType, ByRef Message As String)
    
    With socketList(sckIndex)
   
        With .outgoingData
   
            .WriteByte ServerPacketID.HandleMessage
   
            .WriteByte MsgType
   
            .WriteASCIIString Message
   
        End With
   
        SendBufferSocket sckIndex, .outgoingData
   
    End With
        
End Sub
Public Sub HandleCreateChar(ByRef sckIndex As Integer, Char As characterAparence, Position As characterWorldPos, Optional ByRef playerName As String = vbNullString)

    With socketList(sckIndex)
   
        With .outgoingData
   
            .WriteByte ServerPacketID.CharEvents
            .WriteByte 1 'packet create
            .WriteInteger sckIndex
            .WriteASCIIString playerName
            .WriteInteger Char.Body
            .WriteInteger Char.Head
            .WriteByte Char.Heading
            .WriteByte Position.X
            .WriteByte Position.Y
            
        End With
   
        SendBufferSocket sckIndex, .outgoingData
   
    End With
        
End Sub
Public Sub HandleRemoveChar(ByRef sckIndex As Integer)
    If sckIndex <= 0 Then
        Debug.Print "Something wrong happens here!"
        Exit Sub
    End If

    With socketList(sckIndex)
   
        With .outgoingData
   
            .WriteByte ServerPacketID.CharEvents
            .WriteByte 2 'packet remove
            .WriteInteger sckIndex
            
        End With
   
        SendBufferSocket sckIndex, .outgoingData
   
    End With
        
End Sub
