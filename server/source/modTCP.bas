Attribute VB_Name = "modTCP"
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

Public Enum sckStatus
        sNone = 0
        sOnAcc = 1
        sOnPJ = 2
End Enum

'Socket Struct
Public Type sckStruct
        IP        As String
        sckIndex  As Integer
        Status    As sckStatus
        Latency   As Integer
        
        'Auxiliar buffers
        incomingData As clsByteQueue
        outgoingData As clsByteQueue
End Type: Public socketList() As sckStruct

'Internal Socket
Public sckSize As Integer, sckMax As Integer

Public Enum SendTarget
        toNone = 0
        toSocket
        toCharacter
        toArea
        toAreaButCharacter
        toMap
        toMapButCharacter
        toAll
        toAllCharacter
        toAllButCharacter
        toAdmin
End Enum
Public Function sockInitialize(ByVal Size As Integer, ByVal Port As Integer) As Boolean
On Error GoTo errSock

    Consola "Loading Sockets..."
    
    sckMax = Size
    
    ReDim socketList(1 To sckMax)
    
    frmMain.wsEsc.LocalPort = Port
    frmMain.wsEsc.Listen
    
    Dim i As Long
    
    For i = 1 To sckMax - 1
        Load frmMain.Winsock(i)
    Next i
        
    sockInitialize = True
    Exit Function
    
errSock:
    Debug.Print "Critical error when initializing Sockets"
    Consola "���Server error when trying to establish connection!!!"
    
    sockInitialize = False
End Function
Public Function sockDeInitialize() As Boolean
On Error GoTo errSock

    Dim i As Long
    
    frmMain.wsEsc.Close
        
    For i = 1 To sckSize - 1
        frmMain.Winsock(i).Close
        Unload frmMain.Winsock(i)
    Next

    sockDeInitialize = True
    Exit Function
errSock:
    Debug.Print "Critical error when deinitializing Sockets"
    
    sockDeInitialize = False
End Function
Public Sub sckSendEx(ByRef sckIndex As Integer, ByRef sndRoute As SendTarget, ByRef Data As String)
    
            Select Case sndRoute
                   Case toNone
                        Exit Sub
                   Case toSocket
                        sckSend sckIndex, Data
                        Exit Sub
                        
                   Case Else: Exit Sub
            End Select
                      
End Sub
Public Sub SendBufferSocket(ByRef sckIndex As Integer, outgoingData As clsByteQueue)

    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        sckSend sckIndex, sndData
    End With
        
End Sub
Public Sub sckSend(ByRef sckIndex As Integer, ByRef sndData As String)

    Dim Data() As Byte
    
    ReDim Preserve Data(Len(sndData) - 1) As Byte
    
    Data = StrConv(sndData, vbFromUnicode)
    
On Local Error GoTo sckErr
    
    'SendData
    If frmMain.Winsock(sckIndex).State <> sckClosed And frmMain.Winsock(sckIndex).State <> sckClosing Then
        frmMain.Winsock(sckIndex).SendData Data
    End If

    Exit Sub
sckErr:
    Debug.Print Err.Description; "WState: " & frmMain.Winsock(sckIndex).State
End Sub
Public Sub sckDataArrival(sckIndex As Integer, bytesTotal As Long)

    Dim bufferData() As Byte
    
    If bytesTotal <= 4096 Then 'Limit Long
        ReDim Preserve bufferData(bytesTotal) As Byte
    Else
        Exit Sub
    End If

    frmMain.Winsock(sckIndex).GetData bufferData(), vbByte, bytesTotal

    socketList(sckIndex).incomingData.WriteBlock bufferData, bytesTotal
                                            
    HandleIncomingData sckIndex

End Sub
Public Function sckOpen(ByRef sckIndex As Integer, ByRef requestID As Long)
        
        frmMain.Winsock(sckIndex).Close
        frmMain.Winsock(sckIndex).Accept requestID
                        
        With socketList(sckIndex)
        
            .IP = frmMain.Winsock(sckIndex).RemoteHostIP  'User IP
            .sckIndex = sckIndex
            .Status = sckStatus.sNone
            
            Consola "Connection Accept {sckIndex: " & CStr(.sckIndex) & " - IP: " & .IP & "}"
            
            Set .incomingData = New clsByteQueue
            Set .outgoingData = New clsByteQueue
            
            'Make sure both outgoing and incoming data buffers are clean
            .incomingData.ReadASCIIStringFixed .incomingData.length
            .outgoingData.ReadASCIIStringFixed .outgoingData.length
            
        End With
        
        sckSize = sckSize + 1
        frmMain.lblOnline.Caption = "Conexiones: " & sckSize

End Function
Public Sub sckClose(ByRef sckIndex As Integer)
    
    With socketList(sckIndex)

            Select Case .Status
                Case sckStatus.sOnAcc
                    'accDisconnect sckIndex

                Case sckStatus.sOnPJ
                    If playerList(sckIndex).Active = 1 Then
                        playerDisconnect sckIndex
                    End If
                    
            End Select
            
            Consola "Connection Close {sckIndex: " & CStr(.sckIndex) & " - IP: " & .IP & "}"
                                                                                                  
            .IP = vbNullString
            .Status = sckStatus.sNone
            .sckIndex = 0
            .Latency = 0
            
            Set .incomingData = Nothing
            Set .outgoingData = Nothing
                    
            If frmMain.Winsock(sckIndex).State <> sckClosed Or frmMain.Winsock(sckIndex).State <> sckClosing Then
                frmMain.Winsock(sckIndex).Close
            End If
                    
            Unload frmMain.Winsock(sckIndex)
            Load frmMain.Winsock(sckIndex)
        
    End With
   
    sckSize = sckSize - 1
    frmMain.lblOnline.Caption = "Connections: " & sckSize
    
End Sub
