Attribute VB_Name = "modPlayer"
Option Explicit
Option Base 1

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

Public Const MAX_INVENTORY_SLOT   As Byte = 20
Public Const MAX_INVENTORY_OBJS   As Long = 50000

Private Enum eClass
    Mage = 1    'Mago
    Cleric      'Cl�rigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladr�n
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Palad�n
    Hunter      'Cazador
    Fisher      'Pescador
    Blacksmith  'Herrero
    Lumberjack  'Le�ador
    Miner       'Minero
    Carpenter   'Carpintero
    Pirat       'Pirata
End Enum

Private Enum eCity
    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
End Enum

Private Enum eRace
    Human = 1 'Humano
    Elf       'Elfo
    Drow      'Elfo Oscuro
    Gnome     'Gnomo
    Dwarf     'Enano
End Enum

Private Enum eGenere
    Male = 1
    Female
End Enum

Private Type playerBasicInfo
    Class       As eClass
    Race        As eRace
    Genere      As eGenere
    City        As eCity
End Type

Private Type playerStats

    Level       As Byte
    Exp         As Long
    Gold        As Long
    
    Algin       As characterAlgin
    
    maxHp       As Integer
    minHp       As Integer
    
    MaxMan      As Integer
    MinMan      As Integer
    
    MaxSta      As Integer
    MinSta      As Integer
    
    MaxHam      As Byte
    MinHam      As Byte
    
    MaxSed      As Byte
    MinSed      As Byte
    
    MinDef      As Integer
    MaxDef      As Integer
    
    NPCKill     As Integer
    UserKill    As Integer
    
    numSkills   As Byte     ' Skills libres
    Skills(21)  As Byte     ' 21 Skills
    
    Atr(5)      As Byte     ' 5 Atributes
    
    Spells(35)  As Integer  ' 35 Spells
    
End Type

Private Type playerFlags
    Dead       As Byte ' 1=true
    Hidden     As Byte
    Trading    As Byte
End Type

Public Type playerObj
    Index    As Integer
    Amount   As Long
    Equipped As Boolean
End Type

Public Type playerInventory
    Object(1 To MAX_INVENTORY_SLOT) As playerObj
    WeaponEqpIndex  As Integer
    WeaponEqpSlot   As Byte
    ArmorEqpIndex   As Integer
    ArmorEqpSlot    As Byte
    EscudoEqpIndex  As Integer
    EscudoEqpSlot   As Byte
    CascoEqpIndex   As Integer
    CascoEqpSlot    As Byte
        
    BarcoEqpIndex   As Integer
    BarcoEqpSlot    As Byte
    NroItems        As Integer
End Type

Public Type player
    Active         As Byte 'Indicates if the user is connected
    sckIndex       As Integer ' User Socket
    
    Name           As String
    Desc           As String
    
    Pos            As characterWorldPos
    Char           As characterAparence
    BasicInfo      As playerBasicInfo
    Stats          As playerStats
    flags          As playerFlags
    
    Inventory(20)  As playerInventory
    
   'Outgoing and incoming messages
    outgoingData   As clsByteQueue
    incomingData   As clsByteQueue
    
End Type: Public playerList() As player
Public Function playerOpen(ByRef sckIndex As Integer)
           
        With playerList(sckIndex)
        
            .Active = 1
            .sckIndex = sckIndex
        
            Set .incomingData = New clsByteQueue
            Set .outgoingData = New clsByteQueue
            
            'Make sure both outgoing and incoming data buffers are clean
            .incomingData.ReadASCIIStringFixed .incomingData.length
            .outgoingData.ReadASCIIStringFixed .outgoingData.length
        
        End With

End Function
Private Function playerExist(ByRef playerName As String) As Boolean
    
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `players` WHERE playername='" & playerName & "'")
    
    playerExist = (Not Recordset Is Nothing)
    
    Set Recordset = Nothing
    
End Function
Private Function playerBanCheck(ByRef playerName As String) As Boolean

    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `players` WHERE playername='" & playerName & "'")
    
    If Recordset.EOF Or Recordset.BOF = True Then
        playerBanCheck = False
        Set Recordset = Nothing
        Exit Function
    End If
    
    playerBanCheck = (Recordset!Ban = 1)

    Set Recordset = Nothing
    
End Function
Public Function playerConnect(ByRef sckIndex As Integer, ByRef playerName As String) As Boolean
    
    playerConnect = False
    
    If playerExist(playerName) = False Then
        sckClose sckIndex
        Exit Function
    End If
    
    If playerBanCheck(playerName) Then
        sckClose sckIndex
        Exit Function
    End If
    
    If playerList(sckIndex).Active = 1 Then
        sckClose sckIndex
        Exit Function
    End If
    
    'Get the player
    
    Set Recordset = Nothing
        
    Set Recordset = Connection.Execute("SELECT * FROM `players` WHERE playername='" & playerName & "'")
    
        socketList(sckIndex).Status = sOnPJ
        
        playerOpen sckIndex
    
            With playerList(sckIndex)
                .Name = playerName
                
                ' todo: this currently not works, implement ADODB Command
                ' http://www.timesheetsmts.com/adotutorial.htm
                With .Char
                
                    .Body = 1 'Recordset!Body
                    .Head = 1 'Recordset!Head
                    .Heading = 1 'Recordset!Heading
                
                End With
                
                With .Pos
                    
                    .Map = 1 'Recordset!Map
                    .X = 50 'Recordset!X
                    .Y = 50 'Recordset!Y
                
                End With
                
            End With
            
    Set Recordset = Nothing
    
    characterMake sckIndex, characterNextOpen, characterType.player
  
    playerConnect = True
    
End Function
Public Function playerCreate(ByRef playerName As String) As Boolean
    playerCreate = False
    
    If playerExist(playerName) = True Then
        'send message
        Exit Function
    End If

    Set Recordset = Connection.Execute("INSERT INTO `players` (playername,ban) " & _
                                    "values('" & playerName & "','0')")
     
     playerCreate = CBool(Not Recordset Is Nothing)
         
     Set Recordset = Nothing
End Function
Public Function playerKill(playerName As String) As Boolean
    
    If (playerExist(playerName) = False) Then Exit Function
    
    Set Recordset = Nothing
    
    ' Kill Player
    Set Recordset = Connection.Execute("DELETE FROM `players` WHERE playername='" & playerName & "'")

    Set Recordset = Nothing
    
    playerKill = True
End Function
Public Sub playerDisconnect(ByRef sckIndex As Integer)
    
    With playerList(sckIndex)
        .Active = 0

        Set .incomingData = Nothing
        Set .outgoingData = Nothing
        
        socketList(sckIndex).Status = sNone
        
        characterErase socketList(sckIndex).sckIndex
    End With
    
End Sub
