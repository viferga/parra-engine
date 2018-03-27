Attribute VB_Name = "modCharacter"
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

Public Type characterWorldPos
    Map          As Integer
    X            As Byte
    Y            As Byte
End Type

Public Type characterAparence
    Head        As Integer
    Body        As Integer
    
    Weapon      As Integer
    Shield      As Integer
    Helmet      As Integer
    
    FX          As Integer
    Loops       As Integer
    
    Heading     As Byte
End Type

Public Enum characterAlgin
    Neutral = 1
    Criminal
    Civil
End Enum

Public Enum characterDirection
    NorthWeast = 1
    NorthEast = 2
    SouthWeast = 3
    SouthEast = 4
    North = 5
    South = 6
    East = 7
    West = 8
End Enum

Public Enum characterType
    player = 0
    Npc = 1
End Enum

Private Type characterInfo
    ID As Integer
    Type As characterType
End Type

Public characterList() As characterInfo
Private charMax As Integer, charLast As Integer, charSize As Integer
Public Function characterInitialize() As Boolean
    
    charMax = 5000
    
    ReDim Preserve characterList(1 To charMax) As characterInfo
    ReDim Preserve playerList(1 To sckMax) As player
    ReDim Preserve npcList(1 To charMax - sckMax) As Npc
    
    charLast = 1
    charSize = 1
    
    characterInitialize = True
    
End Function
Public Sub characterEraseAll()
    
    Erase characterList()
    Erase npcList()
    Erase playerList()
    
End Sub
Public Function characterNextOpen() As Integer
    
    Dim i As Long
    
        For i = charLast To charMax
            If characterList(i).ID = 0 Then
                characterNextOpen = i
                Exit Function
            End If
        Next i
        
End Function
Public Sub characterErase(ByRef CharIndex As Integer)

    HandleRemoveChar characterList(CharIndex).ID
    
    DoEvents
    
    characterList(CharIndex).ID = 0
    
    If (CharIndex = charLast) Then
        Do Until characterList(charLast).ID > 0
           If (charLast <= 1) Then Exit Do
           charLast = charLast - 1
        Loop
    End If
    
End Sub

Public Function characterMove(ByRef CharIndex As Integer, ByRef charType As characterType, Direction As characterDirection)
    
    
    Select Case charType
    
        Case characterType.player

            characterMoveDirection playerList(characterList(CharIndex).ID).Pos, Direction
                
        Case characterType.Npc
          
            characterMoveDirection npcList(characterList(CharIndex).ID).Pos, Direction

    End Select
    
End Function
Public Function characterPosition(ByRef CharIndex As Integer) As characterWorldPos
    
    Select Case characterList(CharIndex).Type
        Case characterType.player
            characterPosition = playerList(characterList(CharIndex).ID).Pos
            
        Case characterType.Npc
            characterPosition = npcList(characterList(CharIndex).ID).Pos
            
    End Select
    
End Function
Private Function characterMoveDirection(ByRef Position As characterWorldPos, Direction As characterDirection)

    With Position

        Select Case Direction

                Case characterDirection.NorthWeast
                    .X = .X - 1: .Y = .Y + 1
                Case characterDirection.NorthEast
                    .X = .X + 1: .Y = .Y + 1
                Case characterDirection.SouthWeast
                    .X = .X - 1: .Y = .Y - 1
                Case characterDirection.SouthEast
                    .X = .X + 1: .Y = .Y - 1

        End Select
    
    End With

End Function
Public Function characterMake(ByRef CharIndex As Integer, charID As Integer, charType As characterType)

    Dim charPos As characterWorldPos
    
        With characterList(CharIndex)
                .ID = charID
                .Type = charType
        End With
            
        charPos = characterPosition(CharIndex)
        
        Select Case charType
        
            Case characterType.player
            
                    With playerList(characterList(CharIndex).ID)
                    
                       HandleCreateChar CharIndex, .Char, .Pos, .Name
                
                    End With
                  
                 Exit Function
                 
            Case characterType.Npc
            
                    With npcList(characterList(CharIndex).ID)
                    
                        HandleCreateChar CharIndex, .Char, .Pos
                    
                    End With
                    
                 Exit Function
                 
        End Select
                
End Function


