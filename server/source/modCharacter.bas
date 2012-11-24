Attribute VB_Name = "modCharacter"
Option Explicit

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
Public Sub characterErase(ByRef charIndex As Integer)

    HandleRemoveChar characterList(charIndex).ID
    
    DoEvents
    
    characterList(charIndex).ID = 0
    
    If (charIndex = charLast) Then
        Do Until characterList(charLast).ID > 0
           If (charLast <= 1) Then Exit Do
           charLast = charLast - 1
        Loop
    End If
    
End Sub

Public Function characterMove(ByRef charIndex As Integer, ByRef charType As characterType, Direction As characterDirection)
    
    
    Select Case charType
    
        Case characterType.player

            characterMoveDirection playerList(characterList(charIndex).ID).Pos, Direction
                
        Case characterType.Npc
          
            characterMoveDirection npcList(characterList(charIndex).ID).Pos, Direction

    End Select
    
End Function
Public Function characterPosition(ByRef charIndex As Integer) As characterWorldPos
    
    Select Case characterList(charIndex).Type
        Case characterType.player
            characterPosition = playerList(characterList(charIndex).ID).Pos
            
        Case characterType.Npc
            characterPosition = npcList(characterList(charIndex).ID).Pos
            
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
Public Function characterMake(ByRef charIndex As Integer, charID As Integer, charType As characterType)

    Dim charPos As characterWorldPos
    
        With characterList(charIndex)
                .ID = charID
                .Type = charType
        End With
            
        charPos = characterPosition(charIndex)
        
        Select Case charType
        
            Case characterType.player
            
                    With playerList(characterList(charIndex).ID)
                    
                       HandleCreateChar charIndex, .Char, .Pos, .Name
                
                    End With
                  
                 Exit Function
                 
            Case characterType.Npc
            
                    With npcList(characterList(charIndex).ID)
                    
                        HandleCreateChar charIndex, .Char, .Pos
                    
                    End With
                    
                 Exit Function
                 
        End Select
                
End Function


