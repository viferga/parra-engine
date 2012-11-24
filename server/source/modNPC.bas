Attribute VB_Name = "modNPC"
Option Explicit


Private Type npcFlags
    Dead       As Byte ' 1=true
    Hidden     As Byte
    Trading    As Byte
End Type

Private Type npcStats
    
    Exp         As Long
    Gold        As Long
    
    Algin       As characterAlgin
    
    maxHp       As Integer
    minHp       As Integer
    
    MinDef      As Integer
    MaxDef      As Integer
    
End Type

Public Type Npc
    Index          As Long  ' Array Index
    
    Name           As String
    Desc           As String
    
    Pos            As characterWorldPos
    Char           As characterAparence
    Stats          As npcStats
    flags          As npcFlags
    
End Type: Public npcList() As Npc
