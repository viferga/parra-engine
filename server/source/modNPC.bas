Attribute VB_Name = "modNPC"
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
