VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Fifo

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

Private stackMax  As Long
Private stack()   As Long
Private stackSize As Long

Public Property Get IsEmpty() As Boolean
    IsEmpty = (stackSize = 0)
End Property
Private Sub class_Terminate()
  Erase stack
End Sub
Public Sub Push(ByVal Value As Long)

  If (stackSize < stackMax) Then
      
      stackSize = stackSize + 1
      
      ReDim Preserve stack(stackSize)
    
      stack(stackSize) = Value
  
  End If
  
End Sub
Public Function Pop() As Long
  Dim X As Long
  
  If (stackSize > 0) Then
       Pop = stack(1)
       stackSize = stackSize - 1
       
       For X = 1 To stackSize
          stack(X) = stack(X + 1)
       Next X
       
       ReDim Preserve stack(stackSize)
       
  End If
  
End Function
Public Sub AddNumberRange(ByVal Min As Integer, ByVal max As Integer)
  Dim i As Integer
  
  stackMax = max
  stackSize = max
  
  ReDim stack(max)
  
  For i = 1 To stackSize
      stack(i) = (i - 1) + Min
  Next i
  
End Sub
Public Sub initialize(ByVal Size As Integer)
  
  stackSize = 0
  stackMax = Size
  ReDim stack(Size)
  
End Sub

