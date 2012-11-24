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

