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
Option Explicit

' Constants
Private Const STACK_SIZE_DEFAULT    As Long = 64
Private Const STACK_EMPTY           As Long = -1

' Member Data
Private ptrData()                   As Long
Private lngTop                      As Long
Private lngSize                     As Long

' Constructor
Private Sub Class_Initialize()
    lngSize = 0
    Create
End Sub

' Destructor
Private Sub Class_Terminate()
    Clear
End Sub

' Properties
Public Property Get IsFull() As Boolean
    IsFull = ((lngTop + 1) = lngSize)
End Property
Public Property Get IsEmpty() As Boolean
    IsEmpty = (lngTop = STACK_EMPTY)
End Property

' Methods
Public Sub Push(ByVal Value As Long)
    If IsFull() Then
        Resize
    End If
    
    lngTop = lngTop + 1
    ptrData(lngTop) = Value
End Sub
Public Function Pop() As Long
    Dim Value As Long
    
    Value = 0&

    If IsEmpty() = False Then
        Value = ptrData(lngTop)
        lngTop = lngTop - 1
    End If
            
    Pop = Value
End Function
Public Sub Resize(Optional ByVal Size As Long = 0)
    If Size > lngSize Then
        lngSize = Size
    ElseIf Size = 0 Then
        lngSize = lngSize * 2
    End If
    
    ReDim Preserve ptrData(lngSize) As Long
End Sub
Public Sub Create(Optional ByVal Size As Long = STACK_SIZE_DEFAULT)
    If lngSize > 0 Then
        Clear
    End If
    
    lngTop = STACK_EMPTY
    lngSize = Size
    
    ReDim ptrData(lngSize) As Long
End Sub
Public Sub Clear()
    If lngSize > 0 Then
        Erase ptrData
        ptrData = Null
    End If
    
    lngTop = STACK_EMPTY
    lngSize = 0
End Sub
