Attribute VB_Name = "modMemory"
Option Explicit

' Primitive type size
Public Const byteSize As Long = 1
Public Const intSize  As Long = 2
Public Const lngSize  As Long = 4
Public Const sngSize  As Long = 4
Public Const dblSize  As Long = 8

' API declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                            (Destination As Any, Source As Any, ByVal Length As Long)
                              
Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
                            (ByVal hHeap As Long, ByVal dwFlags As Long, _
                             ByVal dwBytes As Long) As Long
                             
Private Declare Function HeapFree Lib "kernel32" _
                            (ByVal hHeap As Long, ByVal dwFlags As Long, _
                            lpMem As Any) As Long

Private Declare Sub MemoryZero Lib "kernel32.dll" Alias "RtlZeroMemory" _
                            (Destination As Any, ByVal Length As Long)

Public Declare Sub MemoryCopy Lib "kernel32" Alias _
                            "RtlMoveMemory" (Destination As Any, _
                            Source As Any, ByVal Length As Long)


Public Declare Sub MemoryWrite Lib "kernel32" Alias _
                            "RtlMoveMemory" (ByVal Destination As Long, _
                            Source As Any, ByVal Length As Long)

Public Declare Sub MemoryRead Lib "kernel32" Alias _
                            "RtlMoveMemory" (Destination As Any, _
                            ByVal Source As Long, ByVal Length As Long)

' Methods
Public Function strSize(ByRef Value As String) As Long
    ' Length prefix + Datastring (Unicode) + Terminator
    strSize = (4 + (Len(Value) * 2) + 2)
End Function
Public Function MemoryAllocate(ByVal Size As Long) As Long
    Dim hHeap As Long
    
    If Size > 0 Then
        hHeap = GetProcessHeap()
        MemoryAllocate = HeapAlloc(hHeap, 0, Size)
        Exit Function
    End If
    
    MemoryAllocate = 0
End Function
Public Sub MemoryDeallocate(ByRef Pointer As Long)
    Dim hHeap As Long
    
    If Pointer <> 0 Then
        hHeap = GetProcessHeap()
        HeapFree hHeap, 0, Pointer
        Pointer = 0
    End If
End Sub
Public Function MemoryProcAddress(ByVal lngAddressOf As Long) As Long
  MemoryProcAddress = lngAddressOf
End Function
