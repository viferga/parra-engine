VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAllocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

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


' Allocator Class

'**************************************************************
' clsAllocator.cls - Allocator class for managing memory.
' Uses the current process heap to manipulate directly memory.
'
' @author Designed and implemented by �Parra
' @version 1.0.0
' @date 20130417
'**************************************************************

' Definitions
Private Const HEAP_ZERO_MEMORY As Long = &H8

Private Enum AllocEnum
    AllocZero
    AllocPreserve
    AllocDirty
End Enum

' Declarations
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapCompact Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Private Declare Function HeapWalk Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapValidate Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapSetInformation Lib "kernel32" (ByVal hHeap As Long, ByVal dwInformationClass As Long, pInformation As Any, ByVal dwLength As Long) As Long
Private Declare Function HeapQueryInformation Lib "kernel32" (ByVal hHeap As Long, ByVal dwInformationClass As Long, pInformation As Any, ByVal dwLength As Long, pSize As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Member Data
Private ptrData As Long
Private lngDataSize As Long
Private hHeap As Long

' Constructor
Private Sub Class_Initialize()
    hHeap = GetProcessHeap()
    lngDataSize = 0
    ptrData = 0
End Sub

' Destructor
Private Sub Class_Terminate()
    Clear
    hHeap = 0
End Sub

' Properties
Public Property Let Size(ByVal lngSize As Long)
    ReAllocate lngSize
End Property
Public Property Get Address() As Long
    Address = ptrData
End Property
Public Property Get Size() As Long
    Size = lngDataSize
End Property

' Functions
Private Function ReAllocate(ByVal lngSize As Long, Optional ByVal AllocType As AllocEnum = AllocPreserve) As Boolean
    If (ptrData <> 0 And lngDataSize > 0) Then
        Dim ptrDataOld As Long
        Dim lngSizeOld As Long
        
        ptrDataOld = ptrData
        lngSizeOld = lngDataSize
        
        If (AllocType = AllocZero) Then
            ptrData = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, lngSize)
        Else
            ptrData = HeapAlloc(hHeap, 0, lngSize)
        End If
        
        lngDataSize = lngSize
        
        If (AllocType = AllocPreserve) Then
            If (lngSizeOld < lngSize) Then
                lngSize = lngSizeOld
            End If
        
            CopyMemory ptrData, ptrDataOld, lngSize
        End If
        
        HeapFree hHeap, 0, ptrDataOld
    Else
        ptrData = HeapAlloc(hHeap, 0, lngSize)
        lngDataSize = lngSize
    End If
    
    ReAllocate = (ptrData <> 0 And lngDataSize > 0)
End Function
Public Function Allocate(ByVal lngSize As Long) As Boolean
    If (lngSize > 0) Then
        Allocate = ReAllocate(lngDataSize + lngSize)
        Exit Function
    End If
    
    Allocate = False
End Function
Public Function Deallocate(ByVal lngSize As Long) As Boolean
    If (lngDataSize >= lngSize) Then
        Deallocate = ReAllocate(lngDataSize - lngSize)
        Exit Function
    End If
    
    Deallocate = False
End Function

' Procedures
Private Sub ShiftLeft(ByVal lngSize As Long)
    If lngDataSize >= lngSize Then
        CopyMemory ByVal ptrData, ByVal ptrData + lngSize, lngDataSize - lngSize
        Deallocate lngSize
    End If
End Sub
Public Sub Clear()
    If (ptrData <> 0 And lngDataSize > 0) Then
        HeapFree hHeap, 0, ptrData
        lngDataSize = 0
        ptrData = 0
    End If
End Sub
Public Sub GetData(ptr As Long)
    Peek ptr
    Clear
End Sub
Public Sub Peek(ptr As Long)
    If (lngDataSize > 0) Then
        CopyMemory ByVal ptr, ByVal ptrData, lngDataSize
    End If
End Sub
Public Sub Append(ptr As Long, ByVal lngSize As Long)
    Allocate lngSize
    SetData ptr, lngSize, lngDataSize - lngSize
End Sub
Public Sub SetData(ptr As Long, ByVal lngSize As Long, Optional ByVal lngOffset As Long = 0)
    If lngDataSize >= lngSize Then
        CopyMemory ByVal ptrData + lngOffset, ByVal ptr, lngSize
    End If
End Sub
Public Sub GetDataEx(ptr As Long, ByVal lngSize As Long)
    If lngDataSize >= lngSize Then
        PeekEx ptr, lngSize
        ShiftLeft lngSize
    End If
End Sub
Public Sub PeekEx(ptr As Long, ByVal lngSize As Long, Optional ByVal lngOffset As Long = 0)
    If (lngDataSize >= lngSize) Then
        CopyMemory ByVal ptr, ByVal ptrData + lngOffset, lngSize
    End If
End Sub

