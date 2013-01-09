VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Member Data
Private strName        As String * 128
Private strFileName    As String * 256
Private lngHandle      As Long
Private ptrData        As Long
Private lngRefCount    As Long

' Constructor
Private Sub Class_Initialize()
    Create
End Sub

' Destructor
Private Sub Class_Terminate()
    Clear
End Sub

' Properties
Public Property Get Name() As String
    Name = strName
End Property
Public Property Let Name(ByRef Value As String)
    strName = Value
End Property
Public Property Get FileName() As String
    FileName = strFileName
End Property
Public Property Let FileName(ByRef Value As String)
    strFileName = Value
End Property
Public Property Get Handle() As Long
    Handle = lngHandle
End Property
Public Property Let Handle(ByVal Value As Long)
    lngHandle = Value
End Property
Public Property Get Data() As Long
    Data = ptrData
End Property
Public Property Let Data(ByVal Pointer As Long)
    ptrData = Pointer
End Property
Public Property Get RefCount() As Long
    RefCount = lngRefCount
End Property

' Modificators
Public Sub AddRef()
    lngRefCount = lngRefCount + 1
End Sub

Public Sub IncRef()
    lngRefCount = lngRefCount + 1
End Sub

' Methods
Public Function Create(Optional ByVal Size As Long = 0, Optional ByVal Pointer As Long = 0, _
                       Optional ByVal Handle As Long = 0, Optional ByRef Name As String = "", _
                       Optional ByRef Path = "") As Boolean
    lngHandle = Handle
    strName = Name
    strFileName = Path
    lngRefCount = 1
    
    If Size > 0 Then
        ptrData = MemoryAllocate(Size)
        
        If Pointer <> 0 Then
            MemoryWrite ptrData, ByVal Pointer, Size
        End If
    Else
        ptrData = Pointer
    End If
    
    Create = True
End Function

Public Sub Clear()
    strName = ""
    strFileName = ""
    lngHandle = 0
    lngRefCount = 0
    
    MemoryDeallocate ptrData
End Sub
