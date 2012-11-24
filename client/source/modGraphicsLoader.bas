Attribute VB_Name = "modGraphicsLoader"
#If LoadingMetod = 1 Then

Option Explicit

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As DxVBLibA.Direct3DTexture8
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private D3DXLoad As DxVBLibA.D3DX8
Private d3Device As DxVBLibA.Direct3DDevice8

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Type tempTexture
    Surface As DxVBLibA.Direct3DTexture8
    Width As Integer
    Height As Integer
End Type

#Else

Option Explicit

'Manage Textures
Private Type structGraphic
        FileName   As Long
        D3DTexture As DxVBLibA.Direct3DTexture8
        Used       As Long
        Available  As Boolean
        Width      As Integer
        Height     As Integer
End Type: Public oGraphic() As structGraphic: Public lKeys() As Long

Private lSurfaceSize As Long
Private TexInfo      As D3DXIMAGE_INFO
Private Const lMaxEntrys As Integer = 1000

Private Type tCache
    Number        As Long
    SrcHeight     As Single
    SrcWidth      As Single
End Type: Private Cache As tCache

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal numBytes As Long)

#End If

#If LoadingMetod = 1 Then

Public Sub surfaceTerminate()
'**************************************************************
'Clean up
'**************************************************************
    Dim I As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For I = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(I)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next I
End Sub

Public Sub surfaceInitialize(ByVal graphicPath As String, ByRef D3DX As D3DX8, ByRef device8 As Direct3DDevice8, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Initializes the manager
'**************************************************************
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
        
    ResourcePath = graphicPath
    
    'GlobalMemoryStatus pUdtMemStatus
    'mFreeMemoryBytes = pUdtMemStatus.dwAvailPhys
    
    Set D3DXLoad = D3DX
    Set d3Device = device8
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Public Function getSurface(ByRef fileIndex As Long, ByRef surfaceWidth As Integer, ByRef surfaceHeight As Integer) As Direct3DTexture8
'**************************************************************
'Retrieves the requested texture
'**************************************************************
    Dim I As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For I = 1 To .surfaceCount
            If .SurfaceEntry(I).fileIndex = fileIndex Then
                .SurfaceEntry(I).lastAccess = GetTickCount
                Set getSurface = .SurfaceEntry(I).Surface
                Exit Function
            End If
        Next I
    End With
    
    'Not in memory, load it!
    Set getSurface = LoadSurface(fileIndex, surfaceWidth, surfaceHeight)
End Function

Private Function LoadSurface(ByRef fileIndex As Long, ByRef surfaceWidth As Integer, ByRef surfaceHeight As Integer) As Direct3DTexture8
'**************************************************************
'Loads the surface named fileIndex + ".png" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler
    
    Dim newSurface As SURFACE_ENTRY_DYN
    Dim TexInfo As D3DXIMAGE_INFO
    Dim surfDesc As D3DSURFACE_DESC
        
    With newSurface
    
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount
        
            Dim Data() As Byte, lBitMap As Long
            
            openPng Data(), lBitMap, fileIndex
                                
            'Set .Surface = D3DXLoad.CreateTextureFromFileInMemoryEx(d3Device, data(0), UBound(data()) + 1, D3DX_DEFAULT, _
                                                            D3DX_DEFAULT, 6, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                            D3DX_FILTER_DITHER Or D3DX_FILTER_TRIANGLE, D3DX_FILTER_BOX, _
                                                            D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
            
            Set .Surface = D3DXLoad.CreateTextureFromFileInMemory(d3Device, Data(0), lBitMap)
                
        
                
            'Set .Surface = Loader.CreateTextureFromFileEx(D3DDevice, ResourcePath & CStr(fileIndex) & ".png", D3DX_DEFAULT, _
                                                        D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                        D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
                
                
                
        .Surface.GetLevelDesc 0, surfDesc
                
        surfaceWidth = TexInfo.Width
        surfaceHeight = TexInfo.Height

    End With

    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + surfDesc.Size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function

ErrHandler:
    MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & CStr(fileIndex) & ".png" & ". " & vbCrLf & _
                "El código de error es " & CStr(Err.Number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
                vbOKOnly Or vbCritical Or vbExclamation, "Error"
    End
End Function

Private Function RemoveLRU() As Boolean
'**************************************************************
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************

    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim I As Long
    Dim j As Long
    Dim surfDesc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For I = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(I)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).lastAccess Then
                    LRUi = I
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).lastAccess
                End If
            Next j
        End With
    Next I
    
    'Retrieve the surface desc
    surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc 0, surfDesc
       
    'Remove it
    Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
    surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With surfaceList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
    
    'Update the used bytes
    usedBytes = usedBytes - surfDesc.Size
End Function

Private Function openPng(ByRef Data() As Byte, ByRef lBitMap As Long, ByVal fileIndex As Long)
    
    Dim Handle As Integer
    
    Handle = FreeFile() ' Get a free file number
    
        Open ResourcePath & CStr(fileIndex) & ".png" For Binary Access Read As #Handle
            
            lBitMap = LOF(Handle)
            
            ReDim Data(lBitMap - 1) As Byte  ' Create an array just big enough to hold the whole file
            
            Get #Handle, , Data() ' Read the file into that array
   
        Close #Handle

End Function

#Else

Public Function texInitialize() As Boolean
On Error GoTo errHandle

    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
        
    texInitialize = True
    Exit Function
errHandle:
    texInitialize = False
End Function
Private Function texDelete(ByRef FileNumber As Long) As Boolean
    
    lKeys(FileNumber) = 0

    ZeroMemory oGraphic(FileNumber), Len(oGraphic(FileNumber))
    lSurfaceSize = lSurfaceSize + 1
    
End Function
Public Function textureLoad(GrhIndex As Long) As Boolean
    
    If (Cache.Number <> GrhIndex) Then

        If oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture Is Nothing Then
        
            With Cache
                Set oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture = texLoad(Grh(GrhIndex).FileNum)
        
                .SrcHeight = oGraphic(lKeys(Grh(GrhIndex).FileNum)).Height + 1
                .SrcWidth = oGraphic(lKeys(Grh(GrhIndex).FileNum)).Width + 1
                
                Cache.Number = GrhIndex
            End With
        
        End If
    
    End If
    
    If oGraphic(lKeys(Grh(GrhIndex).FileNum)).D3DTexture Is Nothing Then
        textureLoad = False
    Else
        textureLoad = True
    End If

End Function
Private Function texLoad(ByRef FileNumber As Long) As Direct3DTexture8

    oGraphic(lKeys(FileNumber)).Used = oGraphic(lKeys(FileNumber)).Used + 1

    If (oGraphic(lKeys(FileNumber)).Available = False) Then
        If (texCreateFromFile(FileNumber) = False) Then
            Set texLoad = Nothing: Exit Function
        Else
            lSurfaceSize = lSurfaceSize - 1
        End If
    End If

    Set texLoad = oGraphic(lKeys(FileNumber)).D3DTexture

End Function
Private Function texCreateFromFile(ByRef FileNumber As Long) As Boolean
    Dim I As Long
    Dim TexNum As Long
    Dim DelTex As Long
    Dim TexInfo As D3DXIMAGE_INFO
    
        TexNum = 0
    
        For I = 1 To lMaxEntrys
            If (oGraphic(I).Available = False) Then
                TexNum = I
                oGraphic(I).Available = True
                Exit For
            Else
                If (oGraphic(I).Used < 0) Then oGraphic(I).Used = 0: DelTex = I
            End If
        Next I
    
        If (TexNum = 0) Then
            If (texDelete(DelTex) = False) Then
                texCreateFromFile = False: Exit Function
            Else
                lKeys(FileNumber) = DelTex
            End If
        Else
            lKeys(FileNumber) = TexNum
        End If
        
        
    Dim Handle As Integer

    Handle = FreeFile() ' Get a free file number

    Open App.Path & "\Graphics\" & CStr(FileNumber) & ".png" For Binary Access Read As #Handle
        Dim FileData() As Byte

        ReDim FileData(LOF(Handle) - 1) As Byte  ' Create an array just big enough to hold the whole file

        Get #Handle, , FileData() ' Read the file into that array

        Set oGraphic(lKeys(FileNumber)).D3DTexture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, FileData(0), UBound(FileData()) + 1, D3DX_DEFAULT, _
                                                D3DX_DEFAULT, 6, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                D3DX_FILTER_DITHER Or D3DX_FILTER_TRIANGLE, D3DX_FILTER_BOX, _
                                                D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)

    Close #Handle
    
    
    ' Create the Texture using the filedata() array
    'Set oGraphic(lKeys(FileNumber)).D3DTexture = Loader.CreateTextureFromFileInMemory(D3DDevice, FileData(0), lBitmap)
    
    'Set oGraphic(lKeys(FileNumber)).D3DTexture = Loader.CreateTextureFromFileEx(D3DDevice, App.Path & "\Graphics\" & CStr(FileNumber) & ".png", D3DX_DEFAULT, _
                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                D3DX_FILTER_POINT, D3DX_FILTER_POINT, D3DColorXRGB(0, 0, 0), TexInfo, ByVal 0)
    
    
    With oGraphic(lKeys(FileNumber))
        .Width = TexInfo.Width
        .Height = TexInfo.Height
    End With

    texCreateFromFile = True

End Function
Public Sub texReloadAll()
    Dim I As Long

    For I = 1 To lSurfaceSize
        With oGraphic(I)
            If (.Available = True) Then
                Set .D3DTexture = Nothing
                    .Available = False
            End If
        End With
    Next I
    
    lSurfaceSize = lMaxEntrys

    ReDim oGraphic(lMaxEntrys)
    ReDim lKeys(1 To lMaxEntrys)
    
End Sub
Public Sub texDestroyAll()
    Dim I As Long

    For I = 1 To lSurfaceSize
        With oGraphic(I)
            If (.Available = True) Then
                Set .D3DTexture = Nothing
                    .Available = False
            End If
        End With
    Next I
    
    Erase oGraphic
    Erase lKeys
End Sub

#End If
