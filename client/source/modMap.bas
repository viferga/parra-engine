Attribute VB_Name = "modMap"
Option Explicit

Public EditMap As Boolean

Public Const MaxTilesX As Integer = 100
Public Const MaxTilesY As Integer = 100

Public Const TilePixelWidth As Integer = 64
Public Const TilePixelHeight As Integer = 64

Public Const HalfTileWidth = TilePixelWidth / 2
Public Const HalfTileHeight = TilePixelHeight / 2

' Position tile mouse on map
Public MouseTilesPos As structPositionSng
Public MousePosOnMap As structPositionInt

Public Type mapBlock

    Layer(1 To 4) As structGrh
        
    CharIndex As Integer
    
    Blocked   As Byte
    Trigger   As Byte
    
    TileExit  As structPosByte
    
    particleIndex As Integer
    
    LightColor(3) As Long
End Type

Public mapData() As mapBlock

Public mapPreCalcPos(1 To 100, 1 To 100) As structPositionLng 'Posicion predefinida

Public Sub mapLoadAll()

    Dim maxMaps As Long
        
    maxMaps = Val(GetVar(App.Path & "\cliente.ini", "MAPS", "NumMaps"))
    
    'Dim i As Long
    '
    '    For i = 1 To maxMaps
    '        mapLoad i
    '    Next i
    
    
    mapAutoCreate
    
End Sub
Private Sub mapLoad(Mapa As Long)
    
    Dim FreeHandle As Long
    
        FreeHandle = FreeFile()
      
        Open App.Path & "\Maps\Map" & CStr(Mapa) & ".map" For Binary Access Read As FreeHandle
      
            Get FreeHandle, , mapData()

        Close FreeHandle
End Sub
Private Sub mapAutoCreate()

    Dim X As Long, Y As Long
    Dim dX As Integer, dY As Integer
    
      
    ReDim mapData(1 To MaxTilesX, 1 To MaxTilesY) As mapBlock
                
    For X = 1 To MaxTilesY
        For Y = 1 To MaxTilesX
        
            ' Create Layer 1
            mapData(X, Y).Layer(1).GrhIndex = RandomNumber(1, 2)
      
            ' Create Layer 3
            mapData(X, Y).Layer(3).GrhIndex = RandomNumber(1, 30)
                        
            If mapData(X, Y).Layer(3).GrhIndex > 3 Then
                mapData(X, Y).Layer(3).GrhIndex = 0
            Else
                mapData(X, Y).Layer(3).GrhIndex = 3
            End If
               
            ' Create Particle Layer
            mapData(X, Y).particleIndex = RandomNumber(1, 1000)
                        
            If mapData(X, Y).particleIndex > 19 Then
                mapData(X, Y).particleIndex = 0
            End If
                
            ' PreCalculate position
            If (X Mod 2) = 0 Then
                dX = -HalfTileWidth: dY = -HalfTileHeight
            Else
                dX = -HalfTileWidth: dY = 0
            End If
                
            mapPreCalcPos(X, Y).X = TilePixelWidth * X - TilePixelWidth + dX
            mapPreCalcPos(X, Y).Y = TilePixelHeight * Y - TilePixelHeight + dY
                
        Next Y
    Next X
    
End Sub
Public Sub mapUnloadAll()

    Erase mapData
    
End Sub
Public Function mapLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    If Not mapInBounds(X, Y) Then
        Exit Function
    End If
    
    If mapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    If mapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
       
    mapLegalPos = True
End Function
Function mapInBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
    
    If X < 1 Or X > 100 Or Y < 1 Or Y > 100 Then
        Exit Function
    End If
    
    mapInBounds = True
End Function

#If WorldEditor = 1 Then

' ** Edicion de Mapas **

Public Sub mapAddGrh(ByVal Layer As Byte)

    If frmMain.grhList.ListIndex + 1 > 0 Then

        With mapData(MouseTilesPos.X, MouseTilesPos.Y)
            
            .Layer(Layer).GrhIndex = frmMain.grhList.ListIndex + 1
            
        End With
    
    End If
    
End Sub

Public Sub mapRemoveGrh(ByVal Layer As Byte)

    mapData(MouseTilesPos.X, MouseTilesPos.Y).Layer(Layer).GrhIndex = 0

End Sub


Public Sub mapSave(mapIndex As Long)

        Dim FreeHandle As Long
    
        FreeHandle = FreeFile()
      
        Open App.Path & "\Maps\Map" & CStr(mapIndex) & ".map" For Binary Access Write As FreeHandle
            Put FreeHandle, , mapData
        Close FreeHandle
    
End Sub


#End If

