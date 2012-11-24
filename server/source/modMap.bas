Attribute VB_Name = "modMap"
Option Explicit

' MapObj (Define the information of the object)
'     ObjIndex -> Index of Object
'     Amount   -> Amount of Object

Private Type MapObj
    ObjIndex        As Integer
    Amount          As Integer
End Type

' Triggers (Define the Tigger Type of Map)
'     None       -> Not trigger
'     Indoors    -> Doesn't affect the weather in this trigger
'     Unused     -> This Trigger not used
'     InvalidPos -> NPCs can't step on the tiles with this trigger
'     SafeArea   -> Can't fight or steal from this trigger
'     AntiObs..  -> This trigger prevents obstructions in the street
'     BattleZone -> In this trigger is allowed to fight
    
Private Enum mapTrigger
    None = 0
    Indoors = 1
    Unused = 2
    InvalidPos = 3
    SafeArea = 4
    AntiObstruction = 5
    BattleZone = 6
End Enum

' Map Tile Block (Contains tile information)
'      Blocked  -> Defines whether the tile is locked or unlocked
'      Graphic  -> Index of Graphic
'      NpcIndex -> Index of NPC's
'      ObjInfo  -> Manages the information of the object on the map
'      TileExit -> Map shows the shuttle

Private Type MapBlock
    Blocked         As Byte
    Graphic(1 To 4) As Integer

    npcIndex        As Integer
    ObjInfo         As MapObj
    
    TileExit        As characterWorldPos
    Trigger         As mapTrigger
End Type

' Map Info Block (Contains tile properties)
'      Name    -> Name of map
'      Music   -> Music index of map
'      Pk      -> Indicates whether the map is safe or unsafe (0=safe / 1=unsafe)

'      Magia.. -> Indicates whether user can use magic on the map
'      Invi..  -> Indicates wheater user can use invisibility on the map
'      Resu..  -> Indicates wheater user can resurrect in this map

Private Type MapInfo
    Name            As String
    Music           As String
    Pk              As Boolean
    
    MagN            As Byte
    InvN            As Byte
    ResN            As Byte
    
    Terreno         As String
    Zona            As String
    Restringir      As String
    BackUp          As Byte
End Type

Private Type sMap
    MapData()       As MapBlock
    MapInfo         As MapInfo
End Type

Public srvMap()     As sMap    'Array of maps
Public srvMapSize   As Integer 'Number of maps
Public Sub mapInitialize()
    Dim i As Integer

    srvMapSize = Val(GetVar(App.Path & "\server.ini", "CONFIG", "mapSize"))
    
    Consola "Loading Maps..."

    ReDim srvMap(1 To srvMapSize) As sMap
    
    For i = 1 To srvMapSize
        ReDim srvMap(i).MapData(1 To 100, 1 To 100)
        mapLoading i
        DoEvents
    Next i
    
End Sub
Public Sub mapDeInitialize()
    Dim i As Integer
    
    Consola "Liberando Recursos (Maps)..."
    
    For i = 1 To srvMapSize
        Erase srvMap(i).MapData()
        Erase srvMap
    Next i
    
End Sub

Public Function mapInMapBounds(ByRef X As Byte, ByRef Y As Byte) As Boolean

    If (X < 1 Or X > 100 Or Y < 1 Or Y > 100) Then
        mapInMapBounds = False
    Else
        mapInMapBounds = True
    End If
    
End Function
Private Sub mapLoading(ByVal Map As Integer)
    Dim FileMap As Long, FileInf As Long
    Dim Y As Long, X As Long
    
    Dim ByFlags As Byte
    Dim MapPath As String
    
  MapPath = App.Path & "\Mapas\Mapa" & Map
    
  If Not FileExist(MapPath & ".dat") Or Not FileExist(MapPath & ".inf") Or _
        Not FileExist(MapPath & ".map") Then Exit Sub
    
    
  FileMap = FreeFile
    Open MapPath & ".map" For Binary As #FileMap
        Seek #FileMap, 1
    
  FileInf = FreeFile
    Open MapPath & ".inf" For Binary As #FileInf
        Seek #FileInf, 1

    For Y = 1 To 100
        For X = 1 To 100
            
            '.dat file
            Get #FileMap, , ByFlags

            If ByFlags And 1 Then srvMap(Map).MapData(X, Y).Blocked = 1
            
            Get #FileMap, , srvMap(Map).MapData(X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get #FileMap, , srvMap(Map).MapData(X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get #FileMap, , srvMap(Map).MapData(X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get #FileMap, , srvMap(Map).MapData(X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get #FileMap, , ByFlags
                srvMap(Map).MapData(X, Y).Trigger = ByFlags
            End If
            
            Get #FileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get #FileInf, , srvMap(Map).MapData(X, Y).TileExit.Map
                Get #FileInf, , srvMap(Map).MapData(X, Y).TileExit.X
                Get #FileInf, , srvMap(Map).MapData(X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get #FileInf, , srvMap(Map).MapData(X, Y).npcIndex
                
                If srvMap(Map).MapData(X, Y).npcIndex > 0 Then
                    'Si el npc debe hacer respawn en la pos original la guardamos
                    If Val(GetVar(App.Path & "\Dat\NPCs.dat", "NPC" & srvMap(Map).MapData(X, Y).npcIndex, "PosOrig")) = 1 Then
                    '    srvMap(Map).MapData(X, Y).NpcIndex = OpenNPC(srvMap(Map).MapData(X, Y).NpcIndex)
                    '    Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Orig.Map = Map
                    '    Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Orig.X = X
                    '    Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Orig.Y = Y
                    Else
                    '    srvMap(Map).MapData(X, Y).NpcIndex = OpenNPC(srvMap(Map).MapData(X, Y).NpcIndex)
                    End If
                    
                    'Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Pos.Map = Map
                    'Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Pos.X = X
                    'Npclist(srvMap(Map).MapData(X, Y).NpcIndex).Pos.Y = Y
                    
                    'MakeNPCChar True, 0, srvMap(Map).MapData(X, Y).NpcIndex, Map, X, Y
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get #FileInf, , srvMap(Map).MapData(X, Y).ObjInfo.ObjIndex
                Get #FileInf, , srvMap(Map).MapData(X, Y).ObjInfo.Amount
            End If
        Next X
    Next Y
    
       Close #FileMap
    Close #FileInf
    
    Dim Leer As New clsIniReader
    Leer.Initialize (MapPath & ".dat")
    
    srvMap(Map).MapInfo.Name = Leer.GetValue("Mapa" & Map, "Name")
    srvMap(Map).MapInfo.Music = Leer.GetValue("Mapa" & Map, "MusicNum")
    
    If Val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then _
        srvMap(Map).MapInfo.Pk = True Else srvMap(Map).MapInfo.Pk = False
    
    srvMap(Map).MapInfo.MagN = Val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
    srvMap(Map).MapInfo.InvN = Val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
    srvMap(Map).MapInfo.ResN = Val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
    
    srvMap(Map).MapInfo.Terreno = Leer.GetValue("Mapa" & Map, "Terreno")
    srvMap(Map).MapInfo.Zona = Leer.GetValue("Mapa" & Map, "Zona")
    srvMap(Map).MapInfo.Restringir = Leer.GetValue("Mapa" & Map, "Restringir")
    srvMap(Map).MapInfo.BackUp = Val(Leer.GetValue("Mapa" & Map, "BACKUP"))
End Sub

