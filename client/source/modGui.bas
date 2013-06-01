Attribute VB_Name = "modGui"
Option Explicit

'Module to management textures(GUI)

Public RenderGUI As Boolean

Public Type structBoxSize
    X1 As Single: X2 As Single
    Y1 As Single: Y2 As Single
End Type


Public Type structBox
    BoxSize As structBoxSize
    GeometryVert(3) As D3DTLVERTEX
    
    ColorDown As Long
    ColorUp As Long
    
    Text As String
    TextIndex As Integer
End Type: Public guiBox() As structBox

Dim Gui_TextureLogo     As Direct3DTexture8 'textura

Public Function guiInitialize() As Boolean
On Error GoTo errHandle
    
    Dim Path As String: Path = App.Path & "\Init\Gui.ini"
    Dim Size As Integer: Size = GetVar(Path, "INIT", "size")
    
    ReDim guiBox(1 To Size) As structBox
    
    
    Dim I As Integer
        
        For I = 1 To Size
        
            guiCreateBox I, CSng(GetVar(Path, CStr(I), "left")), CSng(GetVar(Path, CStr(I), "top")), _
                            CSng(GetVar(Path, CStr(I), "bottom")), CSng(GetVar(Path, CStr(I), "right")), _
                        D3DColorARGB(CSng(ReadField$(1, GetVar(Path, CStr(I), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(2, GetVar(Path, CStr(I), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(3, GetVar(Path, CStr(I), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(4, GetVar(Path, CStr(I), "colorup"), Asc("-")))), _
                        D3DColorARGB(CSng(ReadField$(1, GetVar(Path, CStr(I), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(2, GetVar(Path, CStr(I), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(3, GetVar(Path, CStr(I), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(4, GetVar(Path, CStr(I), "colordown"), Asc("-")))), _
                            CStr(GetVar(Path, CStr(I), "text")), CSng(GetVar(Path, CStr(I), "textindex"))
                            
                            
        Next I
    
    guiInitialize = True
    Exit Function
    
errHandle:
    guiInitialize = False
    MsgBox "Error in GUI" & vbNewLine & Err.Description
End Function
Public Sub guiCreateBox(index As Integer, Left As Integer, Top As Integer, Bottom As Integer, Right As Integer, ColorUp As Long, ColorDown As Long, Text As String, TextIndex As Integer)
    
    With guiBox(index)
        
        .Text = Text
        .TextIndex = TextIndex
        
        .ColorDown = ColorDown
        .ColorUp = ColorUp
        
        With .BoxSize
        
            .X1 = Left
            .X2 = Right
            .Y1 = Top
            .Y2 = Bottom
        
            guiBox(index).GeometryVert(0) = setVertex(.X1, .Y1 + .X2, 0, 1, ColorDown, 0, 0, 0)
            guiBox(index).GeometryVert(1) = setVertex(.X1, .Y1, 0, 1, ColorUp, 0, 1, 0)
            guiBox(index).GeometryVert(2) = setVertex(.X1 + .Y2, .Y1 + .X2, 0, 1, ColorDown, 0, 0, 1)
            guiBox(index).GeometryVert(3) = setVertex(.X1 + .Y2, .Y1, 0, 1, ColorUp, 0, 1, 1)

         End With
         
    End With

End Sub
Public Sub guiEvents(ByVal x As Single, ByVal y As Single)

    Dim I As Long
        
        For I = 1 To UBound(guiBox())
            
            With guiBox(I).BoxSize
                    
                'Static tempX As Single, tempY As Single
                        
                        'tempX = X - .X1
                        'tempY = Y - .Y1
                        
                        .X1 = x + (x + (.X2 - .X1))
                        .Y1 = y + (y + (.Y2 - .Y1))
                        
                        guiBox(I).GeometryVert(0) = setVertex(.X1, .Y1 + .X2, 0, 1, guiBox(I).ColorDown, 0, 0, 0)
                        guiBox(I).GeometryVert(1) = setVertex(.X1, .Y1, 0, 1, guiBox(I).ColorUp, 0, 1, 0)
                        guiBox(I).GeometryVert(2) = setVertex(.X1 + .Y2, .Y1 + .X2, 0, 1, guiBox(I).ColorDown, 0, 0, 1)
                        guiBox(I).GeometryVert(3) = setVertex(.X1 + .Y2, .Y1, 0, 1, guiBox(I).ColorUp, 0, 1, 1)

            End With
        
        Next I
End Sub

Public Sub guiDestroy()

    Erase guiBox()

End Sub
    'Gui experimental
    'If GUI = True Then
    '    deviceRenderBox 100, 100, 200, 300, D3DColorARGB(210, 150, 150, 150), D3DColorARGB(210, 60, 60, 60), D3DColorARGB(210, 150, 150, 150), D3DColorARGB(210, 60, 60, 60)
    '    deviceRenderBox 100, 100, 2, 300, D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220)
    '    deviceRenderBox 100, 115, 2, 300, D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220)
    '    deviceRenderBox 100, 100, 200, 2, D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220)
    '    deviceRenderBox 100, 298, 2, 300, D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220)
    '    deviceRenderBox 400, 100, 200, 2, D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220), D3DColorARGB(240, 220, 220, 220)
    '
    '    fontRender "Map Editor Controls", 4, 178, 102, 217, 110, DT_RIGHT
    'End If
