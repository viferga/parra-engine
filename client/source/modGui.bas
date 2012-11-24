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
Public Function guiInitialize() As Boolean
On Error GoTo errHandle
    
    Dim Path As String: Path = App.Path & "\Init\Gui.ini"
    Dim Size As Integer: Size = GetVar(Path, "INIT", "size")
    
    ReDim guiBox(1 To Size) As structBox
    
    
    Dim i As Integer
        
        For i = 1 To Size
        
            guiCreateBox i, CSng(GetVar(Path, CStr(i), "left")), CSng(GetVar(Path, CStr(i), "top")), _
                            CSng(GetVar(Path, CStr(i), "bottom")), CSng(GetVar(Path, CStr(i), "right")), _
                        D3DColorARGB(CSng(ReadField$(1, GetVar(Path, CStr(i), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(2, GetVar(Path, CStr(i), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(3, GetVar(Path, CStr(i), "colorup"), Asc("-"))), _
                                     CSng(ReadField$(4, GetVar(Path, CStr(i), "colorup"), Asc("-")))), _
                        D3DColorARGB(CSng(ReadField$(1, GetVar(Path, CStr(i), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(2, GetVar(Path, CStr(i), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(3, GetVar(Path, CStr(i), "colordown"), Asc("-"))), _
                                     CSng(ReadField$(4, GetVar(Path, CStr(i), "colordown"), Asc("-")))), _
                            CStr(GetVar(Path, CStr(i), "text")), CSng(GetVar(Path, CStr(i), "textindex"))
                            
                            
        Next i
    
    guiInitialize = True
    Exit Function
    
errHandle:
    guiInitialize = False
    MsgBox "Error in GUI" & vbNewLine & Err.Description
End Function
Public Sub guiCreateBox(Index As Integer, Left As Integer, Top As Integer, Bottom As Integer, Right As Integer, ColorUp As Long, ColorDown As Long, Text As String, TextIndex As Integer)
    
    With guiBox(Index)
        
        .Text = Text
        .TextIndex = TextIndex
        
        .ColorDown = ColorDown
        .ColorUp = ColorUp
        
        With .BoxSize
        
            .X1 = Left
            .X2 = Right
            .Y1 = Top
            .Y2 = Bottom
        
            guiBox(Index).GeometryVert(0) = setVertex(.X1, .Y1 + .X2, 0, 1, ColorDown, 0, 0, 0)
            guiBox(Index).GeometryVert(1) = setVertex(.X1, .Y1, 0, 1, ColorUp, 0, 1, 0)
            guiBox(Index).GeometryVert(2) = setVertex(.X1 + .Y2, .Y1 + .X2, 0, 1, ColorDown, 0, 0, 1)
            guiBox(Index).GeometryVert(3) = setVertex(.X1 + .Y2, .Y1, 0, 1, ColorUp, 0, 1, 1)

         End With
         
    End With

End Sub
Public Sub guiEvents(ByVal X As Single, ByVal Y As Single)

    Dim i As Long
        
        For i = 1 To UBound(guiBox())
            
            With guiBox(i).BoxSize
                    
                'Static tempX As Single, tempY As Single
                        
                        'tempX = X - .X1
                        'tempY = Y - .Y1
                        
                        .X1 = X + (X + (.X2 - .X1))
                        .Y1 = Y + (Y + (.Y2 - .Y1))
                        
                        guiBox(i).GeometryVert(0) = setVertex(.X1, .Y1 + .X2, 0, 1, guiBox(i).ColorDown, 0, 0, 0)
                        guiBox(i).GeometryVert(1) = setVertex(.X1, .Y1, 0, 1, guiBox(i).ColorUp, 0, 1, 0)
                        guiBox(i).GeometryVert(2) = setVertex(.X1 + .Y2, .Y1 + .X2, 0, 1, guiBox(i).ColorDown, 0, 0, 1)
                        guiBox(i).GeometryVert(3) = setVertex(.X1 + .Y2, .Y1, 0, 1, guiBox(i).ColorUp, 0, 1, 1)

            End With
        
        Next i
End Sub
Public Sub guiRender()
    
    Dim i As Long
        
        For i = 1 To UBound(guiBox())
            
            D3DDevice.SetTexture 0, Nothing
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, guiBox(i).GeometryVert(0), Len(guiBox(i).GeometryVert(0))
        
        Next i
        
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
