Attribute VB_Name = "modLights"
Option Explicit
Public Function CalcVertexLight(ByVal Radio As Byte, LightX As Single, LightY As Single, VertexX As Single, VertexY As Single, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim lXDistance As Single
    Dim lYDistance As Single
    Dim lVertexDistance As Single
    Dim pRadio As Long
    
    pRadio = Radio * 64
        
    ' Calculate distance from vertex
    lXDistance = CLng((Abs(LightX + 64 - VertexX)) / 2)
    lYDistance = CLng(Abs(LightY + 64 - VertexY))
    
    lVertexDistance = CLng(Sqr(lXDistance * lXDistance + lYDistance * lYDistance))
    
    If lVertexDistance <= pRadio Then
    
        Dim CurrentColor As D3DCOLORVALUE
        
        D3DXColorLerp CurrentColor, LightColor, AmbientColor, lVertexDistance / pRadio
        CalcVertexLight = D3DColorXRGB(CurrentColor.r, CurrentColor.g, CurrentColor.b)
    Else
        ' Return lowest value
        CalcVertexLight = D3DColorXRGB(AmbientColor.r, AmbientColor.g, AmbientColor.b)
    End If
End Function
