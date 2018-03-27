Attribute VB_Name = "modSegurity"
Option Explicit

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

Public Function encryptString(ByVal str As String) As String

    Dim strCount As Integer: strCount = Len(str)
    
    Dim strOutput As String

    
    Do While (strCount <> 0)
        
        strOutput = strOutput & CStr(RandomNumber(0, 9) & RandomNumber(0, 9) & RandomNumber(0, 9)) & _
                    Len(CStr(Asc(mid$(str, strCount, strCount)))) & CStr(Asc(mid$(str, strCount, strCount)))
        
        strCount = strCount - 1
    
    Loop
    
    strOutput = strOutput & CStr(RandomNumber(0, 9) & RandomNumber(0, 9) & RandomNumber(0, 9))
    
    encryptString = strOutput

End Function
Public Function decryptString(ByVal str As String) As String

    Dim strCount As Integer: strCount = Len(str)
    
    Dim strOutput As String
    Dim LenChr As Byte
    
    Do While (strCount > 3)
        
        str = mid$(str, 4, Len(str))
        
        strCount = strCount - 3
        
        LenChr = CByte(Left$(str, 1))
        
        strCount = strCount - LenChr
        
        str = mid$(str, 2, Len(str))
        
        strCount = strCount - 1

        strOutput = Chr$(Left$(str, LenChr)) & strOutput
        
        str = mid$(str, LenChr + 1, Len(str))
        
    Loop
    
    decryptString = strOutput

End Function
