Attribute VB_Name = "modGame"
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

'Swich on/off the program
Public prgRun As Boolean

Private Type gameTime
    Milliseconds As Long
    Second As Byte
    Minutes As Byte
    Hour As Byte
    Day As Byte
    Month As Byte
    Year As Byte
End Type

Private tMultiply As Integer
Private srvTime As gameTime
Private Sub gameMainLoop()

    Dim lFrameTimer As Long 'Timer

    'Count miliseconds
    Dim countTime As Long
    Dim precountTime As Long
    
    Do While prgRun
    
        precountTime = GetTickCount()
            
        gameCalculateTime countTime
        
        countTime = GetTickCount() - precountTime
        
        'Timer 1 second
        If GetTickCount - lFrameTimer >= 1000 Then
    
            lFrameTimer = GetTickCount
        End If
    
      DoEvents
    Loop

End Sub
Private Sub gameLoad()

    'Init the Randomize Timer
    IntializeRandom

    'Show Form
     frmMain.Show
    
    'Load Maps
    #If Testing = 0 Then
     modMap.mapInitialize
    #End If
    
    'Socket StartUp
     modTCP.sockInitialize 1000, 666
    
    'Load DataBase
     modDataBase.dbInitialize GetVar(App.Path & "\Server.ini", "CONFIG", "Database")   '1=localhost; 2=byethost
    
    'Load GameTime
     modGame.gameLoadTime
     
    'Load characters
     If modCharacter.characterInitialize = False Then Consola "���Critical error when initializing characters!!!" Else Consola "Initializing characters correctly..."
    
    'AOBot is running!
    #If Testing = 0 Then
        prgRun = True
    #Else
        prgRun = False
    #End If
    
    CreateSystemTrayIcon frmMain, "Mercenario Online 1.0 Server"
    
    Consola "Server loaded."

End Sub
Public Sub gameUnLoad()

    'Unload Form
     Unload frmMain

    'DeInit Maps
    #If Testing = 0 Then
     modMap.mapDeInitialize
    #End If
    
    'Close Server
     modTCP.sockDeInitialize
    
    'Close DataBase
     modDataBase.dbClose
    
    'Save gameTime
     modGame.gameSaveTime
     
    'Erase characters
     modCharacter.characterEraseAll
    
    DeleteSystemTrayIcon frmMain
    
    End
    
End Sub
Public Sub Main()
    
    If (App.PrevInstance) Then
        MsgBox "El servidor ya est� abierto."
        End
    End If
    
    'Load Game
    gameLoad
    
    'Run Main Loop
    gameMainLoop
    
    'Unload Game
    #If Testing = 0 Then
     gameUnLoad
    #End If

End Sub
Public Sub gameLoadTime()

   'day/month/year/sec/min/hour
    
    Dim str As String
    
    tMultiply = Val(GetVar(App.Path & "\server.ini", "TIME", "v"))
    str = GetVar(App.Path & "\server.ini", "TIME", "t")
    
    With srvTime
        .Day = Val(ReadField$(1, str, Asc("/")))
        .Month = Val(ReadField$(2, str, Asc("/")))
        .Year = Val(ReadField$(3, str, Asc("/")))
        .Second = Val(ReadField$(4, str, Asc("/")))
        .Minutes = Val(ReadField$(5, str, Asc("/")))
        .Hour = Val(ReadField$(6, str, Asc("/")))
    End With
    
End Sub
Public Sub gameSaveTime()

    Dim str As String

    With srvTime
         str = .Day & "/" & .Month & "/" & .Year & "/" & .Second & "/" & .Minutes & "/" & .Hour
    End With

    WriteVar App.Path & "\server.ini", "TIME", "t", str
    
End Sub
Public Sub gameCalculateTime(ByVal timeAdder As Long)

    With srvTime
    
        .Milliseconds = .Milliseconds + (timeAdder * tMultiply)

        .Second = .Second + (.Milliseconds \ 1000)
        .Minutes = .Minutes + (.Second \ 60)
        .Hour = .Hour + (.Minutes \ 60)
        .Day = .Day + (.Hour \ 24)
        .Month = .Month + (.Day \ 31)
        .Year = .Year + (.Month \ 12)
            
        .Milliseconds = (.Milliseconds Mod 1000)
        .Second = (.Second Mod 60)
        .Minutes = (.Minutes Mod 60)
        .Hour = (.Hour Mod 24)
        .Day = (.Day Mod 31)
        .Month = (.Month Mod 12)
        
        frmMain.lblWorldTime.Caption = "Tiempo : " & Format(.Hour, "00") & ":" & Format(.Minutes, "00") & ":" & Format(.Second, "00") & "  " & Format(.Day, "00") & "/" & Format(.Month, "00") & "/" & Format(.Year, "0000")
        
    End With
        
End Sub


