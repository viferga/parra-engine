Attribute VB_Name = "modDataBase"
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

' Module for Managing and Securing Data Stored in the DataBase

'======================================
'MySQL Server Configuration
'Server:    localhost
'UserName:  root
'Password:  root
'DataBase:  server_database
'Port:      3306
'======================================

Public Connection As ADODB.Connection
Public Recordset As ADODB.Recordset

Public Function dbInitialize(ByRef Index As Byte) As Boolean
On Error GoTo errSQL

  'Connect to MySQL server using MySQL ODBC 5.1 Driver

  Consola "Loading DataBase..."
  
  Set Connection = New ADODB.Connection
  Set Recordset = New ADODB.Recordset
    
  Connection.CommandTimeout = 40
                                  
  Connection.CursorLocation = adUseClient
                          
  'Connect MySQL Server Without ODBC setup
  Connection.Open getConnString(Index)
    
  With Recordset
      .ActiveConnection = Connection
      .CursorLocation = adUseServer
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
  End With

  dbInitialize = True
  Exit Function
  
errSQL:
    Debug.Print "Error occurred when trying to connect to the Server DataBase"
    MsgBox "Error MySQL: " & vbNewLine & Err.Description
    Consola "���Server error when trying to connect to Database!!!"

    'Close DataBase
    dbClose

    dbInitialize = False
End Function
Private Function getConnString(ByRef Index As Byte) As String

    Dim Path As String
    
    Path = App.Path & "\server.ini"

    getConnString = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & _
                    "DRIVER={MySQL ODBC 3.51 Driver};" & _
                    "DESC=;" & _
                    "SERVER=" & GetVar(Path, "SQL" & CStr(Index), "server") & ";" & _
                    "DATABASE=" & GetVar(Path, "SQL" & CStr(Index), "dbname") & ";" & _
                    "UID=" & GetVar(Path, "SQL" & CStr(Index), "dbuser") & ";" & _
                    "PASSWORD=" & GetVar(Path, "SQL" & CStr(Index), "dbpass") & ";" & _
                    "PORT=" & GetVar(Path, "SQL" & CStr(Index), "dbport") & ";" & _
                    "OPTION=16387;STMT=;" & Chr$(34)

                    

End Function
Public Sub dbClose()

  Set Recordset = Nothing

  If Not Connection Is Nothing Then
  
    'Close DataBase
    If Connection.State <> 0 Then Connection.Close
    Set Connection = Nothing
    
  End If
  
End Sub
