Attribute VB_Name = "modAccount"
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

Private Type Acc
    Name        As String
    Password    As String
    Email       As String
    NumPlayers  As Byte
    Players(1 To 6) As String
End Type
    
Private Account As Acc
Public Function accountConnect(accountName As String, accountPassword As String) As Boolean

    accountConnect = False
        
    'Comprobaciones
    If accountExist(accountName) = False Then Exit Function
    
    'Check if the password suplyed is correct
    If accountCheckPassword(accountName, accountPassword) = False Then Exit Function
    
    'ChekBan
    If accountCheckBan(accountName) = True Then Exit Function
    
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `cuentas` WHERE accountname='" & accountName & "'")
    
    With Account
    
        .Name = accountName
        .Password = accountPassword
        .Email = Recordset!Email
        
        Dim i As Long
        
            For i = 1 To 6
                If Recordset.Fields("pj" & CStr(i)).ActualSize = 0 Then
                    .NumPlayers = (i - 1)
                    Exit For
                End If
 
                .Players(i) = Recordset.Fields("pj" & CStr(i))
 
            Next i
                        
    End With
    
    Set Recordset = Nothing
    
    accountConnect = True

End Function
Public Function accountSendInfo(buffer As clsByteQueue) As Boolean

    accountSendInfo = False

    Dim i As Long

    'Send AccountInfo to Client

    With Account
    
        buffer.WriteByte .NumPlayers
        
        For i = 1 To .NumPlayers
            
            buffer.WriteASCIIString .Players(i)
            
        Next i
        
    End With
    
    accountSendInfo = True

End Function
Public Function accountCreate(accountName As String, Password As String, Email As String) As Boolean
    
    accountCreate = False
    
    'Check if Account already exists
    If accountExist(accountName) = False Then Exit Function
    
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("INSERT INTO `cuentas` (accountname,password,email,ban,pj1,pj2,pj3,pj4,pj5,pj6) " & _
                                    "values('" & accountName & "','" & Password & "','" & Email & "','0'," & _
                                        "'NULL','NULL','NULL','NULL','NULL','NULL')")
                                        
    accountCreate = Not Recordset Is Nothing
    
    Set Recordset = Nothing
    
End Function
Public Function accountKill(accountName As String, accountPassword As String, accountMail As String) As Boolean
    
    If (accountExist(accountName) = False) Then Exit Function

    If (accountCheckPassword(accountName, accountPassword) = False) Then Exit Function
    
    If (accountCheckMail(accountMail) = False) Then Exit Function
    
    ' Eliminamos todos los pjs de la cuenta
    
    Set Recordset = Nothing
    
    ' Kill Account
    Set Recordset = Connection.Execute("DELETE FROM `cuentas` WHERE accountname='" & accountName & "'")

    Set Recordset = Nothing
    
    accountKill = True
End Function
Public Function accountPassChange(ByVal accountName As String, ByVal AccPasswordOld As String, ByVal AccPasswordNew As String) As Boolean

    'Make sure Account exists
    If (accountExist(accountName) = False) Then Exit Function
    
    'Check if the password suplyed is correct
    If (accountCheckPassword(accountName, AccPasswordOld) = False) Then Exit Function
    
    'Change password
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("INSERT INTO `cuentas` (password) values('" & AccPasswordNew & "')")
    
    Set Recordset = Nothing
    
    accountPassChange = True
End Function
Public Function accountCharacterAdd(ByVal accountName As String, ByVal playerName As String) As Boolean

    Dim Slot As Byte
    
    'Obtenemos el slot
'    Dim i As Byte
'        For i = 1 To 8
'            If GetVar(App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", CStr(i)) = "" Then
'                Slot = i
'                Exit For
'            End If
'        Next i
'
'    'Check slot
'    If Slot > 8 Or Slot < 0 Then
'        accountCharacterAdd = False
'        Exit Function
'    End If
'    'Make sure the slot is free
'    If GetVar(App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", CStr(Slot)) <> "" Then
'        accountCharacterAdd = False
'        Exit Function
'    End If

'    Connection.Execute "INSERT INTO cuentas(pj" & CStr(Slot) & "name) values('" & playerName & ")", , adExecuteNoRecords
    
    accountCharacterAdd = True
End Function
Public Function accountCharacterRemove(ByVal accountName As String, ByVal Slot As Byte, ByVal playerName As String) As Boolean
    
'    If Slot > 8 Or Slot < 0 Then
'        accCharacterRemove = False
'        Exit Function
'    End If
'
'    Account.NumPlayers = GetVar(App.Path & "\Accounts\" & accountName & ".acc", "INFO", "NumPlayers")
'
'    If Account.NumPlayers = 0 Then
'        accCharacterRemove = False
'        Exit Function
'    End If
'
'    WriteVar App.Path & "\Accounts\" & accountName & ".acc", "INFO", "NumPlayers", Account.NumPlayers - 1
'    WriteVar App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", CStr(Slot), ""
'
'    'Matamos el user
'    Kill App.Path & "\Characters\" & playerName & ".ini"
'
'    'Acomoda =)
'    Dim s As String, N As String  'cheksum
'    Dim i As Byte
'        For i = 1 To (Account.NumPlayers - 1)
'            s = GetVar(App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", i)
'            N = GetVar(App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", i + 1)
'
'                If s = vbNullString And N <> vbNullString Then
'                    WriteVar App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", i, N
'                    WriteVar App.Path & "\Accounts\" & accountName & ".acc", "PLAYERS", i + 1, vbNullString
'                End If
'        Next i
                
    accountCharacterRemove = True
End Function
Private Function accountExist(Name As String) As Boolean

    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `cuentas` WHERE accountname='" & Name & "'")

    accountExist = Not Recordset Is Nothing
    
    Set Recordset = Nothing
    
End Function
Private Function accountCheckPassword(Name As String, Password As String) As Boolean
    
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `cuentas` WHERE accountname='" & Name & "' AND password='" & Password & "'")
    
    If Recordset.EOF Or Recordset.BOF = True Then
        accountCheckPassword = False
        Set Recordset = Nothing
        Exit Function
    End If
    
    accountCheckPassword = (Recordset!Password = Password And Recordset!accountName = Name)
    
    Set Recordset = Nothing
    
End Function
Private Function accountCheckBan(Name As String) As Boolean
    
    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `cuentas` WHERE accountname='" & Name & "'")
    
    If Recordset.EOF Or Recordset.BOF = True Then
        accountCheckBan = False
        Set Recordset = Nothing
        Exit Function
    End If
    
    accountCheckBan = (Recordset!Ban = 1)

    Set Recordset = Nothing
    
End Function
Private Function accountCheckMail(accountMail As String) As Boolean

    Set Recordset = Nothing
    
    Set Recordset = Connection.Execute("SELECT * FROM `cuentas` WHERE email='" & accountMail & "'")

    If Recordset.EOF Or Recordset.BOF = True Then
        accountCheckMail = False
        Set Recordset = Nothing
        Exit Function
    End If

    accountCheckMail = (Recordset!Email = accountMail)

    Set Recordset = Nothing
    
End Function
