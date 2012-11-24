VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type PlayerInfoAcc
    Name        As String
    PosChar     As String
    Nivel       As Byte
End Type

Private Type Acc
    Name        As String
    Password    As String
    Email       As String
    NumPlayers  As Byte
    Players()   As PlayerInfoAcc
End Type
    
Private Account As Acc
Public Sub accLogin(ByVal Index As Integer, ByVal AccountName As String, ByVal Password As String)
    
    Dim AccPath As String
    AccPath = App.Path & "\Accounts\" & AccountName & ".acc"
    
    'Comprobaciones
    If Not FileExist(AccPath, vbNormal) Then
        sckClose (Index)
        Exit Sub
    End If
    
    If GetVar(AccPath, "INFO", "Ban") = "true" Then
        sckClose (Index)
        Exit Sub
    End If
    
    'Check if the password suplyed is correct
    If Password <> GetVar(AccPath, "INFO", "password") Then
        sckClose (Index)
        Exit Sub
    End If
    
    
 'Read numPlayers
  Account.NumPlayers = Val(GetVar(AccPath, "INFO", "NumPlayers"))


 If Account.NumPlayers > 0 Then
 
   ReDim Preserve Account.Players(1 To Account.NumPlayers)
   
    Dim Packet As String, i As Byte
       
       'Read Character data
        For i = 1 To Account.NumPlayers
        
            With Account.Players(i)
                AccPath = App.Path & "\Accounts\" & AccountName & ".acc" ' AccPath
                    .Name = GetVar(AccPath, "PLAYERS", i)
                
                    'Lo aprobechamos para la lectura nuevo path
                    AccPath = App.Path & "\Players\" & .Name & ".ini"

                    .PosChar = GetVar(AccPath, "INIT", "Position")
                    .Nivel = GetVar(AccPath, "STATS", "ELV")
        
                Packet = Packet & _
                i & "," & .Name & "," & .PosChar & "," & .Nivel & "~"
            End With
            
            'Server packet send
            'Index+InfoAccount

          DoEvents
        Next i

 Else
    'Server packet send
    'Index=0
    Exit Sub
 End If

End Sub
Public Function accCreate(ByVal AccountName As String, ByVal Password As String, ByVal Email As String) As Boolean
    Dim i As Byte
    
    'Check if Account already exists
    If FileExist(App.Path & "\Accounts\" & AccountName & ".acc", vbNormal) Then
        accCreate = False
        Exit Function
    End If
    

    Dim N As Integer
        N = FreeFile()

    Open App.Path & "\Accounts\" & AccountName & ".acc" For Output As N
        Print #N, "[INFO]"
        Print #N, "Account=" & AccountName
        Print #N, "Password=" & Password
        Print #N, "Email=" & Email
        Print #N, "Ban=False"
        Print #N, "NumPlayers=0"
        Print #N, "[PLAYERS]"
        
        For i = 1 To 8
            Print #N, CStr(i) & "="
        Next i
    Close N
    
    accCreate = True
End Function
Public Function accBan(ByVal Ban As Boolean, AccountName As String) As Boolean
    'Make sure Account exists
    If Not FileExist(App.Path & "\Accounts\" & AccountName & ".acc", vbNormal) Then
        accBan = False
        Exit Function
    End If
    
    'Change password
    If Ban = True Then WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "Ban", "true" Else _
         WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "Ban", "false"
    
    accBan = True
End Function
Public Function accPassChange(ByVal AccountName As String, ByVal AccPasswordOld As String, ByVal AccPasswordNew As String) As Boolean

    'Make sure Account exists
    If Not FileExist(App.Path & "\Accounts\" & AccountName & ".acc", vbNormal) Then
        accPassChange = False
        Exit Function
    End If
    
    'Check if the password suplyed is correct
    If Not AccPasswordOld = GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "Password") Then
        accPassChange = False
        Exit Function
    End If
    
    'Change password
    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "password", AccPasswordNew
    
    accPassChange = True
End Function
Public Function accCharacterAdd(ByVal AccountName As String, ByVal playerName As String) As Boolean

    Dim Slot As Byte
    
    'Obtenemos el slot
    Dim i As Byte
        For i = 1 To 8
            If GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", CStr(i)) = "" Then
                Slot = i
                Exit For
            End If
        Next i
        
    'Check slot
    If Slot > 8 Or Slot < 0 Then
        accCharacterAdd = False
        Exit Function
    End If
    'Make sure the slot is free
    If GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", CStr(Slot)) <> "" Then
        accCharacterAdd = False
        Exit Function
    End If

    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", CStr(Slot), playerName
    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "NumPlayers", CStr(Slot)
    
    accCharacterAdd = True
End Function
Public Function accCharacterRemove(ByVal AccountName As String, ByVal Slot As Byte, ByVal playerName As String) As Boolean
    
    If Slot > 8 Or Slot < 0 Then
        accCharacterRemove = False
        Exit Function
    End If
    
    Account.NumPlayers = GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "NumPlayers")
    
    If Account.NumPlayers = 0 Then
        accCharacterRemove = False
        Exit Function
    End If

    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "INFO", "NumPlayers", Account.NumPlayers - 1
    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", CStr(Slot), ""
    
    'Matamos el user
    Kill App.Path & "\Characters\" & playerName & ".ini"
    
    'Acomoda =)
    Dim s As String, N As String  'cheksum
    Dim i As Byte
        For i = 1 To (Account.NumPlayers - 1)
            s = GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", i)
            N = GetVar(App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", i + 1)
            
                If s = vbNullString And N <> vbNullString Then
                    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", i, N
                    WriteVar App.Path & "\Accounts\" & AccountName & ".acc", "PLAYERS", i + 1, vbNullString
                End If
        Next i
        
    accCharacterRemove = True
End Function


