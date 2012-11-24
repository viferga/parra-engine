VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoaderMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DirectShow Objects
Private DSAudio  As IBasicAudio
Private DSEvent As IMediaEvent
Private DSControl As IMediaControl
Private DSPosition As IMediaPosition
Public Function Destroy() As Boolean
On Error GoTo BailOut

    If ObjPtr(DSControl) > 0 Then
        DSControl.Stop
    End If
                
    If ObjPtr(DSAudio) Then Set DSAudio = Nothing
    If ObjPtr(DSEvent) Then Set DSEvent = Nothing
    If ObjPtr(DSControl) Then Set DSControl = Nothing
    If ObjPtr(DSPosition) Then Set DSPosition = Nothing
                
    Destroy = True
    Exit Function
BailOut:
    Destroy = False
    Debug.Print "Error in Buffer Sound Destroy"; Err.Number; Err.Description
End Function
Public Function Load(FileName As String) As Boolean

    '   Destroy existing instances
        If Not (Destroy() = True) Then Exit Function
        
    '   Setup a filter graph for the file
        Set DSControl = New FilgraphManager
        Call DSControl.RenderFile(FileName)
    
    '   Setup the basic audio object
        Set DSAudio = DSControl
        DSAudio.Volume = 0
        DSAudio.Balance = 0
    
    '   Setup the media event and position objects
        Set DSEvent = DSControl
        Set DSPosition = DSControl
        If ObjPtr(DSPosition) Then DSPosition.Rate = 1#
        DSPosition.CurrentPosition = 0

    Load = True
End Function

Public Function Play() As Boolean
On Error GoTo BailOut

    DSControl.Run

    Play = True
    Exit Function
BailOut:
    Play = False
    Debug.Print "Error in Buffer Sound Play"; Err.Number; Err.Description
End Function

Public Function Stopping() As Boolean
On Error GoTo BailOut
    
    DSControl.Stop
    DSPosition.CurrentPosition = 0 'set it back to the beginning
    
    Stopping = True
    Exit Function
BailOut:
    Stopping = False
    Debug.Print "Error in Buffer Sound Stop"; Err.Number; Err.Description
End Function

Public Function Pause() As Boolean
On Error GoTo BailOut
    
    DSControl.Stop
    
    Pause = True
    Exit Function
BailOut:
    Pause = False
    Debug.Print "Error in Buffer Sound Pause"; Err.Number; Err.Description
End Function
Public Function SetSpeed(Speed As Single) As Boolean
On Error GoTo BailOut

    If ObjPtr(DSPosition) > 0 Then
        DSPosition.Rate = Speed
    End If

    SetSpeed = True
    Exit Function
BailOut:
    SetSpeed = False
    Debug.Print "Error in Buffer Sound Speed"; Err.Number; Err.Description
End Function

Public Function SetVolume(Volume As Long) As Boolean
On Error GoTo BailOut
    
    'Set the new volume
    If ObjPtr(DSControl) > 0 Then
        DSAudio.Volume = Volume * 40 ' makes it in the 0 to -4000 range (-4000 is almost silent)
    End If

    SetVolume = True
    Exit Function
BailOut:
    SetVolume = False
    Debug.Print "Error in Buffer Sound Volume"; Err.Number; Err.Description
End Function

Public Function SetBalance(Balance As Long) As Boolean
On Error GoTo BailOut

    If ObjPtr(DSControl) > 0 Then
        DSAudio.Balance = Balance
    End If

    SetBalance = True
    Exit Function
BailOut:
    SetBalance = False
    Debug.Print "Error in Buffer Sound Balance"; Err.Number; Err.Description
End Function
Public Function Status() As Byte
    
    Dim Length As Long
    
    Length = Round(DSPosition.Duration, 2)

    'Assign specified starting position dependent on state
    If CLng(DSPosition.CurrentPosition) > 0 And CLng(DSPosition.CurrentPosition) < Length Then
        Status = 1
    ElseIf CLng(DSPosition.CurrentPosition) = Length Then
        Status = 0
    End If

End Function
