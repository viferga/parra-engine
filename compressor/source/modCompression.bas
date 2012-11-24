Attribute VB_Name = "modCompression"
Option Explicit

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum srcFileType
    BMP
    MIDI
    MP3
End Enum

Private Const SrcPath As String = "\..\Client\Resources\"

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long

' *********** Compress / Uncompress  Data Functions ***********

Private Sub dataCompress(ByRef data() As Byte)
    
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim loopc As Long
    
    Dimensions = UBound(data)
    
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    compress BufTemp(0), DimBuffer, data(0), Dimensions
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    data(0) = data(0) Xor 123
End Sub

Private Sub dataDecompress(ByRef data() As Byte, ByVal OrigSize As Long)
    
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor 123
    
    uncompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Private Sub encryptHeaderFile(ByRef FileHead As FILEHEADER)
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 12345
        .lngFileSize = .lngFileSize Xor 123456789
    End With
End Sub

Private Sub encryptHeaderInfo(ByRef InfoHead As INFOHEADER)
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr(Asc(Mid(InfoHead.strFileName, loopc, 1)) Xor 12)
        Else
            EncryptedFileName = EncryptedFileName & Chr(Asc(Mid(InfoHead.strFileName, loopc, 1)) Xor 23)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 123456789
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 123456789
        .lngFileStart = .lngFileStart Xor 123456789
        .strFileName = EncryptedFileName
    End With
End Sub

' *********** Compress / Extract  Files Functions ***********

Public Function extractFiles(ByVal FileType As srcFileType) As Boolean
    
Dim loopc As Long

Dim SourceFilePath As String
Dim OutputFilePath As String

Dim SourceFile     As Integer
Dim Handle         As Integer
Dim SourceData()   As Byte

Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case FileType
        Case BMP
            SourceFilePath = App.Path & SrcPath & "Graphics.pak"
            OutputFilePath = App.Path & "\Extract Files\Graphics\"
            
        Case MIDI
            SourceFilePath = App.Path & SrcPath & "Midi.pak"
            OutputFilePath = App.Path & "\Extract Files\Midi\"
        
        Case MP3
            SourceFilePath = App.Path & SrcPath & "Mp3.pak"
            OutputFilePath = App.Path & "\Extract Files\Mp3\"
        
        Case Else
            Exit Function
    End Select
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt FILEHEADER
    encryptHeaderFile FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Erase InfoHead
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
        
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each INFOHEADER befWO accessing the data
        encryptHeaderInfo InfoHead(loopc)
        
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        If InfoHead(loopc).lngFileSize < InfoHead(loopc).lngFileSizeUncompressed Then
            dataDecompress SourceData, InfoHead(loopc).lngFileSizeUncompressed
        End If
        
        'Get a free handler
        Handle = FreeFile
        
        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As Handle
        
        Put Handle, , SourceData
        
        Close Handle
        
        Erase SourceData
        
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    extractFiles = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function extractFile(ByVal FileType As srcFileType, ByVal fileName As String) As Boolean

Dim loopc As Long
Dim SourceFilePath As String
Dim OutputFilePath As String
Dim SourceFile As Integer
Dim SourceData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim Handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case FileType
        Case BMP
            SourceFilePath = App.Path & SrcPath & "Graphics.pak"
            OutputFilePath = App.Path & "\Extract Files\Graphics\"
            
        Case MIDI
            SourceFilePath = App.Path & SrcPath & "Midi.pak"
            OutputFilePath = App.Path & "\Extract Files\Midi\"
        
        Case MP3
            SourceFilePath = App.Path & SrcPath & "Mp3.pak"
            OutputFilePath = App.Path & "\Extract Files\Mp3\"
        
        Case Else
            Exit Function
    End Select
    
    'Make sure it's lower case
    fileName = LCase$(fileName)
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    encryptHeaderFile FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Get the file position within the compressed resource file
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each Info Header befWO accessing the data
        encryptHeaderInfo InfoHead(loopc)
        If Left$(InfoHead(loopc).strFileName, Len(fileName)) = fileName Then
            Exit For
        End If
    Next loopc
    
    'Make sure index is valid
    If loopc > UBound(InfoHead) Then
        Erase InfoHead
        Close SourceFile
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
    
    'Get the data
    Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
    
    'Decompress all data
    If InfoHead(loopc).lngFileSize < InfoHead(loopc).lngFileSizeUncompressed Then
        dataDecompress SourceData, InfoHead(loopc).lngFileSizeUncompressed
    End If
    
    'Get a free handler
    Handle = FreeFile
    
    Open OutputFilePath & InfoHead(loopc).strFileName For Binary As Handle
    
    Put Handle, , SourceData
    
    Close Handle
    
    Erase SourceData
    
    'Close the binary file
    Close SourceFile
        
    Erase InfoHead
    
    extractFile = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function
Public Function compressFiles(ByVal FileType As srcFileType) As Boolean
    
Dim SourceFilePath As String
Dim SourceFileExtension As String
Dim OutputFilePath As String
Dim SourceFile As Long
Dim OutputFile As Long
Dim SourceFileName As String
Dim SourceData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim FileNames() As String
Dim lngFileStart As Long
Dim loopc As Long
    
'Set up the error handler
'On Local Error GoTo ErrHandler
    
    Select Case FileType
        Case BMP
            SourceFileExtension = ".bmp"
            SourceFilePath = App.Path & "\..\Client\Graphics\"
            OutputFilePath = App.Path & SrcPath & "Graphics.pak"
            
            Debug.Print SourceFilePath
        Case MIDI
            SourceFileExtension = ".mid"
            SourceFilePath = App.Path & "\..\Client\Sounds\"
            OutputFilePath = App.Path & SrcPath & "Midi.pak"
        
        Case MP3
            SourceFileExtension = ".mp3"
            SourceFilePath = App.Path & "\..\Client\Sounds\"
            OutputFilePath = App.Path & SrcPath & "Mp3.pak"
        
        Case Else
            Exit Function
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase$(SourceFileName)
        
        'Search new file
        SourceFileName = Dir
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'FileHead.intNumFiles = MaxFiles
    
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'StWO file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'StWO the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        dataCompress SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.intNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
        'Once an InfoHead index is ready, we encrypt it
        encryptHeaderInfo InfoHead(loopc)
    Next loopc
    
    'Encrypt the FileHeader
    encryptHeaderFile FileHead
    
    '************ Write Data
    
    'Get all data stWOd so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'StWO the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
'Exit Function
'
'ErrHandler:
'    Erase SourceData
'    Erase InfoHead
'    'Display an error message if it didn't work
'    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Sub deleteFile(ByVal filePath As String)
Dim Handle As Integer
Dim data() As Byte
    
    'We open the file to delete
    Handle = FreeFile
    Open filePath For Binary Access Write Lock Read As Handle
    
    'We replace all the bytes in it with 0s
    ReDim data(LOF(Handle) - 1)
    Put Handle, 1, data
    
    'We close the file
    Close Handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill filePath
End Sub

