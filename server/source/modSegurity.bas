Attribute VB_Name = "modSegurity"
Option Explicit
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
