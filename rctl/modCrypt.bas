Attribute VB_Name = "modCrypt"
Function Encrypt(ByVal S As String) As String
    Dim C As Long
    Dim M As String
    Dim Enc As Long
    Encrypt = S
    If Encrypted(S) Then Exit Function
    Encrypt = Chr(1)
    For C = 1 To Len(S)
        M = Mid(S, C, 1)
        Enc = Enc + ((CLng(Asc(M)) * 2) * CLng(Asc(M)))
        If Enc > (2 ^ 63) Then Enc = Enc \ 2
        Encrypt = Encrypt & Chr((Enc Mod 255) + 1) & Chr(Enc Mod 255)
    Next
    Encrypt = Encrypt & Chr(2)
End Function

Function Encrypted(ByVal S As String) As Boolean
    If ((Left(S, 1) = Chr(1)) And (Right(S, 1) = Chr(2))) Then Encrypted = True
End Function

Function Compare(ByVal Clean As String, ByVal Encrypted As String) As Boolean
    If Encrypted = Encrypt(Clean) Then Compare = True
End Function
