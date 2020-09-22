Attribute VB_Name = "modEncryption"
'####### Remote Control Encryption System #######

Function rmt_Encrypt(ByVal S As String) As String
    Dim C As Long
    Dim M As String
    Dim Enc As Long
    rmt_Encrypt = S
    If rmt_Encrypted(S) Then Exit Function
    rmt_Encrypt = Chr(1)
    For C = 1 To Len(S)
        M = Mid(S, C, 1)
        Enc = Enc + ((CLng(Asc(M)) * 2) * CLng(Asc(M)))
        If Enc > (2 ^ 63) Then Enc = Enc \ 2
        rmt_Encrypt = rmt_Encrypt & Chr((Enc Mod 255) + 1) & Chr(Enc Mod 255)
    Next
    rmt_Encrypt = rmt_Encrypt & Chr(2)
End Function

Function rmt_Encrypted(ByVal S As String) As Boolean
    If ((Left(S, 1) = Chr(1)) And (Right(S, 1) = Chr(2))) Then rmt_Encrypted = True
End Function

Function rmt_Compare(ByVal Clean As String, ByVal Encrypted As String) As Boolean
    If Encrypted = rmt_Encrypt(Clean) Then rmt_Compare = True
End Function
