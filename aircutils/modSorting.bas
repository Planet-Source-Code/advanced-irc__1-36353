Attribute VB_Name = "modSorting"
Option Explicit
Option Compare Text

Public Type WndLDef
Wnd As Object
Col As Long
Indx As Integer
SrvNum As Integer 'Server number
Flags As Integer 'Window type
Ico As StdPicture
Title As String
End Type

Public Wnds() As WndLDef
Public WndsBK() As WndLDef

Private Const ERROR_NOT_FOUND As Long = &H80000000 ' DO NOT CHANGE, for internal usage only !

Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Public Sub SortArray(ByRef sArray() As WndLDef, Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim i          As Long   ' Loop Counter
   Dim j          As Long
   Dim iLBound    As Long
   Dim iUBound    As Long
   Dim iMax       As Long
   Dim sTemp      As WndLDef
   Dim distance   As Long
   Dim bSortOrder As Boolean
   
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   If Not iLBound = LBound(sArray) Then Exit Sub
   If Not iUBound = UBound(sArray) Then Exit Sub

   bSortOrder = IIf(SortOrder = SortAscending, False, True)
   iMax = iUBound - iLBound + 1
   
   Do
      distance = distance * 3 + 1
   Loop Until distance > iMax

   Do
      distance = distance \ 3
      For i = distance + iLBound To iUBound
         sTemp = sArray(i)
         j = i
         Do While (sArray(j - distance).Title > sTemp.Title) Xor bSortOrder
            sArray(j) = sArray(j - distance)
            j = j - distance
            If j - distance < iLBound Then Exit Do
         Loop
         sArray(j) = sTemp
      Next i
   Loop Until distance = 1

End Sub
