Attribute VB_Name = "Module2"
' #VBIDEUtils#************************************************************
' * Programmer Name  : VBDiamond
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           :
' * Date             : 03/11/2001
' **********************************************************************
' * Comments         : Shell Sort of an array of any type
' *
' * Shell Sort of an array of any type
' *
' **********************************************************************
Option Explicit
Public ORGorFIX As Boolean
Sub ShellSortAny(arr As Variant, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of any type
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Variant

   ' exit if it is not an arr
   If VarType(arr) < vbArray Then Exit Sub

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
 
   
   
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls
Do
     distance = distance - 1
    distance = distance / 3
For index = distance To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         
         Loop
         arr(index2) = value
      Next
Loop Until distance <= 1

If ORGorFIX = True Then
'******This will Fixed FIRST ARRAY PROBLEM**********
SWAP arr
Exit Sub
End If


End Sub
Private Sub SWAP(var As Variant)
Dim d As Integer
Dim tmpvar As Variant
d = 1
Do While var(d) < var(d - 1)
tmpvar = var(d - 1)
var(d - 1) = var(d)
var(d) = tmpvar
d = d + 1
Loop
End Sub

Public Function FindFast(var As Variant, wvar As Variant) As Long
'Algorithm FIND FAST PLACE FOR VARIANT in SORTED ARRAY LIST!!!
Dim index As Long
Dim xl As Long
index = CLng(UBound(var) / 2 + 0.1)
xl = CLng(index / 2 + 0.1)
Do
If wvar > var(index) Then
index = index + xl
ElseIf wvar < var(index) Then
index = index - xl
Else
Exit Do
End If
If xl = 1 Then Exit Do
xl = CLng(xl / 2 + 0.1)
Loop
FindFast = index
End Function
Sub ShellSortI(arr() As Integer, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of integers
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Integer

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortL(arr() As Long, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Long
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Long

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortS(arr() As Single, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Single
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Single

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortD(arr() As Double, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Double
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Double

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortC(arr() As Currency, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Currency
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Currency

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortDate(arr() As Date, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Date
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Date

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortStr(arr() As String, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of String
   '  Case-sensitive version
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As String

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortStrC(arr() As String, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of String
   '  Case-unsensitive version
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As String

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (StrComp(arr(index2 - distance), value, 1) = 1) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub

Sub ShellSortV(arr() As Variant, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Shell Sort of an array of Variant
   '---------------------------------------

   Dim index            As Long, index2 As Long
   Dim firstItem        As Long
   Dim inverseOrder     As Boolean
   Dim distance         As Long
   Dim value            As Variant

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)
   ' find the best value for distance
   Do
      distance = distance * 3 + 1
   Loop Until distance > numEls

   Do
      distance = distance / 3
      For index = distance + 1 To numEls
         value = arr(index)
         index2 = index
         Do While (arr(index2 - distance) > value) Xor inverseOrder
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
         Loop
         arr(index2) = value
      Next
   Loop Until distance = 1

End Sub
