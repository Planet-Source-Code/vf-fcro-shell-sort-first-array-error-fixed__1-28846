Attribute VB_Name = "Module1"
' #VBIDEUtils#************************************************************
' * Programmer Name  : VBDiamond
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           :
' * Date             : 03/11/2001
' **********************************************************************
' * Comments         : Bubble Sort of an array of any type
' *
' * Bubble Sort of an array of any type
' *
' **********************************************************************
Option Explicit

Sub BubbleSortAny(arr As Variant, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of any type
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Variant

   ' exit if it is not an arr
   If VarType(arr) < vbArray Then Exit Sub

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortI(arr() As Integer, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of integers
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Integer

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortL(arr() As Long, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Long
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Long

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortS(arr() As Single, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Single
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Single

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortD(arr() As Double, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Double
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Double

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortC(arr() As Currency, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Currency
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Currency

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortDate(arr() As Date, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Date
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Date

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortStr(arr() As String, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of String
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As String

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortStrC(arr() As String, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of String
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As String

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (StrComp(value, arr(index + 1), 1) = 1) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub

Sub BubbleSortV(arr() As Variant, Optional ByVal numEls As Variant, Optional ByVal descending As Variant)

   '---------------------------------------
   '  Bubble Sort of an array of Variant
   '---------------------------------------

   Dim index            As Long
   Dim firstItem        As Long
   Dim indexLimit       As Long, lastSwap As Long
   Dim inverseOrder     As Boolean
   Dim value            As Variant

   ' account for optional arguments
   If IsMissing(numEls) Then numEls = UBound(arr)
   If IsMissing(descending) Then descending = False
   inverseOrder = (descending <> False)

   firstItem = LBound(arr)

   lastSwap = numEls
   Do
      indexLimit = lastSwap - 1
      lastSwap = 0
      For index = firstItem To indexLimit
         value = arr(index)
         If (value > arr(index + 1)) Xor inverseOrder Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
         End If
      Next
   Loop While lastSwap

End Sub
