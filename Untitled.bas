Attribute VB_Name = "__1"
Public Const TotalExperts As Integer = 34

Sub Macro1()

'Dim TotalExperts As Integer
'TotalExperts = 322
'Public Const TotalExperts As Integer = 322

'Dim i, j, k, arr(TotalExperts, 82) As Integer
Dim i, j, k, arr(TotalExperts, 81) As Integer
Dim b(0 To 4, 1 To 3) As Double

For i = 0 To 4
For j = 1 To 3
  b(i, j) = Sheet2.Cells(1 + i, j)
Next
Next

For i = 1 To TotalExperts
'For j = 0 To 81
For j = 1 To 81
  arr(i, j) = Sheet1.Cells(3 + i, 1 + j)
  Sheet4.Cells(i + 12, j) = arr(i, j)
Next
Next

'Dim maxR, minL, minLk(1 To 9, 1 To 9) As Double
'maxR = 0
'minL = 1

'For i = 1 To 9
'For j = 1 To 9
 ' minLk(i, j) = 1
 ' Sheet4.Cells(i, j) = minLk(i, j)
'Next
'Next

Dim maxRk(1 To TotalExperts), minLk(1 To TotalExperts) As Double
For k = 1 To TotalExperts
maxRk(k) = 0
minLk(k) = 1
Next

Dim maxR, minL As Double


For k = 1 To TotalExperts
maxR = 0
minL = 1
For i = 1 To 9
For j = 1 To 9

Sheet3.Cells((k - 1) * 10 + i, j) = arr(k, (i - 1) * 9 + j)

v = arr(k, (i - 1) * 9 + j)

If i <> j And b(v, 1) < minL Then minL = b(v, 1)
'If b(v, 1) < minL Then minL = b(v, 1)
If b(v, 3) > maxR Then maxR = b(v, 3)
If v < 5 Then
Sheet3.Cells((k - 1) * 10 + i, j + 10) = "(" & b(v, 1) & "," & b(v, 2) & "," & b(v, 3) & ")"
End If
Next
Next
maxRk(k) = maxR
minLk(k) = minL
Next

Dim arrL(1 To TotalExperts, 1 To 9, 1 To 9), arrM(1 To TotalExperts, 1 To 9, 1 To 9), arrR(1 To TotalExperts, 1 To 9, 1 To 9) As Double

For k = 1 To TotalExperts
For i = 1 To 9
For j = 1 To 9
v = arr(k, (i - 1) * 9 + j)
'arrL(k, i, j) = (b(v, 1) - minLk(i, j)) / (maxR - minL)
'arrM(k, i, j) = (b(v, 2) - minLk(i, j)) / (maxR - minL)
'arrR(k, i, j) = (b(v, 3) - minLk(i, j)) / (maxR - minL)

'arrL(k, i, j) = (b(v, 1) - minLk(k)) / (maxRk(k) - minLk(k))
'arrM(k, i, j) = (b(v, 2) - minLk(k)) / (maxRk(k) - minLk(k))
'arrR(k, i, j) = (b(v, 3) - minLk(k)) / (maxRk(k) - minLk(k))

arrL(k, i, j) = (b(v, 1) - minLk(k)) / (maxRk(k) - minLk(k))
arrM(k, i, j) = (b(v, 2) - minLk(k)) / (maxRk(k) - minLk(k))
arrR(k, i, j) = (b(v, 3) - minLk(k)) / (maxRk(k) - minLk(k))

Sheet3.Cells(1, 20) = "Lijk"
Sheet3.Cells(1, 30) = "Mijk"
Sheet3.Cells(1, 40) = "Rijk"

Sheet3.Cells((k - 1) * 10 + i, j + 20) = arrL(k, i, j)
Sheet3.Cells((k - 1) * 10 + i, j + 30) = arrM(k, i, j)
Sheet3.Cells((k - 1) * 10 + i, j + 40) = arrR(k, i, j)

Next
Next
Next

Dim arrSL(1 To TotalExperts, 1 To 9, 1 To 9), arrSR(1 To TotalExperts, 1 To 9, 1 To 9), X(1 To TotalExperts, 1 To 9, 1 To 9) As Double



For k = 1 To TotalExperts
For i = 1 To 9
For j = 1 To 9

v = arr(k, (i - 1) * 9 + j)
'arrSL(k, i, j) = b(v, 2) / (1 + b(v, 2) - b(v, 1))
'arrSR(k, i, j) = b(v, 3) / (1 + b(v, 3) - b(v, 2))

arrSL(k, i, j) = arrM(k, i, j) / (1 + arrM(k, i, j) - arrL(k, i, j))
arrSR(k, i, j) = arrR(k, i, j) / (1 + arrR(k, i, j) - arrM(k, i, j))

X(k, i, j) = (arrSL(k, i, j) * (1 - arrSL(k, i, j)) + arrSR(k, i, j) * arrSR(k, i, j)) / (1 + arrSR(k, i, j) - arrSL(k, i, j))

Sheet3.Cells(1, 50) = "Xijk"
Sheet3.Cells((k - 1) * 10 + i, j + 50) = X(k, i, j)

Next
Next
Next

Dim BNP(1 To TotalExperts, 1 To 9, 1 To 9) As Double
For k = 1 To TotalExperts
For i = 1 To 9
For j = 1 To 9

BNP(k, i, j) = X(k, i, j) * (maxRk(k) - minLk(k))  'minLk(k) +

'BNP(k, i, j) = minL + X(k, i, j) * (maxR - minL)

Sheet3.Cells(1, 60) = "BNPijk"
Sheet3.Cells((k - 1) * 10 + i, j + 60) = BNP(k, i, j)

Next
Next
Next

Dim a(1 To 9, 1 To 9), sum As Double
 For i = 1 To 9
 For j = 1 To 9
 
 sum = 0
 For k = 1 To TotalExperts
 sum = sum + BNP(k, i, j)
 Next
 a(i, j) = sum / TotalExperts
 If i = j Then a(i, j) = 0
 
 Sheet3.Cells(1, 70) = "Aij"
 Sheet3.Cells(i, j + 70) = a(i, j)
 Next
 Next
 
 Dim max As Double
 max = 0
 For i = 1 To 9
    sum = 0
    For j = 1 To 9
        sum = sum + a(i, j)
    Next
    If sum > max Then max = sum
 Next
  
 Dim D1(1 To 9, 1 To 9) As Double
For i = 1 To 9
For j = 1 To 9
a(i, j) = a(i, j) / max  '(max + 2)  '6
If (i = j) Then
D1(i, j) = 1 - a(i, j)
Else
D1(i, j) = -a(i, j)
End If

Sheet3.Cells(1, 80) = "D"
Sheet3.Cells(i, j + 80) = a(i, j)

Sheet3.Cells(1, 90) = "1-D"
Sheet3.Cells(i, j + 90) = D1(i, j)
Next
Next

Sheet3.Cells(1, 100) = "1-D____"ÿ’Û"

Sheet3.Range(Sheet3.Cells(1, 101), Sheet3.Cells(9, 109)).FormulaArray = "=MINVERSE(RC[-10]:R[8]C[-2])"

Sheet3.Range(Sheet3.Cells(1, 111), Sheet3.Cells(9, 119)).FormulaArray = "=MMULT(CC1:CK9,CW1:DE9)"


End Sub
'Sub Macro2()


    'Range("K1:S9").Select
    'Selection.FormulaArray = "=MINVERSE(RC[-10]:R[8]C[-2])"
    'Sheet3.Range(Cells(1, 11), Cells(9, 19)).FormulaArray = "=minverse(cells(1,1),cells(9,9))"
    'For k = 1 To TotalExperts
    'Sheet3.Range(Cells(1, 101), Cells(9, 109)).FormulaArray = "=MINVERSE(RC[-10]:R[8]C[-2])"
    'Next
'End Sub

