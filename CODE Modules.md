**CODE Modules Access**

**1. Hàm định dạng toàn bộ chữ Hoa**
```
'Chu toan hoa
Public Function Hoa(ST) As String
On Error Resume Next
    Dim temp As String
    If IsNull(ST) Then
      Exit Function
    Else
    temp = CStr(UCase(ST))
 Hoa = temp
End If
End Function
```

**2. Hàm định dạng chữ Hoa đầu mỗi từ**

```
'Hoa dau tu
Public Function HDTu(ST) As String
On Error Resume Next
    Dim temp, t As String
        If IsNull(ST) Then
            Exit Function
      Exit Function
    Else
    temp = CStr(LCase(ST))
    A = Split(temp, " ")
    For i = 0 To UBound(A)
        A(i) = UCase(Left(A(i), 1)) & Right(A(i), Len(A(i)) - 1)
    Next
HDTu = Join(A, " ")
End If
End Function
```
**3. Hàm định dạng chữ Hoa đầu mỗi đoạn**

```
'Hoa dau chuoi
Public Function HDChuoi(ST) As String
On Error Resume Next
    Dim temp, Hoa, A As String
    If IsNull(ST) Then
      Exit Function
    Else
    temp = CStr(LCase(ST))
    Hoa = Left(temp, 1)
    A = UCase(Hoa) & Right(temp, Len(temp) - 1)
    HDChuoi = A
End If
End Function
```
