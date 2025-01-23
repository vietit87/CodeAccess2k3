**CODE Form Access**

**1. Hàm LOOKUP**

_Tìm Năm sinh theo điều kiện cố định là > 2000_

```
DLookup("NAMSINH", "DS_TV", "DS_TV.NAMSINH>2000")
```

_Tìm Năm sinh theo điều kiện của giá trị nhập vào trên Form_

```
DLookup("NAMSINH", "DS_TV", "DS_TV.NAMSINH ='" & NAMNHAP & "'") = -1
```

**2. Hàm COUNT, LOOKUP, MIN, MAX**

_Đếm số lượng thành viên có trong hộ_

```
DCount("MATV", "DS_TV", "DS_TV.MASO ='" & MASO & "'") > 1
```

**4. Hàm MIN, MAX**

_Tìm giá trị STT lớn nhất của thành viên trong hộ_

```
DMax("STT", "DS_TV", "DS_TV.MASO like '" & Form_NAME.MASO & "'") 
```

