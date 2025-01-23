**BỘ CODE ACCESS 2003**

Bộ Code lưu trữ cá nhân, thường sử dụng khi viết phần mềm trên Access. Đối với Access mới hơn 2003, chế độ bảo mật cao hơn, có thể code sẽ không chạy đúng.

**A. CODE Query**

1. Tạo 2 trường năm sinh cho nam và nữ, code động khi chỉ có tháng, ngày tháng hoặc chỉ có năm sinh

_Năm sinh cho nam:_

`NS_NAM: IIf([GIOITINH]=1,IIf(IsNull([NGAYSINH]) And IsNull([THANGSINH]),[NAMSINH],IIf(Not IsNull([NGAYSINH]) And Not IsNull([THANGSINH]),IIf(Len([NGAYSINH])=1,"0" & [NGAYSINH],[NGAYSINH]) & "/" & IIf(Len([THANGSINH])=1,"0" & [THANGSINH],[THANGSINH]) & "/" & [NAMSINH])))`

_Năm sinh cho nữ:_

`NS_NU: IIf([GIOITINH]=2,IIf(IsNull([NGAYSINH]) And IsNull([THANGSINH]),[NAMSINH],IIf(Not IsNull([NGAYSINH]) And Not IsNull([THANGSINH]),IIf(Len([NGAYSINH])=1,"0" & [NGAYSINH],[NGAYSINH]) & "/" & IIf(Len([THANGSINH])=1,"0" & [THANGSINH],[THANGSINH]) & "/" & [NAMSINH])))`

**B. CODE Form**


**C. CODE SQL**


**D. CODE Modules**
