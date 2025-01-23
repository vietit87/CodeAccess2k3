**CODE SQL Access**
Đặt trong cặp "DoCmd.SetWarnings False" và "DoCmd.SetWarnings True" để hàm chạy mà không cần hiện bảng thông báo, Hàm SQL cần làm thử trước, chạy chính xác trong query rồi mới đưa vào SQL của Form

**1. Hàm UPDATE SQL của Form**

```
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE HUYEN SET HUYEN.NHO = null"
DoCmd.SetWarnings True
```

