# Heading trong Word bị mảng đen biết phải làm sao?

Một ngày đẹp trời Khoá luận của bạn mở lên và Heading bị đen lại.

![](https://lh5.googleusercontent.com/proxy/VrxUZO5MOT1AIUt-xSgn3RBBbURqdwnrcoBiDzGUwfgyCHpAjkv23NN07U8NnZE4Tt2NN1n8CYJza8u-RCLsM36hCgsV31dH9uCPnVA=s0-d)

Tốt nhất đập đi làm lại... Lần 1. OK

Lần khác mở lên, đen tiếp... Đập làm lại...

Đã bị thì nó sẽ bị hoài...

Cách giải quyết như sau:

**Bước 1:** Lưu lại 1 file mới cho chắc ăn

**Bước 2:** Trên file mới đang mở tạo Macro như sau:

![](https://3.bp.blogspot.com/-Mb1qfPPmhJc/W0cbT4ZkLWI/AAAAAAAAD40/wls-DOBROlwB4BAOLaMr1whcKZnP-3qAQCLcBGAs/s640/Screen%2BShot%2B2018-07-12%2Bat%2B4.10.10%2BPM.png)

Chọn tiếp Dấu cộng\(+\) 

![](https://2.bp.blogspot.com/-FHSRBxsTgKk/W0cbXHeahCI/AAAAAAAAD44/JfcgGkY-pSINX6aGnKnoFtUGmFa2m5RBACLcBGAs/s640/Screen%2BShot%2B2018-07-12%2Bat%2B4.10.25%2BPM.png)

 Cửa sổ mới như sau:

![](https://3.bp.blogspot.com/-d6v3looYDPU/W0cbXW74FKI/AAAAAAAAD5A/r49quUJ6gdUL1d66gf1CZo3NKTDtTTX1wCLcBGAs/s640/Screen%2BShot%2B2018-07-12%2Bat%2B4.10.37%2BPM.png)

 Dán đoạn code sau vào:

> Sub RemoveBlackBox\(\)
>
> '
>
> ' RemoveBlackBox Macro
>
> '
>
> '
>
> For Each templ In ActiveDocument.ListTemplates
>
> For Each lev In templ.ListLevels
>
> lev.Font.Reset
>
> Next lev
>
> Next templ
>
> End Sub

  
 Cuối cùng nhấn Run \(Hình tam giác ở góc trái màn hình\).

![](https://1.bp.blogspot.com/-bnLXQTZDE_U/W0cbXTN5fwI/AAAAAAAAD48/kqMjJzo0-OI-MfIsePaNzjQVuS73VVyfACLcBGAs/s640/Screen%2BShot%2B2018-07-12%2Bat%2B4.10.56%2BPM.png)

![](https://3.bp.blogspot.com/-GeN_vm9gPMg/W0cbXxCDwII/AAAAAAAAD5E/KrOoQNRv0S4imuRbYPWVFyUJ1GFmknW6QCLcBGAs/s1600/Screen%2BShot%2B2018-07-12%2Bat%2B4.11.09%2BPM.png)

Chờ đoạn code chạy xong thì tắt cửa sổ đi là xong!

## 

