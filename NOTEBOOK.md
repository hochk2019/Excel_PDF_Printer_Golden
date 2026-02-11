# NOTEBOOK

## Muc tieu
- Thiet ke lai giao dien app de ro rang, de thao tac, thong nhat phong cach.
- Goi y tinh nang va UI de hoan thien san pham.
- Tach bớt tien ich sang module rieng va them test.
- Gom cac thao tac loc/preset/thu muc vao `app_actions.py` de giam dong file chinh.
- Them tooltip va hieu ung cho nut bam qua `ui_effects.py`.
- Bo sung vien mong cho nut bam (highlightthickness=1) de noi bat hon.
- Them focus ring ro rang khi dung Tab cho nut/entry/combobox.
- Tang do day vien nut (2px) va focus ring (3px) de noi bat hon.
- Danh dau phien ban V4.1, them changelog va designer.
- Mo rong panel ben trai va them nut "Xóa hết" de thao tac nhanh.
- Thay logo sang "Logo moi 1.png" va uu tien hien thi trong header.
- Thu nghiem logo "Logo moi 2- size bé.png" de net hon.
- Doi ten app va file phat hanh: "Excel & PDF Print V4.1- Golden Logistics".
- Cap nhat icon ung dung moi (app_icon_v41) cho .exe va khi chay.
- Tao icon hien dai moi (app_icon_v41_modern) va cap nhat icon .exe.
- Dua thong tin designer len header ben phai de luon hien ro.

## Boi canh
- Ung dung desktop (Tkinter) cho in/xuat PDF tu Excel/Word/PDF.
- UI hien tai gom nhieu khung chuc nang trong mot cua so.

## Quyet dinh / Huong di
- Ap dung phong cach toi gian (Swiss/minimal), mau logistics (xanh + cam), font Segoe UI.
- To chuc lai bo cuc theo cac khoi: Nguon > Loc > Danh sach | Thiet lap > Dau ra > Thao tac > Nhat ky.
- Bo emoji khoi nhan dien nut bam, tang do tuong phan.
- Tach ham tien ich sang `app_utils.py` va tach handler in/xuat sang `print_handlers.py`.
- Them test cho module moi trong `tests/`.

## Tien do
- [x] Tao `app_utils.py` + test `tests/test_app_utils.py`
- [x] Tao `print_handlers.py` + test `tests/test_print_handlers.py`
- [x] Cap nhat UI trong `print_excel_first_page_pro4_fix.py`
- [x] Tao `app_actions.py` + test `tests/test_app_actions.py`
- [x] Tao `ui_effects.py` + test `tests/test_ui_effects.py`
- [x] Chay test va kiem tra dem dong (file chinh < 800 dong)
- [x] Tong hop goi y tinh nang/UI (trong phan tra loi)
