# KET Session Yola — 7 decks review

Review nội dung 7 decks `SESSION-2` … `SESSION-8` (A2 KET Trainer, Yola):
tìm chỗ không ổn / không logic / gây khó hiểu cho học sinh và đề xuất fix.

- Chi tiết đầy đủ: xem `REVIEW-7-SESSION-DECKS.md`
- 233 slides, đã check text layer + render PNG + OCR + đối chiếu format Cambridge A2 Key 2020.

## Cấu trúc
- `SESSION-*-A2-KET-TRAINER_Yola.pptx` — 7 decks (**đã cập nhật bản fix 03/10/2026**)
- `REVIEW-7-SESSION-DECKS.md` — báo cáo review slide-by-slide + checklist sửa
- `FIXES-APPLIED.md` — tổng kết các fix đã áp dụng + phần chưa sửa được (cần tài liệu gốc)
- `review-fix/` — scripts áp dụng & kiểm định fix (python-pptx) + `fix_log.json`

## Trạng thái hiện tại ✅
- **78/78 fix dạng text-layer đã áp dụng** + 6 phần tử bổ sung (S3·3, S3·18, S7·10, S7·23–24, S8·16)
- Kiểm định: 63/63 checks OK — 233/233 slides, fade transition + font Baloo 2 / Nunito nguyên vẹn
- Chưa sửa được (cần answer key / audio / ảnh gốc / URL Kahoot / handout p.4–5): xem `FIXES-APPLIED.md`

## Tiếp tục develop
- Các mục còn lại nằm trong ảnh hoặc cần tài liệu gốc — xem bảng "Chưa sửa được" trong `FIXES-APPLIED.md`.
- Nếu chỉnh thêm, chạy scripts trong `review-fix/` (backup PPTX trước khi chạy lại).
- Mỗi nhóm fix nên là 1 commit riêng (vd: `fix S8-16 script`, `fix writing letter->email`, `add answer keys S4-24...`).
