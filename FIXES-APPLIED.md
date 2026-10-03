# Đã áp dụng fixes — 7 decks SESSION-2 … SESSION-8

**Ngày:** 03/10/2026
**Nguồn:** các mục sửa trong `REVIEW-7-SESSION-DECKS.md` (review 233 slides)
**Kết quả:** **78/78 fix dạng text-layer đã áp dụng** + **6 phần tử bổ sung** đã thêm. Kiểm định `verify_review_fixes.py`: **63/63 checks OK**.

## Nội dung đã sửa

### Nhóm fix lặp lại (mọi deck)

| Nhóm | Nội dung | Số slide |
|---|---|---|
| Footer unify | Thống nhất footer "A2 Key for Schools Trainer" trên mọi slide sai format | 24 (S2) + 10 (S5) + 10 (S8) |
| "Write a letter" → "Write an email" | Đúng format Writing Part 6 của đề thi | nhiều slide S3–S7 |
| Thiếu người nhận | Chèn tên người nhận (Jane / James) vào "Write a letter to ___" | S6·32, S7·29 |
| Part 3/4 lẫn lộn (S7) | Thống nhất "Reading Part 4" cho Session 7 (objectives + mọi label) | sl. 7–10, 21–24, 35 |
| US → UK spelling | favourite, colour, theatre, cinema… | nhiều slide |
| Grammar/A2-accuracy | "meaning **as**", "match **to**", "Who is **the** receiver?", "take turns", "once only"… | nhiều slide |

### Fix theo slide (đặc trưng)

- **S2·13** — US cinema → UK cinema (BINGO row)
- **S2·18** — example fix: viết hoa đầu câu + dấu câu đầy đủ
- **S3·7** — retitle "Checking answers" (section rỗng → có mục đích)
- **S3·12, S4·9** — greeting / sign-off / email muddle sửa theo model
- **S3·15, S4·33** — rename slide có MODEL ANSWER chip
- **S4·2, S6·2** — 3 câu grammar hỏi sai → sửa thành câu đúng form
- **S4·10** — "the receiver" + câu hỏi tense + answer
- **S4·19, S4·21** — "match to", "the same meaning as"
- **S5·1, S5·4, S5·9** — typo + clarity, gap-fill instructions, "in London"
- **S6·3, S6·4** — state answer + reveal answer
- **S7·5 → S7-10**, **S8·1a, S8·3b, S8·3c** — script Listening S8: sửa 2/3 lỗi ("as large as **ours**" + ending logic); "Who will **we** do" đã xử lý ở bullet fix
- **S8·8** — rubric story đúng format (35+ words)

### 6 phần tử được thêm vào (add_review_elements.py)

1. **S3·3** — clone số section 03 → tạo đủ 01/02 (trước đó divider thiếu số)
2. **S3·18** — thêm dòng answer cho Q3 (trước để trống)
3. **S7·10** — điền đáp án gap-fill (owned / career)
4. **S7·23–24** — thêm rail ANSWERS: 19C earn · 20B hours · 21C worse · 22B young · 23A enough · 24B stop
5. **S8·16** — thêm strip đáp án Listening Part 3: 11B · 12B · 13A · 14B · 15A

## Kiểm định sau khi sửa

- `verify_review_fixes.py` → **63/63 checks OK**
- Toàn vẹn file: **233/233 slides**, transition **fade 233/233**, font **Baloo 2 + Nunito** nguyên vẹn
- Render PNG (LibreOffice) kiểm tra trực quan: s3-03, s3-18, s7-10, s7-23, s8-16 — không chồng lấn, đáp án khớp options

## Chưa sửa được (cần tài liệu gốc / nằm trong ảnh)

| Mục | Lý do | Cần gì để sửa |
|---|---|---|
| S4·24 answer key (A/B/C cả 3 tên, không đánh dấu) | Key không có sẵn, đáp án nằm trong ảnh | Answer key chính thức của sách |
| Kahoot link ("Let's play kahoot!!!!!!") | Không có URL thật | URL Kahoot của giáo viên |
| "turn to page 4 / page 5" (S2·19, S4·20) | Không có trang handout 4/5 trong deck/folder | File handout gốc |
| "WITH EDITABLE STROKE" in trên ảnh (S6·30) | Artefact nằm trong pixel ảnh | Re-export ảnh từ source |
| Độ legibility "KET WRITING GUIDE" | Ảnh quá nhỏ/đè nét | Ảnh gốc độ phân giải cao |
| "Piak" spelling, hình "10 seconds" | Bên trong ảnh | Ảnh gốc chỉnh sửa |
| Gộp/bỏ slide trùng, chuyển mục giảng dạy (Reading P5, Listening P4/P5 chưa dạy) | Thay đổi cấu trúc chương trình | Quyết định sắp xếp giáo trình |
| Audio files | Deck không embed/link audio nào | File audio của Trainer |

## Chạy lại các script

```bash
pip install python-pptx
python review-fix/dump_all_texts.py          # extract text 233 slides -> all_texts.json
python review-fix/apply_review_fixes.py      # pass 1: run-level fixes
python review-fix/apply_review_fixes2.py     # pass 2: cross-run offset replacements
python review-fix/add_review_elements.py     # 6 phần tử bổ sung (S3, S7, S8)
python review-fix/verify_review_fixes.py     # 63/63 checks
```

Lưu ý: các script đọc/ghi trực tiếp `SESSION-*-A2-KET-TRAINER_Yola.pptx` — hãy backup trước khi chạy lại. `fix_log.json` ghi log từng fix (pass 1: 53 mục; pass 2 vá tiếp 23 mục; 2 mục cuối sửa inline).
