# Content review — Yola A2 KET Trainer decks, Sessions 2–8

**Scope:** `SESSION-2` … `SESSION-8` (`.pptx`) — 233 slides in total.
**Purpose:** find content that is illogical, confusing, inaccurate or otherwise problematic for A2 students, and give concrete fixes.

## How this review was done (so you can re-check it)

1. Extracted the **text layer** of every slide with `python-pptx` (all 233 slides, including grouped shapes and tables).
2. **Rendered every slide to PNG** via PowerPoint and **OCR'd all 233 slides**, so that text baked into pictures (exam tasks, answer keys, notices, worksheets) could be reviewed too — a large share of the real content is images, not editable text.
3. Pixel-level checks (colour, ink density, glyph shape) on slides where the answer marking was not readable by OCR.
4. Cross-checked exam-format claims against the Cambridge A2 Key (2020) format.
5. File-level audit: slides, hidden slides, speaker notes, embedded/linked audio.

**Verified facts about the files themselves**

- **No speaker notes anywhere.** All 233 slides have an empty notes page (only the slide number) ⇒ **no trainer instructions and no answer keys in the notes**.
---

## Executive summary — the 10 things I would fix first

| # | Where | Problem | Fix |
|---|-------|---------|-----|
| 1 | S8·16 (script) | The Listening Part 3 script students must read and underline contains **real English errors**: “*So who will we do, Tim?*”, “*it’s not as large as **ours***”, and the ending contradicts the story (“party at Jack’s house” → “*I thought I would have the party in the garden*”). | Correct the script (§S8-3) or use the published script unedited. |
| 2 | S3/S4/S6/S7 (Writing Part 6) | Slides say **“Write a letter …”** while the task and the exam say **email**; **two slides are missing the recipient**: “*Write a letter to **using the structure below***” (S6·32, S7·29). | “letter” → “email”; insert the name (Jane / James). |
| 3 | S4·28, S6·31 | “*What tense do you use to write this email? → **Future tense***” — misleading: the reply needs present/present continuous (“I’m very excited”), *would like* and future forms. | Ask “Which tenses will you use? Give an example of each.” or “How many things must you include?” (as S7·28 does). |
| 4 | S5·11 | “*fill in the note with **a single words***” — grammar error **and** it contradicts the deck’s own answer key (60, 40, 7, Saturday). | “Write **one word, or a number/a date/a time** in each gap”, + “Check your spelling.” |
| 5 | S4·24, S3·20–21, S6·13·25·27, S7·10 | **“Let’s check:” / exam-practice slides with no answer marking.** On S4·24 every question row shows A, B and C under the three names and nothing marks the right one (verified at pixel level: all letters plain black, no highlight/bold/ring). | Show the correct answer (bold / coloured / tick) or add an answer strip. |
| 6 | S3 (2 sections), S4·17, S6·22, S7·19 | **Empty sections**: “Vocabulary Game” is announced but has **no game** in 4 decks; “Warm up” is announced with **no activity** in S3 and S7; S4 and S6 have **two** opening sections (Warm up + STOP THE BUS). | Add the activity or delete the divider. |
| 7 | S7 throughout | **Part number chaos**: “Reading **Part 4**” (sl. 7, 21, 23, 24, 35) vs “Reading **part 3**” / “Part 3: How to do / Practice” (sl. 8, 9, 10, 22). Objectives (sl. 2) also say “Reading Part 3” (copy-paste from Session 6). | Session 7 = **Reading Part 4**; fix the objective and all labels. |
| 8 | S4·32–33, S6·35–36, S7·32–33 | “Before writing” **gives the complete plan**, and the next slide, labelled **“Your turn”, already contains the finished model story**. No student production; the label is misleading. | “Your turn” = task only; move the model to a separate “Model answer” slide shown *after* students write. |
| 9 | S2·13–14, S2·19, S4·20 | Dead references: the **Kahoot link is missing** (“Let’s play kahoot!!!!!!”), **“Student A: turn to page 4 / Student B: turn to page 5”** — no such pages exist in the deck/folder, and the **BINGO grid (S2·13) has no instructions** anywhere. | Put the Kahoot URL in the notes, attach the pages, or replace with in-deck content. |
| 10 | **S6·30** | Visible **editing artefact in the imported image: “WITH EDITABLE STROKE”** printed on the slide. | Re-export the image from source without the layer name. |

**Tóm tắt nhanh cho giáo viên (VN)**

- 3 lỗi nặng nhất: (1) script Listening S8 sai ngữ pháp/lô-gíc; (2) nhiều slide “Let’s check” **không có đáp án**; (3) Session 7 ghi sai số Part (Part 3/Part 4 lẫn lộn).
- Lỗi lặp lại: “Write a **letter**” thay vì “email”; “Write a letter to ___” **thiếu tên người nhận**; “Who is a receiver?” (thiếu *the*); “the same meaning **with**” (đúng: *as*).
- Không có **file audio** và **không có ghi chú / answer key** trong speaker notes → giáo viên phải tự chuẩn bị.
- 4 buổi (S3, S4, S6, S7) đều dạy Writing Part 6&7, trong khi **Reading Part 5, Listening Part 4, Listening Part 5** chưa buổi nào dạy.

---

## A. Issues that appear in several / all decks

### A1. No answer keys, no trainer notes, no audio
- Every deck’s notes pages are empty; no audio file is embedded or linked (checked at file level).
- Tasks that therefore cannot be checked by students or by a substitute teacher: **S3·20–21, S4·23–24, S5·15·17, S6·12–13·25·27, S7·10·22, S8·16·18**.
- **Fix:** one “Answer key” slide per task, or put the key in the speaker notes (invisible to students), and name the audio file on each listening slide.

### A2. Writing Part 6 templates: inconsistent, and partly wrong

| Slide | Text as printed | Problem |
|---|---|---|
| S4·29 | “Write **a letter** to Bobby using the structure below” | task and exam say **email** |
| S6·32 | “Write a letter **to** using the structure below” | **missing “Jane”** |
| S7·29 | “Write a letter **to** using the structure below” | **missing “James”** |
| S3·26 | “Hi, **Tim!**” … closing “**Let me know asap,**” | should be “Hi Tim,”; “Let me know asap” is the *asker’s* phrase and very informal for a reply → “Write soon, / Best wishes,” |
| S4·29 | “**Ask Bobby to wait for you at the airport at what time**” | muddled; the task says “suggest a time to meet at the airport” |
| S6·32 | “Recommend a good restaurant and **their** famous dishes” | “their” → “its” |
| S3 vs S4/S6/S7 | two different sign-offs taught for the same task type (“Let me know asap,” vs “I am looking forward to seeing you.”) | teach one email frame + one sign-off consistently (the model answer on S3·26 uses “Best wishes,” — a third version) |

Also **“Who is a receiver?”** (S4·28, S6·31, S7·28) → “Who is **the** receiver?” / “Who will receive the email?” (3 occurrences).

### A3. Grammar / wording errors in rubrics (students copy these)

| Location | Printed | Should be |
|---|---|---|
| S2·5 | “You **take turn** to tell a story … that has 3 words provided”; “Each set of words must be used once only” | “You **take turns** to tell a story (30 seconds or more) **using the three words given**”; “Each set of three words can be used only once.” |
| S2·19, S5·6 | “**Take turn** to ask and answer”; “Take turn to read for **you friend**” | “Take turns…” ; “for **your** friend” |
| S4·6, S6·6 | “Write down 1 **animal starts** with letter …” | “…one animal **that** starts with the letter …” |
| S4·21, S8·5 | “has the same meaning **with**” | “the same meaning **as**” |
| S5·11 | “a **single words**” | “one word (or a number / a date / a time)” |
| S5·13 | “**Why to give?**” | “Why give a leaving present?” |
| S5·24 | “friends have just arrived **at** London … **Suggest them** the best way” | “arrived **in** London… **Suggest to them / give them advice on**…” |
| S6·35 | “sailing in a small boat **in** a sunny day” | “**on** a sunny day” (the model story two slides later is correct) |
| S7·10 | “appointment: same meaning **with** meeting” | “…the same meaning **as** meeting” |
| S8·30 | “Be careful **for** distractors” / “Listen to understand not to choose the answer” | “Be careful **of / about** distractors” / “Listen to understand, don’t just listen for the answer” |
| S2·22 | “Watch out for distractors (wait, but,…)” | Distractors are the *wrong options*; “but/wait/actually” are the signals → “Listen for words like *but, actually, oh wait* — they usually introduce the real answer.” |

### A4. US / UK spelling mixed (Cambridge = British English)
- “**Movie theater**” (S2·13), “**favorite room**” (S2·20), “**Information center**” (S6·15) — while the same decks use *cinema*, *centre*, *favourite* (e.g. the Hair-Museum text on S6·23 uses “tourist information centre”). Choose British English throughout.
- S4·14 “A CHEMIST”: in British English the **shop** is *the chemist’s / pharmacy*; “a chemist” is the person, and the picture shows the shop → label it “A CHEMIST’S”.

### A5. “Let’s check:” slides where nobody can check
- **S4·24** — each of the 7 question rows carries **A**, **B** and **C** under the three names (Naomi / Lisa / Mary Lu); nothing marks the correct one. Verified with pixel analysis: all three letters are plain black, same weight, and there is no highlight, ring or box in any of the 7 rows. The three texts are also headed with **names instead of the exam’s A/B/C labels**, even though S4·19 teaches that Part 2 gives “three short texts”.
- **S3·20–21** — notices, questions and options are shown, but no option is bolded/highlighted/lettered as the answer (ink-density check: all options 14–17 % ink, i.e. the same weight).
- **S6·13** (“Please do this question!!!”) and **S7·10** show questions but never give the answers; on S6·13 the question is the last item before a new section, so students never find out.
- **Fix:** give every “Let’s check:” slide a visible key (bold the correct option, or add a small answer strip).

### A6. Structural repetition and dead dividers

| Deck | Problem |
|---|---|
| S3 | “Warm up” divider (sl. 4) with **no warm-up activity**; “Vocabulary Game” divider (sl. 14) with **no game** |
| S4 | **Two opening sections**: “Warm up” (sl. 4) *and* “STOP THE BUS!!” (sl. 5); “VOCABULARY GAME” (sl. 17) with no game |
| S6 | Same double opening (sl. 4–5); “VOCABULARY GAME” (sl. 22) with no game |
| S7 | “Warm up” (sl. 4) with no activity; “VOCABULARY GAME” (sl. 19) with no game |
| S4 vs S6 | Identical “STOP THE BUS” warm-up (same 3 prompts, no letter, no scoring rules) |
| all | The **same exam-format tables** appear verbatim 4× (Reading & Writing: S3·6, S4·8, S6·8, S7·6) and 3× (Listening + Speaking: S2·8–9, S5·8–9, S8·7–8) — each time consuming a numbered section of the lesson |
| S5 vs S8 | The whole Speaking Part 2 block is nearly identical (“I don’t know what to say!”, WHAT/WHERE/WHEN/WHY/WHO, “Two teams — pick 1 box — 30 s”) |
| S3/S4/S6/S7 | The Writing Part 7 sequence is the same four-slide template every time (instruction → “Before writing” → plan → model) |
| S5·20/S8·21 | Same Speaking Part 2 “Phase 1” slide; S5·26 = S8·28 |

*Duplicated content inside one deck:* S3·17 and S3·18 (identical email twice), S6·23 and S6·24 (identical reading text twice), S6·36 and S6·39 (identical model story), S7·20 and S7·21 (identical text twice), S7·23 and S7·24 (identical exam-practice slide twice), S8·13–14–15 (identical question set three times).
→ **Fix:** replace duplicates with the missing pieces (answer keys, a real vocabulary game, a differentiated speaking focus).

### A7. Branding / consistency details
- Footer convention changes mid-course: **“KET · TEST 2 / 3 / 4”** (S2, S5, S8) vs **“KET · SESSION n”** (S3, S4, S6, S7).
- **S3·3** “ONLINE CLASS RULES” shows only **03** and **04** — the 01/02 numbers are missing, so two of the four rules are unnumbered.
- “**Revised** Exam Format from 2020” (S3·6, S4·8, S6·8, S7·6) reads as dated for a 2026 course; the 2020 format is the current one.
- Punctuation noise: “Let’s play kahoot!!!!!!” (S4·20), “Understanding the questions is a MUST!!!” (S4·20, S6·11), “Please do this question!!!” (S6·13). Also “A MUST” is not student-friendly English → “essential”.
- Section-label spacing is inconsistent: “SECTION 04 / 07” vs “SECTION 10 / 10” vs “SECTION IO / 10”.

### A8. Third-party images imported from other worksheets
Several model-answer / recap slides are screenshots from other materials. They carry their own wording, their own answers and sometimes their own errors, and they compete with the deck’s own structure box:
- S3·25 “Part 6 WRITING TUTOR … **Make sure WHO you will send the message to.**” (awkward) and tip 6 about writing on “a different piece of paper … answer sheet”.
- S3·26 / S3·32 model-answer sheet containing the typo “**White 25 words or more.**”
- S4·36 “KET WRITING GUIDE” — partly illegible at slide size.
- **S6·30 “WITH EDITABLE STROKE”** printed on the writing prompt (editing artefact).
- S6·39 / S7·36 “How to Write a Story” worksheets; **S7·36 shows a different model story (Mai / Jack)** from the Tom-&-Mary story the class has just written → confusing.
→ **Fix:** keep the deck’s own template as the single source of truth; re-export images cleanly or rebuild those two templates as native slides.

---

## B. Session-by-session findings
(Severity: **H** = fix before next class, **M** = fix when you next edit, **L** = polish.)

### Session 2 — Listening Part 1 + Speaking Part 1 (24 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S2-1 | 5 | M | “You take turn to tell a story … that has 3 words provided” + “Each set of words must be used once only” — two unclear sentences; “take turn” is ungrammatical | “Take turns to tell a story (30 seconds or more) using the three words given. Each set can be used only once.” |
| S2-2 | 5–6 | M | The three-word sets live only in a picture on slide 6 and are **not labelled on the instruction slide**; please re-check that every set really has exactly 3 words (the image could not be read reliably) | Put the 4 sets on the instruction slide with clear “Set 1/2/3/4” labels, 3 words each |
| S2-3 | 11–14 | M | **BINGO (sl. 12–13) has no rules at all** — students see a BINGO card and a 5×5 word grid with no instruction (who reads? when do we shout? how do we win?) | Add one line: “Teacher reads the words one by one; cross out what you hear; the first to complete a line shouts BINGO!” |
| S2-4 | 13 | L | The word grid mixes “interesting / interested”, “22nd / 33rd”, “80 %”, and uses **Movie theater** (US) | Keep the minimal pairs, replace “Movie theater” with “cinema”, and tell students the grid contains distractors |
| S2-5 | 18 | M | Phase description: “2 topical questions and 1 question that requires candidates to give **long answer**” (grammar); “For example: “where are you from?” **Or** “how old are you?”” (capital *Or* mid-sentence); “Tell me about your bedroom” (no full stop) | “…asks each candidate two questions on the same topic and one that needs a **longer** answer.”; fix the punctuation of the examples |
| S2-6 | 19 | H | “Student A: turn to **page 4** / Student B: turn to **page 5**” — those pages do not exist in the deck (no student handout in this folder) → the pair-work activity cannot run | Insert the two question pages at the end of the deck (or as a separate PDF), or replace with on-slide questions |
| S2-7 | 20 | H | A **speaking** lesson asks students to **write** “a short paragraph (**40 words max**)”: wrong modality, and “40 words max” contradicts “practice giving a long answer” | Make it oral (talk/record for 40 seconds); drop the word limit; if you want a written step, say “make notes (max 40 words), then answer aloud” |
| S2-8 | 20 | L | “favorite room” (US spelling) | “favourite room” |
| S2-9 | 22 | L | “Watch out for distractors (wait, but,…)” — labels the *signal words* as “distractors” | “Listen for words like *but, actually, oh wait* — they usually introduce the right answer.” |
| S2-10 | 1 / footer | L | Footer is “KET · TEST 2” while S3/S4/S6/S7 use “KET · SESSION n” | Pick one convention for all decks |

### Session 3 — Reading Part 1 + Writing Part 6 & 7 (33 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S3-1 | 3 | L | Class rules show only “03” and “04”; rules 1–2 have no numbers | Add “01” / “02” (copy the S2·3 layout) |
| S3-2 | 4 | M | Section “01 / 09 Warm up” has **no warm-up activity** (slide 5 starts section 02) | Add a short warm-up or delete the section |
| S3-3 | 8, 9, 13 | M | Vocabulary slides with **no answer reveal**, whereas slides 10–12 show “✓ INVITE / RECEPTIONIST / COACH”. Slide 13 (“VOCABULARY ____”) never shows its word at all | Add the ✓ answer line to 8, 9 and 13 (check what slide 13’s picture is meant to be) |
| S3-4 | 8 | L | Slide 8 teaches **LEND *and* BORROW** on one picture with two blank lines — which line is which? | Split into two slides, or label the lines “(person giving)” / “(person receiving)” |
| S3-5 | 12 | L | The picture (a bus with a “WINDSOR” destination sign) does not clearly read as **COACH** | Use an unambiguous coach picture |
| S3-6 | 14 | M | “04 / 09 Vocabulary Game” — **no game slide** | Add the game (e.g. reuse the BINGO) or delete the section |
| S3-7 | 17–18 | M | The **same email appears on two consecutive slides**; both are titled “Look at the picture and finish the questions” — there is no picture and nothing to “finish” | Delete slide 17; retitle slide 18 “Look at the message and answer the questions” |
| S3-8 | 18 | M | The answer to question 3 (“Choose the correct answer”, A/B/C about Jack and Anna) is **never shown**; only the Q1/Q2 answers are (“It is an email.” / “Jack is sending it to Anna.”) | Add the answer (B: Jack needs Anna to make a decision) and why A/C are wrong |
| S3-9 | 20–21 | M | Both “Let’s check:” slides show tasks + options with **no marked answer** | Mark the correct option (see §A5) |
| S3-10 | 21 | L | Three notices + 3 questions + 9 options are crammed into one image; the text is very small when projected | Split into two slides |
| S3-11 | 24 | L | The Writing Part 6 prompt is fine, but nothing reminds students that **all three** of Tim’s questions must be answered | Add “Answer all three of Tim’s questions.” |
| S3-12 | 26 | M | The structure box teaches “**Let me know asap,**” as the sign-off for a *reply* email, and the greeting “Hi, **Tim!**” | Use one standard frame: “Hi Tim, … Write soon, / Best wishes, (Your name)” |
| S3-13 | 25, 26, 32 | M | Three imported worksheets act as “tips / model answer” (one contains “**White** 25 words or more.”), and the Writing recap (32) has **no tips of its own**, unlike the Reading recap with six | Keep one model answer; add 3–4 Writing recap bullets (“answer all three points”, “25+ words”, “open and close the email”, “check spelling”) |
| S3-14 | 29 | L | The “HOW TO WRITE A STORY (35 WORDS OR MORE)” worksheet is illegible at slide size | Rebuild as native slides or crop to a readable size |
| S3-15 | 27–29 | L | After Part 7 the deck jumps straight from the pictures to a “Your turn” slide that **already shows a model story**; there is no space for students to plan/write and no checklist | Rename it “Model answer”, and add a short “Your turn” slide with the pictures + the 3 planning questions |

### Session 4 — Reading Part 2 + Writing Part 6 & 7 (37 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S4-1 | 4–5 | M | **Two opening sections** back to back — “Warm up” (sl. 4) then “STOP THE BUS!!” (sl. 5). Only one warm-up is needed | Make “STOP THE BUS” section 01 and delete the empty “Warm up” divider |
| S4-2 | 6 | M | “STOP THE BUS” has **no letter and no rules**: “Write down 1 animal starts with letter …..” — which letter? how many points? when does the round end? | Add “Letter: ___ (teacher chooses). 1 point per word; first team to finish says STOP THE BUS.” |
| S4-3 | 10–16 | L | Vocabulary is shown as *picture + word*, i.e. the answer is given away (S3/S7 gap the word instead). “A CHEMIST” should be “A CHEMIST’S” for the shop | Optionally gap the word for one round; relabel slide 14 |
| S4-4 | 19 | L | “Matching the questions **with** the texts” (official wording: *to*) | “Match the questions **to** the texts” |
| S4-5 | 20 | M | “Let’s play kahoot!!!!!!” — the **Kahoot link is not in the deck**, and the next slides (“Read the texts and answer…”) have nothing to do with Kahoot | Put the Kahoot URL in the notes or delete the reference; make the link between the slides explicit |
| S4-6 | 21 | M | “Read the texts and find out phrases that have **same meaning with** the following” — grammar, and two of the three “answers” are **not synonyms**: “rich: → **40,000**” (a number) and “unusual things: → things that other people don’t want” | “Find the word or phrase in the texts that tells you this: *hard to understand / unusual things / rich*”, and give the answer sentence, not just “40,000” |
| S4-7 | 23–24 | L | Slide 23 (“Let’s do Reading Part 2!”) shows the three texts, then slide 24 (“Exam practice”) shows the same texts **plus** the questions — two nearly identical slides | Use 23 for the pre-reading step only; use 24 for the quiz |
| S4-8 | 24 | H | The **answer key is invisible**: every row shows A, B and C under the three names and nothing marks the right one (verified at pixel level). The texts are headed with names, not the exam’s A/B/C labels | Mark the answers (green tick / bold) and/or label the texts A/B/C so the A-B-C grid means something |
| S4-9 | 27–29 | M | The task says “**Write an email**”, slide 29 says “Write **a letter**”; the structure point “**Ask Bobby to wait for you at the airport at what time**” is muddled vs the task’s “suggest a time to meet at the airport” | “Write an email to Bobby…”; “Suggest a time and place to meet at the airport.” |
| S4-10 | 28 | M | “Who is **a** receiver?” + “What tense …? → **Future tense**” (see executive summary #3) | “Who is **the** receiver?”; instead ask “Which tenses will you use? Give one example.” |
| S4-11 | 32–33 | H | “Before writing” gives the whole story away picture by picture, and “**Your turn**” (33) already contains the finished model story — students produce nothing. Also the plan says “**She saw her friends**” while the model says “**her friends saw her** and shouted” | Keep only “Where / What happened / Who”; let students write; then show a “Model answer” slide (fix the who-saw-whom mismatch) |
| S4-12 | 35–36 | L | Reading recap has 3 tips; the Writing recap (36) is only an imported, half-legible “KET WRITING GUIDE” | Add 3–4 native Writing tips (3 content points, 25+ words, linkers, spelling) |
| S4-13 | 37 | L | “SECTION IO / 10” (letter O instead of zero) | “SECTION 10 / 10” |

### Session 5 — Listening Part 2 + Speaking Part 2 (30 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S5-1 | 6 | M | “Take turn to read for **you friend** to fill out the table” — typo + unclear: read aloud while the friend completes the table? | “Take turns to tell your partner about your school day; your partner completes the table. Then swap.” |
| S5-2 | 11 | H | “In KET Listening part 2, you have to fill in the note with **a single words**” — grammar error, and factually it **contradicts the deck’s own key** (60, 40, 7, Saturday). The exam allows a word, a number, a date or a time | “Fill each gap with **one word or a number / a date / a time**. Check your spelling.” |
| S5-3 | 12 | M | “you will be given **10 seconds** to study the note” — verify this against your recording; the standard rubric is “You have X seconds to look at the notes” | Align the number with your audio; keep the useful prediction step |
| S5-4 | 13 | L | “PROVIDING BACKGROUND KNOWLEDGE / LEAVING PRESENT / **Why to give?** When to give? What to give?” — ungrammatical; the slide also mixes the topic with a “Gifts for Friends / A Unique Goodbye” game graphic | “Why give a leaving present? When do we give one? What do we give?” |
| S5-5 | 15 | L | The key presents labels and answers in one mixed block (“month / (8) DVD / (10) pm / 60 / 40 / player / Saturday / 7”), which is hard to read out in class | List answers as “6 … 7 … 8 … 9 … 10 …” under the note items |
| S5-6 | 17 | M | The last answer printed is the name “**Piak**” (“For more info ask: Miss ___”) — looks like a mis-transcription/typo, and names must be spelled correctly in Part 2 | Check it against the recording and correct it (with a capital letter) |
| S5-7 | 20 | M | Under “**Phase 1**” the deck merges both phases (“…After 1-2 minutes, the examiner will ask each candidate two questions…”), and it never says that the exam gives you **five pictures** that you must say you like/don’t like | Split into “Phase 1: five pictures — say what you like/don’t like” and “Phase 2: two questions each”, matching the picture task on slide 23 |
| S5-8 | 21–22 | L | “I don’t know what to say!” + WHAT/WHERE/WHEN/WHY/WHO — the technique is never explained or demonstrated | Add a worked example (question → full answer) |
| S5-9 | 24 | M | “A group of friends have just arrived **at** London … **Suggest them** the best way to travel around.” — grammar, and this is a *giving advice* task (B1 style), not the A2 Key Part 2 likes/dislikes task | “A group of friends have just arrived **in** London. **Give them advice on** the best way to travel around. Say which way you like best and why.” |
| S5-10 | 26 | L | Identical to S8·28 — the same “Two teams / 1 box / 30 seconds” activity appears in two sessions | Differentiate the two speaking-practice sessions |

### Session 6 — Reading Part 3 + Writing Part 6 & 7 (40 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S6-1 | 4–5 | M | Same double opening as S4 (“Warm up” + “STOP THE BUS!!”) | Keep one warm-up section |
| S6-2 | 6 | M | “Write down 1 **animal starts** with letter …” — grammar + the same missing letter and missing rules as S4·6 | “Write down one animal **that** starts with the letter ___”; add the scoring rules |
| S6-3 | 12 | M | The answer analysis is inconsistent and incomplete: “✕ Her parents are heart doctors, not her” / “✕ **C. She used to, not now**” — option C is explained but the **correct answer (B) is never stated**, and the two bullets use different formats | Show “**Answer: B** — she works as a local doctor in the countryside. A ✗ (her parents…), C ✗ (she used to work with children)” |
| S6-4 | 13 | M | A second question (“What happened after Diane was sick?”) ends the slide with “Please do this question!!!” and the answer never appears | Add the answer (A: she became a better doctor — “which made her better at her job”) |
| S6-5 | 15 | L | “Information **center**” (US) — the reading text in the same session says “tourist information **centre**” | “information centre” |
| S6-6 | 18 | L | Vocabulary item “**Fill up**” presented without context; in the text it means “filled up an area (= used all the space)” | “fill (something) up = use all the space in it” + the example from the text |
| S6-7 | 23–24 | M | The same long reading text (“A Very Unusual Museum”) is shown on **two slides** (slide 24 only adds the question numbers 14-18) — a redundant slide that costs class time | Merge: one slide for the text, one slide for questions 14-18 + the answer key |
| S6-8 | 25, 27 | M | The two question sets (Hair Museum Q14-18, Goby Q14-18) have **no answer key**; both use the numbers 14-18, which may confuse students (“we already did these”) | Add answer keys (ideally in the notes) and label the sets “Practice 1 / Practice 2” |
| S6-9 | 30 | M | **“WITH EDITABLE STROKE”** is visible in the imported prompt image (editing artefact) | Re-export the image from source |
| S6-10 | 31 | M | “Who is **a** receiver?” + “What tense …? → Future tense” (Jane is the receiver; the reply mixes present and future) | “Who is **the** receiver?”; “What do you have to include? (3 things)” |
| S6-11 | 32 | H | “Write a letter **to** using the structure below” — **the recipient (Jane) is missing**; also “letter” vs email; “Recommend a good restaurant and **their** famous dishes” | “Write an email **to Jane** using the structure below”; “their” → “its” |
| S6-12 | 33 | M | The Writing Part 7 task slide sits right after the Part 6 structure, with no transition or answer plan for the new pictures until slide 35 | Add a “Look at the pictures. What can you see?” step |
| S6-13 | 35 | L | “the boy went sailing in a small boat **in** a sunny day” | “**on** a sunny day” |
| S6-14 | 36, 39 | M | The model story is duplicated (36 “Your turn” and 39 “Recap – Writing”), and it ends with **no full stop** (“…he swam to the shore”). The last sentence is also an invention (a life jacket) that may not be in picture 3 | Label 36 “Model answer”, add the full stop, and check that the life jacket does not contradict the pictures |
| S6-15 | 39 | L | The recap adds a third-party “How to Write a Story” worksheet whose **own model story is different** (“Sunny morning. a boy went sailing…”) | Keep one model story in the recap |
| S6-16 | 40 | L | Footer is “KET · SESSION 6” here but “KET · TEST n” in S2/S5/S8 | Unify |

### Session 7 — Reading Part 4 + Writing Part 6 & 7 (37 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S7-1 | 2 | H | The objective says “Practice finding details in KET – **Reading Part 3**” — copy-pasted from Session 6; this session teaches **Part 4** | “…in KET – Reading **Part 4** (choosing the right word for each gap)” |
| S7-2 | 7–10, 22 | H | **Part number chaos**: section header “Reading **Part 4**” (7), but “Reading **part 3**” (8), “Reading Part **3**: How to do” (9), “Reading Part **3**: Practice” (10, 22); other slides correctly say Part 4 (21, 23, 24, 35) | Make all of them **Part 4** |
| S7-3 | 4 | M | Section “01 / 09 Warm up” — no warm-up activity | Add one or delete the section |
| S7-4 | 19 | M | “05 / 09 VOCABULARY GAME” — no game slide | Add the game or delete the section |
| S7-5 | 10 | M | Cezanne gap-fill: the notes explain only why the **wrong** options are wrong (“borrowed: take something then return later”, “earned: often go with money”, “appointment: same meaning with meeting”, “company: don’t choose because…”) — the **correct answers (owned / career) are never stated**, and the reason for “company” is muddled | State the answers and the reason they fit: “**1 B owned** (his father had/ran the bank), **2 A career** (a job/profession)”; keep the distractor notes as a second bullet |
| S7-6 | 20–21 | M | Two slides show the **same text** (“A simple life”), one to “Guess the content”, one to “choose the correct answers” — redundant; question 22’s slide adds the options but **no answers** | Merge into two slides: text (with the gaps) → questions + answer key |
| S7-7 | 22 | H | The questions slide is labelled “Reading **Part 3**: Practice” while every neighbouring slide says Part 4 | Fix the label |
| S7-8 | 23–24 | M | The same exam-practice task (“Football players”, Q19-24) appears on **two identical slides**, with **no answer key** | Keep one slide and add the key (19 C earn / 20 B hours / 21 C worse / 22 B young / 23 A enough / 24 B stop — please verify against your key) |
| S7-9 | 27–29 | M | Task says “**Write an email** to James”, slide 29 says “Write **a letter to** using the structure below” — **recipient missing** | “Write an email **to James** using the structure below” |
| S7-10 | 28 | L | “Who is **a** receiver?”; the second question here (“How many questions do you have to answer? → 3”) is the good version — use it everywhere | “Who is **the** receiver?”; keep the “3 things” question |
| S7-11 | 32–33 | M | “Picture 2: Tom and Mary went to a mobile phone shop.” and the model’s last sentence “Later, they could talk over the phone **without seeing each other**” — unnatural/confusing; the model also adds “their trip to Japan”, which is not in the plan or the pictures | Simplify: “She bought a phone and later called Tom.” Remove details that aren’t in the pictures |
| S7-12 | 36 | L | The recap worksheet contains a **different model story** (Mai / Jack) than the Tom-&-Mary story just written | Keep one model story |
| S7-13 | 35 | — | ✔ Reading recap correctly says “Reading Part 4 – Exam Tips” | Keep as the reference wording for the whole deck |

### Session 8 — Listening Part 3 + Speaking Part 2 (32 slides)

| # | Slide | Sev | Issue | Suggested fix |
|---|---|---|---|---|
| S8-1 | 5 | M | “**ODD ONE OUT**: Choose a sentence that has the same meaning **with** the original one” — the activity name (“odd one out”) does not match the task (choosing the synonymous sentence), grammar needs “as”, and the slide gives **only one example and no items to choose from** | Retitle “Same meaning”; add 4–6 items with the pairs |
| S8-2 | 12 | L | Tip 3 “Important nouns, verbs, adjectives in each answer (**Tim didn’t receive a message from Tom.**)” — the example has no “nouns/verbs/adjectives” focus and refers to **Tom**, who does not exist in the script (characters: Tim, Sarah, Jack); slide 12’s example should also match the option wording used on slides 13-15 (“receive a text message”) | Replace with an example from the actual script, e.g. “**Sarah didn’t receive a text message.** — underline *receive a text message*” |
| S8-3 | 16 | H | The script students must read and underline contains real English/logic errors: (a) “Sarah: Oh! So **who will we do**, Tim?” → “**what will we do**”; (b) “it’s not as large as **ours**” → if she means Tim’s house, “**yours**” (and it is unclear whose house “ours” is); (c) the ending contradicts the plan: the party is at **Jack’s house**, but “Tim: Actually, **I thought I would have the party in the garden**” — whose garden? and “I thought I would” is unnatural → “**Actually, I’m going to have the party in the garden at Jack’s**” or “**…in my garden**” (with the parents issue resolved) | Fix all three before using the slide, since students are asked to analyse this exact text |
| S8-4 | 13–15 | M | The same question set is repeated on **three slides** (13/14/15), and the numbering column is confusing (a stray column of 13/14/15 beside questions 11-15); on question 13 the printed stem appears to be only “**Sarah**” | Keep one question slide; check that each stem is a complete question; make sure numbers 11-15 appear once |
| S8-5 | 16, 18 | M | Neither the practice script nor the exam practice (“Alan and his mum”) has an **answer key** on the slide | Add keys: practice 11 B, 12 B, 13 A, 14 B, 15 A *(verify)* and the Alan set |
| S8-6 | 18 | L | Option grammar/register: “his teacher speaks too **slow**” → “too **slowly**” | Fix; also show which answer is correct |
| S8-7 | 21 | M | Same “Phase 1” wording problem as S5·20 (both phases merged; the five-picture task not mentioned) | Align with S5’s corrected slide |
| S8-8 | 24–26 | M | Three consecutive “Speaking Part 2” slides repeat the instructions, and slide 26 (“Choose one of the following outdoor activities and **record your talk**”) is a homework-style task with no rubric (how long? where to submit?) | Merge 24–25; make 26 explicit: “Choose one activity, record a 1-minute answer, submit before the next lesson (checklist: 3 reasons, 1 example, linking words)" |
| S8-9 | 22–23 | L | “I don’t know what to say!” + WHAT/WHERE/WHEN/WHY/WHO — same as S5·21–22, technique not explained | Add a worked example |
| S8-10 | 28 | L | Identical to S5·26 | Differentiate |
| S8-11 | 30 | L | “Be careful **for** distractors (but, oh wait!,…)” — grammar; and again the signal words are labelled “distractors” | “Be careful **of/with** distractors. Listen for *but, oh wait* — they usually introduce the right answer.” |
| S8-12 | 31 | L | Speaking recap is identical to S5·29 | Differentiate or drop one |

---

## C. Exam-format fact-check (what is correct, and what to correct)

| Claim in the decks | Where | Verdict |
|---|---|---|
| Reading & Writing: 7 parts, 32 questions, 60 minutes, 60 marks (Part 1: 6 × 3-option MCQ; Part 2: 7 × matching; Part 3: 5 × 3-option MCQ; Part 4: 6-gap vocabulary cloze with A/B/C; Part 5: 6-gap open cloze, one word; Part 6: email 25+ words; Part 7: story 35+ words from 3 pictures) | S3·6, S4·8, S6·8, S7·6 | ✔ **Accurate** (matches the Cambridge 2020 A2 Key format and task descriptions) |
| Listening: 5 parts × 5 questions = 25 marks, approx. 30 min incl. 6 min to transfer answers; Part 1 3-option MCQ with visuals; Part 2 note gap-fill; Part 3 dialogue + 3-option MCQ; Part 4 main idea/gist, 3-option MCQ; Part 5 matching | S2·8, S5·8, S8·7 | ✔ **Accurate** |
| Speaking: 8–10 min; Part 1 = examiner asks each candidate questions (3–4 min); Part 2 = discussion task with visual stimulus, 5–6 min, examiner–candidate + candidate–candidate; 25 marks | S2·9, S5·9, S8·8 | ✔ **Accurate** |
| “In Listening Part 2 you have to fill in the note with **a single word**” | S5·11 | ✘ **Wrong / too restrictive.** The paper allows a word **or a number / date / time**, and the deck’s own key contains 60, 40, 7 and Saturday. Also ungrammatical (“a single words”) |
| “You will be given **10 seconds** to study the note” | S5·12 | ⚠ **Verify** against your recording; use the paper’s wording (“You have … seconds to look at the notes”) |
| Speaking Part 1: Phase 1 general questions, Phase 2 two topical questions + one longer-answer question | S2·18 | ✔ **Accurate in substance** (fix the grammar/punctuation — see S2-5) |
| Speaking Part 2 “Phase 1” = talk about pictures, then two questions each | S5·20, S8·21 | ⚠ **Merges the two phases** and omits that you get **five pictures** and must say what you like/don’t like about them (the practice slides S5·23 / S8·24–25 do show the five-picture task) |
| Reading Part 3 = longer text, 5 × 3-option questions in text order | S6·10 | ✔ **Accurate** |
| Reading Part 4 = text with 6 gaps, 3 options, tests vocabulary | S7·8 | ✔ **Accurate as a task description**, but four slides label this part “Part 3” (see S7-2/S7-7) |
| “This part will test your vocabulary rather than grammar” (S7·8) | S7·8 | ⚠ **Simplification.** Part 4 tests vocabulary *in context*, but collocation/grammar knowledge is needed too → “Choose the word that fits the **meaning and the grammar** of the sentence.” |

---

## D. Coverage map and curriculum recommendations

| KET component | Covered in | Gap |
|---|---|---|
| Reading Part 1 | S3 | |
| Reading Part 2 | S4 | |
| Reading Part 3 | S6 | |
| Reading Part 4 | S7 | |
| **Reading Part 5** (open cloze, one word) | — | **Never taught** in Sessions 2–8 |
| Writing Part 6 (email) + Part 7 (story) | S3, S4, S6, S7 | Taught **four** times with the same slide sequence |
| Listening Part 1 | S2 | |
| Listening Part 2 | S5 | |
| Listening Part 3 | S8 | |
| **Listening Part 4** (main idea/gist, 5 monologues/dialogues) | — | **Never taught** |
| **Listening Part 5** (matching) | — | **Never taught** |
| Speaking Part 1 | S2 | |
| Speaking Part 2 | S5 **and** S8 (near-duplicate) | |

**Recommendations**

1. **Re-allocate two sessions** to Reading Part 5 and Listening Parts 4–5 — they are worth 16 of the exam’s marks (Reading Part 5 = 6; Listening Parts 4+5 = 10) and currently no session touches them.
2. **Split the writing work**: keep Part 6 (email) in one session and Part 7 (story) in another; use the freed slots in S6/S7 for Reading Part 5 or for **timed writing with feedback** instead of re-teaching the same four-slide template.
3. **Differentiate the two Speaking Part 2 sessions** (e.g. S5 = likes/dislikes language + the five-picture task; S8 = asking your partner questions, giving reasons, the two-question phase) so that S5 and S8 do not repeat the same slides.
4. **Stop re-showing the exam-format tables in every session** (they currently appear 4× and 3× verbatim). Show the full format once and use a one-line recap afterwards — it frees ~5 minutes per lesson, which is exactly what the missing answer-key checking needs.
5. **Fill or delete the “Vocabulary Game” and “Warm up” sections** — a divider that promises an activity and delivers none is the most confusing thing for a substitute teacher.

---

## E. Things to verify against your own audio / key (I could not verify them)

1. **S5·17** — the name printed as the answer (“Piak”); check the spelling in the recording (Part 2 penalises spelling).
2. **S5·12** — the “10 seconds” figure, against the recording you use.
3. **S2·6** — the story-telling word sets: confirm each set really contains exactly three words (the table is an image).
4. **S8·13–15** — the exact stems of questions 11–15 (the printed stem of question 13 looks like just “Sarah”), and the answer key.
5. **S4·24, S3·20–21, S6·25·27, S7·22·24, S8·18** — the official answer keys; if your edition has a key sheet, add it to the speaker notes.
6. **S6·30** — the Jane email prompt: check whether the “From / To / Subject” headers are missing (only “Jane / This Sunday” is visible).

---

## F. Quick-win checklist for whoever edits the decks

- [ ] Find & replace: “Write a letter” → “Write an email”; “a receiver” → “the receiver”; “meaning with” → “meaning as”; “Movie theater” → “cinema”; “favorite” → “favourite”; “information center” → “information centre”; “SECTION IO” → “SECTION 10”; “Take turn” → “Take turns”.
- [ ] S6·32 / S7·29 — insert the missing recipient name (Jane / James).
- [ ] S5·11 — rewrite the Listening Part 2 instruction (word / number / date / time).
- [ ] S7 — change every “Reading Part 3” to “Reading Part 4” (objective on slide 2 + slides 8, 9, 10, 22).
- [ ] S8·16 — correct the three script errors; then check the practice key (11–15).
- [ ] Add answer keys (on-slide or in the notes) to: S3·18·20·21, S4·21·24, S5·15·17, S6·12·13·25·27·36, S7·10·22·24·33, S8·16·18.
- [ ] Rename every “Your turn” slide that actually contains the model answer → “Model answer”.
- [ ] Delete / merge the duplicated slides listed in §A6 and the per-session tables.
- [ ] Put the audio file names on the listening slides and the Kahoot URL (S4·20) in the notes.
- [ ] Re-export the imported worksheet images without artefacts (S6·30 “WITH EDITABLE STROKE”, S3·26 “White 25 words or more.”).
- [ ] Add “01 / 02” to the class-rules slide in S3.
- [ ] Delete the empty “Warm up” / “Vocabulary Game” dividers (or add the activity).

---

*Prepared from the seven `.pptx` files in this folder. Every slide reference is “S<session>·<slide number>”, and the quoted text is what the slide actually shows (text layer or in-image text). Items marked ⚠ are the ones I recommend you confirm against your own audio/answer sheets before changing them.*

- **No audio is embedded or linked in any deck** (0 audio/video relationships, 0 external links), although several slides say “Listen and …” or show a speaker icon (e.g. S2·15, S5·15, S5·17, S8·15, S8·18). The trainer must always bring the recordings separately.
- **No slide is marked hidden**; every section divider is visible in the show.
- Where answer keys do exist they sit **on the same slide** as the task; several “Let’s check:” slides have **no marking at all** (see §A5).
