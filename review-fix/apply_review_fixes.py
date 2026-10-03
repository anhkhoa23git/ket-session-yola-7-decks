#!/usr/bin/env python3
"""
Apply content-review fixes to 7 SESSION PPTX decks (Yola template).
Two kinds of fixes:
  A) Run-level text replacements (exact or substring)
  B) New elements (answer keys, missing rule numbers) added by cloning/creating shapes
Every fix is logged APPLIED / NOT FOUND for the report.
"""
import copy
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BASE = "/home/z/my-project/download/"
NAVY = RGBColor(0x0F, 0x2B, 0x4C)
GREEN = RGBColor(0x1F, 0xA3, 0x4B)

# ---------- A) TEXT REPLACEMENTS ----------
# (deck, slide_index 1-based or None=all slides, old, new, exact_run, label)
R = []
def fx(deck, slide, old, new, label, exact=False):
    R.append((deck, slide, old, new, exact, label))

# --- Session 2 ---
fx(2, None, "KET · TEST 2", "KET · SESSION 2", "S2 footer unify (24x)")
fx(2, 5, "You take turn to tell a story (longer than 30 seconds) that has 3 words provided",
      "You take turns to tell a story (longer than 30 seconds) using the three words given", "S2-1a take turns + 3 words")
fx(2, 5, "Each set of words must be used once only",
      "Each set of three words can be used only once", "S2-1b once only")
fx(2, 13, "Movie theater", "Cinema", "S2-4 US->UK cinema")
fx(2, 18, "For example: “where are you from?” Or “how old are you?”",
        "For example: “Where are you from?” / “How old are you?”", "S2-5a example punctuation")
fx(2, 18, "2 topical questions and 1 question that requires candidates to give long answer.",
        "2 topical questions and 1 question that requires a longer answer.", "S2-5b long answer")
fx(2, 18, "Tell me about your bedroom", "Tell me about your bedroom.", "S2-5c full stop")
fx(2, 19, "Take turn to ask and answer questions", "Take turns to ask and answer questions", "S2-6a take turns")
fx(2, 20, "Please choose one of the following topics and write a short paragraph (40 words max)",
        "Please choose one of the following topics. Make short notes (max 40 words), then answer aloud", "S2-7 speaking not writing")
fx(2, 20, "Tell me about your favorite room", "Tell me about your favourite room", "S2-8 favourite")
fx(2, 22, "Watch out for distractors (wait, but,...)",
        "Listen for words like “but, actually, oh wait” — they often introduce the answer", "S2-9 distractors")

# --- Session 3 ---
fx(3, 18, "Look at the picture and finish the questions",
        "Look at the message and answer the questions", "S3-7 retitle")
fx(3, 26, "Hi, Tim!", "Hi Tim,", "S3-12a greeting")
fx(3, 26, "Let me know asap,", "Write soon,", "S3-12b sign-off")
fx(3, 26, "Your turn", "Model answer", "S3-15 rename (has MODEL ANSWER chip)", exact=True)

# --- Session 4 ---
fx(4, 6, "1 animal starts with letter", "one animal that starts with the letter", "S4-2a grammar")
fx(4, 6, "1 city starts with letter", "one city that starts with the letter", "S4-2b grammar")
fx(4, 6, "1 food starts with letter", "one food that starts with the letter", "S4-2c grammar")
fx(4, 19, "Matching the questions with the texts", "Match the questions to the texts", "S4-4 match to")
fx(4, 21, "Read the texts and find out phrases that have same meaning with the following:",
        "Read the texts and find the words or phrases that have the same meaning as the following:", "S4-6 meaning as")
fx(4, 28, "Who is a receiver?", "Who is the receiver?", "S4-10a the receiver")
fx(4, 28, "What tense do you use to write this email?", "How many things must you include in this email?", "S4-10b tense question")
fx(4, 28, "Future tense", "3 things", "S4-10c answer", exact=True)
fx(4, 29, "Write a letter to Bobby using the structure below:",
        "Write an email to Bobby using the structure below:", "S4-9a email")
fx(4, 29, "Why you write this letter", "Why you write this email", "S4-9b email")
fx(4, 29, "Ask Bobby to wait for you at the airport at what time",
        "Suggest a time and place to meet Bobby at the airport", "S4-9c muddled point")
fx(4, 32, "Picture 2: She saw her friends", "Picture 2: Her friends saw her and shouted", "S4-11 plan/model mismatch")
fx(4, 33, "Your turn", "Model answer", "S4-11 rename", exact=True)

# --- Session 5 ---
fx(5, None, "KET · TEST 3", "KET · SESSION 5", "S5 footer unify (10x)")
fx(5, 6, "Take turn to read for you friend to fill out the table",
        "Take turns to tell your friend about your school day so they can fill out the table", "S5-1 typo+clarity")
fx(5, 11, "you have to fill in the note with a single words",
        "you have to fill each gap with one word, a number, a date or a time", "S5-2 single words")
fx(5, 13, "Why to give?", "Why do we give it?", "S5-4a")
fx(5, 13, "When to give?", "When do we give it?", "S5-4b")
fx(5, 13, "What to give?", "What do we give?", "S5-4c")
fx(5, 20, "You and the other candidate will talk about some pictures based on the questions from the examiner.",
        "You and the other candidate will talk about five pictures. Say what you like or don’t like about them.", "S5-7a five pictures")
fx(5, 20, "After 1-2 minutes, the examiner will ask each candidate two questions related to the topic",
        "After 1-2 minutes (Phase 2), the examiner will ask each candidate two questions related to the topic", "S5-7b phase 2")
fx(5, 24, "arrived at London", "arrived in London", "S5-9a in London")
fx(5, 24, "Suggest them the best way to travel around.", "Give them advice on the best way to travel around.", "S5-9b advice")

# --- Session 6 ---
fx(6, 6, "1 animal starts with letter", "one animal that starts with the letter", "S6-2a grammar")
fx(6, 6, "1 city starts with letter", "one city that starts with the letter", "S6-2b grammar")
fx(6, 6, "1 food starts with letter", "one food that starts with the letter", "S6-2c grammar")
fx(6, 12, "Why are option A and C incorrect?", "Answer: B — a village doctor now. A and C are incorrect:", "S6-3 state answer")
fx(6, 12, "✕ Her parents are heart doctors, not her", "✕ A — her parents are heart doctors, not her", "S6-3 bullet A")
fx(6, 12, "✕ C. She used to, not now", "✕ C — she used to work with children, not now", "S6-3 bullet C")
fx(6, 13, "Please do this question!!!", "Answer: A — she became a better doctor", "S6-4 reveal answer")
fx(6, 15, "Information center", "Information centre", "S6-5 UK spelling")
fx(6, 31, "Who is a receiver?", "Who is the receiver?", "S6-10a")
fx(6, 31, "What tense do you use to write this email?", "How many things must you include in this email?", "S6-10b")
fx(6, 31, "Future tense", "3 things", "S6-10c", exact=True)
fx(6, 32, "Write a letter to using the structure below:", "Write an email to Jane using the structure below:", "S6-11a missing Jane")
fx(6, 32, "Recommend a good restaurant and their famous dishes", "Recommend a good restaurant and its famous dishes", "S6-11b its")
fx(6, 35, "in a small boat in a sunny day", "in a small boat on a sunny day", "S6-13 on a sunny day")
fx(6, 36, "Your turn", "Model answer", "S6-14 rename", exact=True)
fx(6, 36, "so he swam to the shore", "so he swam to the shore.", "S6-14 full stop")
fx(6, 39, "so he swam to the shore", "so he swam to the shore.", "S6-14 full stop (recap)")

# --- Session 7 ---
fx(7, 2, "Practice finding details in KET - Reading Part 3",
      "Practice choosing the right word for each gap in KET - Reading Part 4", "S7-1 objective Part 4")
fx(7, 8, "Reading part 3", "Reading part 4", "S7-2 title", exact=True)
fx(7, 8, "This part will test your vocabulary rather than grammar",
      "Choose the word that fits the meaning and the grammar of the sentence", "S7-2 vocabulary claim")
fx(7, 9, "Reading Part 3: How to do", "Reading Part 4: How to do", "S7-2 label")
fx(7, 10, "Reading Part 3: Practice", "Reading Part 4: Practice", "S7-7 label")
fx(7, 10, "appointment: same meaning with meeting", "appointment: similar in meaning to meeting", "S7-5 meaning")
fx(7, 22, "Reading Part 3: Practice", "Reading Part 4: Practice", "S7-7 label")
fx(7, 28, "Who is a receiver?", "Who is the receiver?", "S7-10a")
fx(7, 29, "Write a letter to using the structure below:", "Write an email to James using the structure below:", "S7-9 missing James")
fx(7, 33, "Your turn", "Model answer", "S7-11 rename", exact=True)
fx(7, 33, "One day, Tom and Mary met at the park and talked about their trip to Japan. Then, they went to a mobile phone shop to buy one for Mary. Later, they could talk over the phone without seeing each other.",
        "One day, Tom and Mary met at the park. Then, they went to a mobile phone shop to buy a phone for Mary. Later, she called Tom and they talked happily.", "S7-11 simplify story")

# --- Session 8 ---
fx(8, None, "KET · TEST 4", "KET · SESSION 8", "S8 footer unify (10x)")
fx(8, 5, "ODD ONE OUT:", "SAME MEANING:", "S8-1a retitle")
fx(8, 5, "has the same meaning with the original one", "has the same meaning as the original one", "S8-1b meaning as")
fx(8, 12, "Important nouns, verbs, adjectives in each answer (Tim didn’t receive a message from Tom.)",
        "Important nouns, verbs, adjectives in each answer (Sarah didn’t receive a text message.)", "S8-2 example fix")
fx(8, 16, "Sarah: Oh! So who will we do, Tim?", "Sarah: Oh! So what will we do, Tim?", "S8-3a script")
fx(8, 16, "it’s not as large as ours.", "it’s not as large as yours.", "S8-3b script")
fx(8, 16, "Actually, I thought I would have the party in the garden.",
        "Actually, I’m going to have the party in Jack’s garden.", "S8-3c script")
fx(8, 21, "You and the other candidate will talk about some pictures based on the questions from the examiner.",
        "You and the other candidate will talk about five pictures. Say what you like or don’t like about them.", "S8-7a five pictures")
fx(8, 21, "After 1-2 minutes, the examiner will ask each candidate two questions related to the topic",
        "After 1-2 minutes (Phase 2), the examiner will ask each candidate two questions related to the topic", "S8-7b phase 2")
fx(8, 26, "Choose one of the following outdoor activities and record your talk",
        "Choose one of the following outdoor activities and record a 1-minute talk: 3 reasons + 1 example + linking words", "S8-8 rubric")
fx(8, 30, "Be careful for distractors (but, oh wait!,...)",
        "Be careful of distractors — listen for “but, oh wait!”", "S8-11a")
fx(8, 30, "Listen to understand not to choose the answer", "Listen to understand, not just to choose the answer", "S8-11b")


def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def apply_text_fixes():
    log = []
    decks = sorted({r[0] for r in R})
    for deck in decks:
        fname = f"SESSION-{deck}-A2-KET-TRAINER_Yola.pptx"
        prs = Presentation(BASE + fname)
        fixes = [r for r in R if r[0] == deck]
        applied = {id(f): 0 for f in fixes}
        for si, slide in enumerate(prs.slides, 1):
            for sh in iter_shapes(slide.shapes):
                frames = []
                if sh.has_text_frame:
                    frames.append(sh.text_frame)
                if getattr(sh, "has_table", False) and sh.has_table:
                    for row in sh.table.rows:
                        for cell in row.cells:
                            frames.append(cell.text_frame)
                for tf in frames:
                    for para in tf.paragraphs:
                        for run in para.runs:
                            for f in fixes:
                                _, sl, old, new, exact, label = f
                                if sl is not None and sl != si:
                                    continue
                                if exact:
                                    if run.text == old:
                                        run.text = new
                                        applied[id(f)] += 1
                                elif old in run.text:
                                    run.text = run.text.replace(old, new)
                                    applied[id(f)] += 1
        prs.save(BASE + fname)
        for f in fixes:
            _, sl, old, new, exact, label = f
            cnt = applied[id(f)]
            loc = f"S{deck}·{sl}" if sl else f"S{deck} all"
            log.append((loc, label, cnt))
    return log


if __name__ == "__main__":
    log = apply_text_fixes()
    missing = [l for l in log if l[2] == 0]
    for loc, label, cnt in log:
        status = "OK " if cnt else "MISS"
        print(f"[{status}] {loc:<10} {label} (x{cnt})")
    print(f"\nTotal: {len(log)} fixes, {len(missing)} not found")
    json.dump(log, open("/home/z/my-project/work_ss/fix_log.json", "w"), ensure_ascii=False, indent=1)
