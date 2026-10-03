#!/usr/bin/env python3
"""Verify all review fixes landed + PPTX integrity (slides, fade, fonts)."""
import json, re, zipfile
from pptx import Presentation

BASE = "/home/z/my-project/download/"
N = {2: 24, 3: 33, 4: 37, 5: 30, 6: 40, 7: 37, 8: 32}

# (deck, slide, must_be_ABSENT, must_be_PRESENT)
CHECKS = [
    (2, 5, "You take turn to tell", "using the three words given"),
    (2, 5, "must be used once only", "can be used only once"),
    (2, 13, "Movie theater", "Cinema"),
    (2, 18, "Or “how old", "“Where are you from?” / “How old are you?”"),
    (2, 18, "candidates to give long answer", "requires a longer answer."),
    (2, 19, "Take turn to ask", "Take turns to ask"),
    (2, 20, "write a short paragraph (40 words max)", "Make short notes (max 40 words), then answer aloud"),
    (2, 20, "favorite room", "favourite room"),
    (2, 22, "Watch out for distractors", "they often introduce the answer"),
    (2, 1, "KET · TEST 2", "KET · SESSION 2"),
    (3, 3, None, "01"),
    (3, 18, "finish the questions", "answer the questions"),
    (3, 18, None, "Jack needs Anna to make a decision"),
    (3, 26, "Hi, Tim!", "Hi Tim,"),
    (3, 26, "Let me know asap,", "Write soon,"),
    (3, 26, "Your turn", "Model answer"),
    (4, 6, "1 animal starts with letter", "one animal that starts with the letter"),
    (4, 19, "Matching the questions with", "Match the questions to the texts"),
    (4, 21, "same meaning with the following", "same meaning as the following"),
    (4, 28, "Who is a receiver?", "Who is the receiver?"),
    (4, 28, "Future tense", "3 things"),
    (4, 29, "Write a letter to Bobby", "Write an email to Bobby"),
    (4, 29, "Ask Bobby to wait for you at the airport at what time", "Suggest a time and place to meet Bobby at the airport"),
    (4, 32, "Picture 2: She saw her friends", "Her friends saw her and shouted"),
    (4, 33, None, "Model answer"),
    (5, 6, "read for you friend", "Tell your friend about your school day"),
    (5, 11, "a single words", "one word, a number, a date or a time"),
    (5, 13, "Why to give?", "Why do we give it?"),
    (5, 20, "talk about some pictures", "talk about five pictures"),
    (5, 24, "arrived at London", "arrived in London"),
    (5, 24, "Suggest them the best way", "Give them advice on the best way"),
    (6, 6, "1 animal starts with letter", "one animal that starts with the letter"),
    (6, 12, "Why are option A and C incorrect?", "Answer: B — a village doctor now"),
    (6, 12, "C. She used to, not now", "she used to work with children, not now"),
    (6, 13, "Please do this question!!!", "Answer: A — she became a better doctor"),
    (6, 15, "Information center", "Information centre"),
    (6, 31, "Who is a receiver?", "Who is the receiver?"),
    (6, 32, "Write a letter to using", "Write an email to Jane using"),
    (6, 32, "their famous dishes", "its famous dishes"),
    (6, 35, "in a small boat in a sunny day", "in a small boat on a sunny day"),
    (6, 36, None, "Model answer"),
    (6, 36, "swam to the shore |", "swam to the shore."),
    (7, 2, "Reading Part 3", "Reading Part 4"),
    (7, 8, "Reading part 3", "Reading part 4"),
    (7, 9, "Part 3: How to do", "Part 4: How to do"),
    (7, 10, "same meaning with meeting", "similar in meaning to meeting"),
    (7, 22, "Part 3: Practice", "Part 4: Practice"),
    (7, 23, None, "ANSWERS"),
    (7, 24, None, "ANSWERS"),
    (7, 29, "Write a letter to using", "Write an email to James using"),
    (7, 33, "trip to Japan", "buy a phone for Mary"),
    (8, 5, "ODD ONE OUT", "SAME MEANING:"),
    (8, 5, "meaning with the original", "meaning as the original"),
    (8, 12, "Tim didn’t receive a message from Tom", "Sarah didn’t receive a text message"),
    (8, 16, "who will we do", "what will we do"),
    (8, 16, "as large as ours", "as large as yours"),
    (8, 16, "I thought I would have the party in the garden", "going to have the party in Jack’s garden"),
    (8, 16, None, "11 · B"),
    (8, 21, "talk about some pictures", "talk about five pictures"),
    (8, 26, "record your talk", "record a 1-minute talk"),
    (8, 30, "Be careful for distractors", "Be careful of distractors"),
    (8, 30, "Listen to understand not to choose", "not just to choose the answer"),
    (8, 1, "KET · TEST 4", "KET · SESSION 8"),
]

def slide_text(prs, i):
    out = []
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                walk(sh.shapes)
            elif sh.has_text_frame:
                out.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        out.append(c.text_frame.text)
    walk(prs.slides[i - 1].shapes)
    return "\n".join(out)

fails = 0
prs_cache = {}
for deck, sl, absent, present in CHECKS:
    if deck not in prs_cache:
        prs_cache[deck] = Presentation(BASE + f"SESSION-{deck}-A2-KET-TRAINER_Yola.pptx")
    txt = slide_text(prs_cache[deck], sl)
    ok = True
    if absent and absent in txt:
        ok = False
        print(f"[FAIL] S{deck}·{sl}: OLD STILL PRESENT: {absent!r}")
    if present and present not in txt:
        ok = False
        print(f"[FAIL] S{deck}·{sl}: NEW MISSING: {present!r}")
    if ok:
        print(f"[OK]   S{deck}·{sl}")

# integrity: slide counts, fade, fonts
print("\n--- integrity ---")
for d, want in N.items():
    fn = BASE + f"SESSION-{d}-A2-KET-TRAINER_Yola.pptx"
    z = zipfile.ZipFile(fn)
    slides = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    fades = sum(1 for s in slides if "<p:fade/>" in z.read(s).decode("utf8", "ignore"))
    fonts = set()
    for s in slides:
        fonts.update(re.findall(r'typeface="([^"]+)"', z.read(s).decode("utf8", "ignore")))
    bad = fonts - {"Baloo 2", "Nunito", "Segoe UI Symbol", "+mj-lt", "+mn-lt"}
    st = "OK " if (len(slides) == want and fades == want and not bad) else "FAIL"
    if st == "FAIL":
        fails += 1
    print(f"[{st}] SESSION-{d}: slides={len(slides)}/{want} fade={fades}/{want} fonts={sorted(fonts)}")
print("ALL CHECKS PASSED" if fails == 0 else f"{fails} INTEGRITY FAILURES")
