# Analyzing an exported chat

Three steps to turn a multi-year conversation into something you can act on,
without reading it. Each step below says exactly what to attach, what to copy
and paste, and what to do with the result.

Works with Claude, ChatGPT, or anything else that takes file uploads.

## What you're working with

After an export you have:

| File | Used in |
|---|---|
| `chunks/part-NN_*.md` | Step 1 — one run each |
| `images/index.md` and `images/*.png` | Step 2 — optional |
| `records.md` | you create this; it's the input to Step 3 |
| `transcript.html` | not used here — that's the one for reading yourself |

You'll create one new file, `records.md`, by pasting each step's output into it.
That file is the real product. Keep it: you can re-run Step 3 against it later
with a different question and never touch the transcript again.

---

# Step 1 — Extract records from each chunk

**Do this once per chunk file.** Three chunks means three runs. Use the same
conversation for all of them.

**Attach:** one chunk file, e.g. `chunks/part-01_2019-03-14_to_2021-08-02.md`

**Paste this, replacing `{N}` and `{TOTAL}`:**

```
You are extracting a structured record from one segment of a multi-year
Microsoft Teams conversation. This is segment {N} of {TOTAL}. Work only from the
attached segment. Do not reference or assume anything about other segments.

Do not summarize. Produce records.

A message earns a record only if it does one of these:
- changes state: a decision, commitment, deadline, approval, escalation, or a
  reversal of an earlier one
- carries knowledge that exists nowhere else: how or why something works, a
  workaround, an owner, history
- is a stated assessment of a person, vendor, or situation
- is an unresolved thread: asked and never answered, handed off and never
  confirmed
- is sensitive or risky content
- marks a point where the conversation moved off-channel to a call or meeting

Ignore greetings, scheduling logistics, acknowledgments, reactions, and
pleasantries. They produce no records.

Do not stretch to fill a category. If a category has nothing in this segment,
write NONE. Inventing findings to fill the schema is worse than an empty
section.

Every record must carry a date and a verbatim quote. If you can't quote it,
don't record it. Keep quotes to 25 words.

Use these exact headings, one line per record, in this form:
YYYY-MM-DD | who | what happened | "verbatim quote"

DECISIONS - something was settled. Note if it reverses something earlier.
COMMITMENTS - someone took ownership or named a date.
KNOWLEDGE - how or why something works. The tribal knowledge.
PEOPLE - third parties named, their role, and any assessment made of them.
FRICTION - disagreements, complaints, escalations, frustration.
OPEN LOOPS - asked and never answered, handed off and never confirmed.
SENSITIVE - personnel matters, credentials, confidential material. Describe the
category; do not reproduce secrets.
TERMS - internal acronyms, system names, project codenames, shorthand. Give the
term and what it appears to mean from context. Mark a guess as a guess.
GAPS - where the conversation moved to a call, meeting, or another channel. Note
the date and the topic that was about to be discussed.
IMAGES - for each [IMAGE: ...] marker, what the surrounding text suggests it
contained, and whether it looks load-bearing. List the filenames.
```

**What you do with the output:** create `records.md` and paste the result in,
with a line above it naming the segment:

```
## From part-01_2019-03-14_to_2021-08-02.md
<paste output here>
```

Repeat for every chunk, in date order, appending each to the same file.

---

# Step 2 — Recover the screenshots that matter (optional)

Skip this if the chat has few screenshots or the Step 1 `IMAGES` sections say
none of them are load-bearing.

Step 1 told you which image files look important. Collect those filenames.

**Attach:** `images/index.md` plus only the PNGs Step 1 flagged. Don't upload
all of them.

**Paste this:**

```
The attached PNG files are screenshots pasted into a Microsoft Teams
conversation. index.md gives each screenshot's filename, when it was sent, who
sent it, and the conversation immediately before and after it.

For each attached image, read its contents and produce one entry:

FILENAME | YYYY-MM-DD | sender | what the image actually contains

Then, underneath, transcribe any text in the image that carries meaning -
document contents, message threads, error text, figures. Skip decorative
material and UI chrome.

If an image is illegible or contains nothing of substance, say so in one line
rather than guessing.

Do not interpret or draw conclusions. Report what is in the images.
```

**What you do with the output:** append it to `records.md` under a heading:

```
## Screenshot contents
<paste output here>
```

---

# Step 3 — The assessment

**Do this once**, in a fresh conversation.

**Attach:** `records.md`

**Paste this, replacing `{DATE RANGE}` and `{PEOPLE}`:**

```
The attached file contains structured records extracted from a Microsoft Teams
conversation spanning {DATE RANGE} between {PEOPLE}, in chronological order.
Each record has a date and a verbatim quote. Some entries describe the contents
of screenshots that were pasted into the conversation.

Analyze across the whole timeline. I have not read the underlying transcript and
will act on what you tell me, so every claim must cite the dates of the records
supporting it. Where the records don't support a conclusion, say so rather than
filling the gap.

Produce:

1. ARC. What this relationship or workstream actually was, and how it changed.
Name the inflection points by date.

2. DECISION REGISTER. Every decision and what became of it: held, quietly
reversed, superseded, or never closed. Flag reversals that were never
acknowledged as reversals.

3. RECURRING THEMES. Only things appearing across multiple records over time.
Give the date span and how many times. A one-off complaint is not a theme -
distinguish them explicitly.

4. WORKING PROFILES. For each participant, how they operate as evidenced by the
records: how they escalate, whether they commit in writing or move to a call,
whether they raise problems early or late, how they handle disagreement, what
they reliably follow through on. Cite dates. Do not infer personality types,
clinical traits, or motives - stay with observable behavior.

5. KNOWLEDGE THAT LEAVES WITH THEM. What does this person know that the records
show nobody else does? Rank by how expensive it would be to reconstruct.

6. CONTRADICTIONS. Where the record disagrees with itself across time.

7. OPEN AT THE END. What was still unresolved when the conversation stopped, and
who was holding it.

8. BLIND SPOTS. Where the substance moved off-channel. List the dates and topics
so I know what this transcript cannot tell me.

9. SENSITIVE REGISTER. What is in here that should not be circulated, by
category and date range. Do not reproduce the content.

Finish with the five things I most need to know, in order.
```

---

# Why it's built this way

**Two passes, not one.** Hand a model all the chunks and ask for an assessment
and you get a shallow summary of each plus a generic synthesis. What dies there
is everything only visible across time — a decision made in year one and quietly
reversed in year three appears in no single chunk. Summaries are lossy and don't
merge. Dated records do.

**Records, not summaries.** Step 1 never asks for prose. It asks for typed lines
with dates and quotes, because those merge into a single timeline that Step 3 can
reason over. Hundreds of thousands of tokens of transcript become a few thousand
tokens of evidence.

**The quote requirement is an anti-hallucination device.** A model that must
quote can't invent a decision. It also means you can verify any single claim by
searching `transcript.html` for that quote instead of rereading the year around
it.

**Permission to return NONE matters.** Without it, a model manufactures findings
to fill the schema, and the manufactured ones look exactly like the real ones.

**Small talk needs no instruction.** Nothing in the schema fits "sure", "thanks",
or "got a minute" — so chatter produces no records and disappears on its own.
