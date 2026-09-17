Here is a cleaner and more complete version that captures the **mathematical uniqueness + practical/human usefulness + parent-child structure + keyword priority** rules.

## Minimum Keyword Identification Rules

### 1. Start with the current keyword list

```text
Current Keyword List
        ↓
Build complete identifiable items
(Parent description + Child qualifier)
        ↓
Generate candidate keywords
        ↓
Apply keyword priority hierarchy
        ↓
1-word identification
        ↓
2-word identification
        ↓
3-word identification
        ↓
...
        ↓
Minimum useful unique combination
        ↓
If no combination works
        ↓
NO UNIQUE IDENTIFICATION
```

### 2. Parent–child rule

A child row should **not automatically be treated as an independent item**.

```text
Parent Description
        +
Child Qualifier
        =
Complete Identifiable Item
```

For example:

```text
Earth work in excavation...
        +
Hard rock
        ↓
Earth work in excavation... Hard rock
```

The complete item is what must be compared against other complete items.

This prevents a generic child such as `hard rock`, `all kinds of soil`, or `fly ash` from being incorrectly treated as a standalone item.

---

## 3. Progressive identification

For **each complete item**, test the smallest number of keywords required to distinguish it from every other item.

```text
ROW
 ↓
Try every allowed 1-keyword combination
 ↓
Does exactly one item match?
 ├── YES → DONE
 └── NO
       ↓
Try every allowed 2-keyword combination
       ↓
Does exactly one item match?
 ├── YES → DONE
 └── NO
       ↓
Try every allowed 3-keyword combination
       ↓
Does exactly one item match?
 ├── YES → DONE
 └── NO
       ↓
Continue...
       ↓
Minimum unique combination
```

A combination is **unique** only when:

> **Exactly one complete item contains all the selected keywords.**

---

# 4. Minimum does not mean blindly shortest

The objective is **not simply**:

> “Find the fewest words.”

The objective is:

> **Find the fewest useful discriminating words that uniquely identify the item.**

Therefore, keyword selection has two stages:

```text
MINIMUM NUMBER OF WORDS
              +
KEYWORD QUALITY / PRIORITY
              ↓
BEST IDENTIFIER
```

For example, suppose:

```text
Item A → Fly ash and earth filling with 98% density
```

and `98%` occurs nowhere else.

Mathematically:

```text
98%
```

is enough to identify the item.

But `98%` is an **accidental/numeric identifier**, whereas:

```text
fly + ash + filling
```

is much more meaningful to a human.

Therefore, the system should distinguish between:

```text
STRICT MINIMUM
```

and

```text
PREFERRED / MEANINGFUL MINIMUM
```

---

# 5. Keyword priority hierarchy

When several combinations have the same number of words, prefer the combination with the **better discriminating keywords**.

### Priority 1 — Strong meaningful discriminator

Words that clearly describe the nature of the work or material.

Examples:

```text
chlorpyriphos
timbering
ploughing
masonry
stacking
levelling
apron
masonry
```

### Priority 2 — Meaningful distinguishing qualifier

Words that distinguish the item from its siblings.

Examples:

```text
external
existing
ordinary
hard
prohibited
wood
floors
foundation
```

### Priority 3 — Meaningful combinations

When one word is insufficient:

```text
fly + ash
external + wall
existing + masonry
wood + contact
```

### Priority 4 — Technical/context words

Useful when required, but generally less distinctive.

Examples:

```text
excavation
trenches
surface
foundation
filling
treatment
```

### Priority 5 — Numeric values

Numbers can be used when necessary:

```text
30
1.5
10
80
300
600
98%
```

But:

> **A numeric value must not automatically win merely because it happens to be unique.**

Numbers are still valid keywords if they are required for identification, but they should generally have lower priority than meaningful descriptive words.

### Priority 6 — Generic/structural words

Avoid these unless nothing better is available:

```text
work
earth
soil
material
description
including
complete
required
depth
lead
lift
and
the
for
in
of
```

---

# 6. The actual selection rule

For every item:

```text
Generate all possible 1-word combinations
        ↓
Keep combinations that uniquely identify the item
        ↓
Rank them by keyword priority
        ↓
If none exist:
        ↓
Generate all 2-word combinations
        ↓
Keep unique combinations
        ↓
Rank them
        ↓
If none exist:
        ↓
Generate all 3-word combinations
        ↓
...
```

Therefore:

### First priority

**Fewest number of words**

### Second priority

**Most useful/discriminating words**

### Third priority

**Human meaningfulness**

### Fourth priority

**Avoid accidental numeric/code identifiers where a meaningful alternative exists**

---

# 7. What makes a good identifier?

A good keyword set should satisfy all three:

```text
                 GOOD IDENTIFIER
                       │
          ┌────────────┼────────────┐
          ↓            ↓            ↓
      UNIQUE        USEFUL       DISTINCTIVE
          │            │            │
   identifies     meaningful    separates item
   one item       to human      from siblings
```

So the system should prefer:

> **Words that uniquely identify a row, are actually useful to a human, and distinguish the row from its siblings.**

---

# 8. Important distinction: global uniqueness vs sibling uniqueness

A word may be unique globally but still be a poor identifier.

For example:

```text
98%
```

may occur only once in the entire dataset.

But:

```text
fly
```

may occur in multiple items.

Therefore:

```text
98% → mathematically unique
fly → not mathematically unique
```

However, if the purpose is human identification, the system should recognize that:

```text
fly + ash + filling
```

may be a much better identifier than:

```text
98%
```

Thus the algorithm should **not throw away mathematical uniqueness**, but should rank meaningful combinations higher.

---

# 9. Parent–child items must be compared as complete items

For example:

```text
Parent:
Anti-termite treatment

Children:
External wall
Below apron
Existing floors
Existing masonry
Wood contact points
```

The comparison should effectively become:

```text
Anti-termite + external wall
Anti-termite + below apron
Anti-termite + existing floors
Anti-termite + existing masonry
Anti-termite + wood contact
```

Then the algorithm determines the minimum useful identifier.

Similarly:

```text
Fly ash
   +
Supply and stacking
```

and

```text
Fly ash
   +
Earth filling
```

are two different complete items.

Therefore:

```text
fly
```

is insufficient.

The algorithm may find:

```text
supply + stacking
```

for one item and:

```text
filling + fly
```

for the other, depending on the complete dataset.

---

# 10. Final rule set

I would define the system with these rules:

```text
RULE 1
Treat Parent Description + Child Qualifier as one complete identifiable item.

RULE 2
Only keywords from the current allowed keyword list may be selected.

RULE 3
For each item, test 1-word combinations first.

RULE 4
If no 1-word combination uniquely identifies the item,
test every 2-word combination.

RULE 5
Continue progressively:
1 word → 2 words → 3 words → ... → N words.

RULE 6
A combination is unique only if exactly one complete item
matches all keywords in that combination.

RULE 7
Stop at the first word-count level where a unique combination exists.

RULE 8
If several combinations have the same minimum number of words,
rank them using the keyword priority hierarchy.

RULE 9
Prefer meaningful discriminating words over generic words.

RULE 10
Numeric values and codes are valid keywords, but must not
automatically receive higher priority merely because they are unique.

RULE 11
Prefer identifiers that are useful to a human and distinguish
the item from its siblings.

RULE 12
If no combination can uniquely identify the item, report:
"NOT UNIQUELY IDENTIFIABLE."

RULE 13
Do not add a new keyword merely to make the result unique
unless the system enters the "additional keyword" process.

RULE 14
If the current keyword list cannot uniquely identify an item,
find the smallest additional keyword(s) required to make it unique.

RULE 15
The final objective is:
minimum number of useful discriminating keywords,
not merely minimum number of arbitrary keywords.
```

## Overall logic

```text
CURRENT KEYWORD LIST
        ↓
COMPLETE ITEMS
(Parent + Child)
        ↓
CANDIDATE KEYWORDS
        ↓
┌─────────────────────────────┐
│  1 WORD                     │
│  Is any useful combination  │
│  uniquely identifying?      │
└──────────────┬──────────────┘
               │ NO
               ↓
┌─────────────────────────────┐
│  2 WORDS                    │
│  Is any useful combination  │
│  uniquely identifying?      │
└──────────────┬──────────────┘
               │ NO
               ↓
┌─────────────────────────────┐
│  3 WORDS                    │
│  Is any useful combination  │
│  uniquely identifying?      │
└──────────────┬──────────────┘
               │
              ...
               ↓
       UNIQUE COMBINATION
               ↓
    RANK BY KEYWORD QUALITY
               ↓
       BEST IDENTIFIER
               │
               │ no combination
               ↓
   FIND SMALLEST ADDITIONAL
        KEYWORD(S)
               ↓
      UPDATED KEYWORD LIST
```

### The core principle in one sentence

> **Find the smallest number of keywords that uniquely identify the complete item, but among equally small combinations, prefer keywords that are meaningful, discriminating, human-useful, and relevant to distinguishing the item from its siblings rather than accidental numeric or generic words.**
