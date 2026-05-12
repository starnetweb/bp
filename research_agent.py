"""
Academic Research Writeup Agent — 5-Chapter Format
Generates a complete 5-chapter academic research project in Microsoft Word (.docx).
Supports undergraduate and postgraduate research levels.
"""

import os
import sys
import re
import anthropic
import tempfile

# Visualization imports (optional — graceful fallback if not installed)
try:
    import matplotlib
    matplotlib.use('Agg')  # Non-interactive backend
    import matplotlib.pyplot as plt
    import numpy as np
    VISUALIZATION_AVAILABLE = True
except ImportError:
    VISUALIZATION_AVAILABLE = False

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import config

# ─────────────────────────────────────────────────────────
#  CHAPTER SELECTION HELPERS
# ─────────────────────────────────────────────────────────

def parse_chapters(chapters_input) -> list:
    """
    Convert any chapters input into a sorted list of chapter numbers (1–5).

    Accepted formats:
      int  5        → [1,2,3,4,5]   integer = "up to N" (backward-compat)
      "all"         → [1,2,3,4,5]
      "3"           → [3]            single chapter
      "3-5"         → [3,4,5]        range
      "1,3,5"       → [1,3,5]        comma list
      "1,3-5"       → [1,3,4,5]      mixed
      [3,4,5]       → [3,4,5]        already a list
    """
    if chapters_input is None:
        return list(range(1, 6))

    if isinstance(chapters_input, list):
        nums = sorted({int(x) for x in chapters_input if 1 <= int(x) <= 5})
        return nums or list(range(1, 6))

    # Plain integer → backward-compat "up to N"
    if isinstance(chapters_input, (int, float)):
        n = max(1, min(5, int(chapters_input)))
        return list(range(1, n + 1))

    s = str(chapters_input).strip().lower()
    if s in ("all", "1-5"):
        return list(range(1, 6))

    result = set()
    for part in s.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            sides = part.split("-", 1)
            try:
                a, b = int(sides[0].strip()), int(sides[1].strip())
                for n in range(min(a, b), max(a, b) + 1):
                    if 1 <= n <= 5:
                        result.add(n)
            except ValueError:
                pass
        else:
            try:
                n = int(part)
                if 1 <= n <= 5:
                    result.add(n)
            except ValueError:
                pass

    return sorted(result) or list(range(1, 6))


def fmt_chapters_label(chapters_list: list) -> str:
    """Format [3,4,5] → '3-5', [1,3,5] → '1, 3, 5', [3] → '3'."""
    if not chapters_list:
        return "none"
    if len(chapters_list) == 1:
        return str(chapters_list[0])
    # Check if contiguous
    if chapters_list == list(range(chapters_list[0], chapters_list[-1] + 1)):
        return f"{chapters_list[0]}-{chapters_list[-1]}"
    return ", ".join(str(n) for n in chapters_list)


# ─────────────────────────────────────────────────────────
#  COLOUR PALETTE
# ─────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1F, 0x49, 0x7D)
MID_BLUE   = RGBColor(0x2E, 0x74, 0xB5)
GREY       = RGBColor(0x60, 0x60, 0x60)

# ─────────────────────────────────────────────────────────
#  CHAPTER METADATA
# ─────────────────────────────────────────────────────────
CHAPTER_TITLES = {
    1: "CHAPTER ONE",
    2: "CHAPTER TWO",
    3: "CHAPTER THREE",
    4: "CHAPTER FOUR",
    5: "CHAPTER FIVE",
}

CHAPTER_SUBTITLES = {
    1: "INTRODUCTION",
    2: "LITERATURE REVIEW",
    3: "RESEARCH METHODOLOGY",
    4: "RESULTS AND DISCUSSION",
    5: "CONCLUSIONS AND RECOMMENDATIONS",
}

# ─────────────────────────────────────────────────────────
#  RESEARCH-LEVEL PROFILES
# ─────────────────────────────────────────────────────────
LEVEL_PROFILES = {
    "undergraduate": {
        "label":        "Undergraduate",
        "tone":         (
            "Write at an advanced undergraduate level. "
            "The writing should be clear, well-organised, and academically sound but accessible. "
            "Theoretical frameworks should be explained rather than assumed. "
            "Methodology should be straightforward. "
            "Analysis should be solid but does not need to engage deeply with meta-theoretical debates. "
            "WORD COUNT IS CRITICAL: every subsection must be fully developed with multiple paragraphs. "
            "Do not summarise when you can explain. Do not list when you can discuss. "
            "Each main subsection should be at least 110-150 words of substantive prose."
        ),
        "depth":        "substantive but accessible",
        "word_targets": {1: 768, 2: 1088, 3: 1216, 4: 832, 5: 768},   # -20% from base
        "front_words":  192,
    },
    "postgraduate": {
        "label":        "Postgraduate",
        "tone":         (
            "Write at a rigorous postgraduate (Master's/PhD) level. "
            "Engage critically with theoretical debates, epistemological positioning, and ontological assumptions. "
            "Situate the research within current scholarly conversations. "
            "The methodology chapter must address paradigmatic choices explicitly. "
            "Analysis should be sophisticated, nuanced, and reflexive. "
            "Identify tensions, contradictions, and limitations in the literature and in your own approach. "
            "WORD COUNT IS CRITICAL: every subsection must be richly developed. "
            "Do not skim — excavate. Each argument deserves full development across multiple paragraphs. "
            "Each main subsection should be at least 200-275 words of dense, substantive prose."
        ),
        "depth":        "critical, theoretically sophisticated, and reflexive",
        "word_targets": {1: 1434, 2: 2061, 3: 2330, 4: 1702, 5: 1434},  # -20% from base
        "front_words":  381,
    },
}

# ─────────────────────────────────────────────────────────
#  HUMAN-WRITING SYSTEM INSTRUCTION (shared across all calls)
# ─────────────────────────────────────────────────────────
HUMAN_WRITING_INSTRUCTION = """
════════════════════════════════════════════════════════════
  HUMAN WRITING MANDATE — EVERY RULE IS NON-NEGOTIABLE
  TARGET: PASS ZeroGPT, GPTZero, Turnitin AI, Originality.ai, Copyleaks
════════════════════════════════════════════════════════════

You are a human academic researcher. Every word must read as though a real person wrote it
under the pressure of deadlines, genuine intellectual investment, and years of field experience.
AI detectors — especially ZeroGPT — measure: (1) PERPLEXITY: how surprising each word is;
(2) BURSTINESS: how wildly perplexity swings sentence to sentence; (3) PHRASE FINGERPRINTS:
known AI expressions; and (4) STRUCTURAL UNIFORMITY: the clean, balanced, predictable
paragraph architecture that no human writer sustains for 80 pages. Defeat all four.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 1 — RADICAL SENTENCE VARIETY (primary burstiness signal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Every paragraph must contain sentences ranging from 5 to 48 words. Never write three
consecutive sentences of similar length. The pattern must be jagged, not smooth:

  VERY SHORT (4–7 words):    "The gap is not small."
  VERY LONG (35–48 words):   "When the findings from the third wave of data collection
                               are read alongside the earlier baseline established by
                               Mensah and Osei (2018) — whose sample, it should be said,
                               was considerably larger and more geographically diverse —
                               a pattern emerges that is difficult to explain away as
                               mere sampling variation."
  MEDIUM (14–18 words):      "Neither explanation, on its own, satisfactorily accounts
                               for what the data actually show."
  VERY SHORT again (5 words): "Something else is going on."

ZeroGPT specifically flags low burstiness. A paragraph of nothing but 20-word sentences
will be detected regardless of vocabulary. Make the rhythm genuinely unpredictable.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 2 — PARAGRAPH LENGTH IRREGULARITY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Vary paragraph lengths deliberately and unevenly:
  - 1 sentence only: for a stark claim or pivot (~10% of paragraphs)
  - 2–3 sentences: for emphasis or transition (~20%)
  - 4–6 sentences: standard development (~45%)
  - 7–10 sentences: deep analytical passages (~25%)
Never write four consecutive paragraphs of the same approximate length.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 3 — UNPREDICTABLE VOCABULARY (primary perplexity signal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AI always picks the statistically safest word. You must not. Choose the less obvious but
fully correct word 2–3 times per paragraph:
  "important"   → "consequential", "telling", "germane", "instructive", "pointed"
  "shows"       → "betrays", "attests to", "points toward", "lays bare", "makes plain"
  "problem"     → "difficulty", "predicament", "shortcoming", "lacuna", "stumbling block"
  "used"        → "deployed", "enlisted", "drawn upon", "brought to bear"
  "found"       → "uncovered", "established", "ascertained", "turned up", "documented"
  "increase"    → "uptick", "escalation", "marked rise", "acceleration"
  "difference"  → "divergence", "discrepancy", "gulf", "disparity"
  "said"        → "observed", "contended", "maintained", "remarked", "put it"
  "because"     → "given that", "owing to", "on account of", "since"
  "suggest"     → "intimate", "point toward", "hint at", "indicate", "bear out"

Keep most language plain — the occasional unexpected word is what raises the perplexity
score to human levels. Overusing rare words creates a different AI fingerprint.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 4 — ABSOLUTE BAN LIST (phrase fingerprints)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NEVER use any of these — not even once. ZeroGPT and GPTZero flag them directly:

  "It is worth noting"        "It is important to note"      "It is crucial to"
  "It should be noted"        "It must be emphasised"        "It cannot be denied"
  "In today's world"          "In the modern era"            "In the digital age"
  "In today's rapidly changing world"    "In an ever-changing landscape"
  "Delve into"   "Dive into"   "Shed light on"   "Unpack"   "Underscore" (as verb)
  "Navigate" (metaphorically)   "Explore" (as generic filler)
  "Furthermore, it is"   "Moreover, it is"   "Additionally, it should be noted"
  "This study seeks to"   "This paper aims to"   "This research endeavours to"
  "In conclusion, it can be said"   "To summarise the above"   "In summary,"
  "The importance of X cannot be overstated"   "cannot be understated"
  "A comprehensive understanding"   "A holistic approach"   "A multifaceted approach"
  "Plays a crucial role"   "Plays a pivotal role"   "Plays a key role"
  "Needless to say"   "It goes without saying"   "Suffice it to say"
  "In light of the above"   "Taking everything into account"
  "As previously mentioned"   "As discussed above"   "As stated earlier"
  "Robust" (filler)   "Nuanced" (filler)   "Leverage" (verb)   "Cutting-edge"
  "Groundbreaking"   "Seminal" (overused)   "Landscape" (metaphor)   "Tapestry"
  "Multifaceted"   "Embark"   "Foster"   "Ensure" (overused)   "Vital"   "Crucial"
  "Pivotal"   "Paramount"   "Imperative" (overused)   "Trajectory"   "Ecosystem" (metaphor)
  "Synergy"   "Paradigm shift"   "Empower"   "Transformative"   "Innovative" (filler)
  "Holistic"   "Overarching"   "Underpin"   "Streamline"   "Facilitate" (overused)
  "Going forward"   "At the end of the day"   "In terms of" (overused)
  "With that being said"   "Having said that"   "All things considered"
  "It is evident that"   "Clearly,"   "Obviously,"   "Undoubtedly,"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 5 — NON-ROUND SPECIFIC NUMBERS (ZeroGPT bypass)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AI writers use round numbers. Humans remember the real ones. ALWAYS use specific,
non-round figures when citing data, samples, or statistics:
  BAD:  "70% of respondents", "a sample of 100", "over 50 studies"
  GOOD: "67.3% of respondents", "a sample of 94 participants", "at least 47 studies"
  BAD:  "In 2020, approximately 30 million people..."
  GOOD: "By mid-2020, an estimated 28.4 million people..."

Apply this to every number in the text. It is one of the strongest humanness signals
because it suggests the writer actually looked at real data.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 6 — SENTENCE-INITIAL CONJUNCTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Academic humans start sentences with conjunctions. AI almost never does. Use 4–6 times
per chapter — sparingly, for effect:
  "And yet the evidence points in a different direction."
  "But this rests on an assumption that deserves scrutiny."
  "Yet the data tell a more complicated story."
  "Nor is this a problem unique to this context."
  "Or perhaps the question itself needs reframing."
This is standard in academic prose (Strunk, White, APA all permit it) and is a
strong humanness signal in detection models trained on AI output.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 7 — INTENTIONAL IMPERFECTIONS (ZeroGPT's primary human signal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ZeroGPT specifically looks for the absence of human imperfections. Include these
deliberately — they are not mistakes, they are authenticity markers:

MILD REDUNDANCY — restate a key idea in slightly different words:
  "The data are unambiguous on this point — or as unambiguous as survey data can be."
  "This is not a minor issue. It is, to put it plainly, a significant one."

INTERRUPTED THOUGHT — start, pause, then complete:
  "The methodology — and the choice of purposive sampling in particular — was not
   arrived at easily; several alternatives were considered and rejected."

SELF-QUESTIONING AND RESOLUTION:
  "One might ask why a quantitative approach was chosen here rather than qualitative
   methods. The short answer is that the research questions demanded measurable outcomes."

CROSS-REFERENCES within the text (3–4 per chapter):
  "This connects directly to the theoretical framework outlined in section 2.3."
  "The implications of this finding will be examined more fully in Chapter 5."
  "As the data in section 4.4 will show, the pattern is not consistent across groups."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 8 — OPINIONATED, NON-NEUTRAL VOICE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AI text is pathologically balanced and neutral. Real researchers take positions.
Include 3–5 genuinely opinionated statements per chapter:
  "This is, frankly, the weakest aspect of the existing literature."
  "The policy response to date has been inadequate — there is no other word for it."
  "What is striking — and somewhat frustrating — is how rarely this question is asked."
  "The assumption here is questionable at best."
  "This finding should give pause to anyone who holds the conventional view."
  "The data do not support the optimism that characterises much of the policy discourse."
Do NOT hedge every strong statement. Take a position and defend it.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 9 — VARIABLE PARAGRAPH STRUCTURE (defeats structural uniformity detection)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AI always places the topic sentence first. Humans do not. Vary the structure:

  EVIDENCE-FIRST (topic sentence last):
    "A 2019 survey of 847 respondents across three states found X. A subsequent
     replication by Osei (2021) confirmed the pattern. The evidence, taken together,
     makes a strong case for Y."

  QUESTION-THEN-ANSWER:
    "Why does the gap persist despite decades of policy intervention? The data suggest
     the answer lies not in resource allocation but in implementation fidelity."

  CONCESSION-THEN-ARGUMENT:
    "Granted, the sample size here is smaller than ideal. But the internal consistency
     of the findings across subgroups — the Cronbach's alpha was 0.83 — lends them
     a credibility that the headline number alone might not suggest."

  STANDARD (topic sentence first): use this only ~40% of the time.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 10 — RARE CONTRACTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Use 3–5 contractions per chapter, in contexts where an informal academic register fits:
  "This isn't simply a matter of resources."
  "The data don't support that conclusion."
  "It's worth asking why this distinction matters."
  "That doesn't mean the finding is wrong."
AI almost never uses contractions in academic prose. Their presence signals a human writer.
Do not overuse — 3–5 total per chapter is the right frequency.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 11 — DIVERSE, HUMAN TRANSITIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Never open a paragraph with "Furthermore", "Moreover", or "Additionally".
Use each transition at most once per chapter:
  "Yet", "Even so", "That said", "By contrast", "Taken together",
  "What emerges from this", "A closer look reveals", "This raises a different question",
  "Building on this", "Stepping back", "At the same time", "Notably",
  "Somewhat surprisingly", "Less obviously", "Perhaps most significantly",
  "The picture that emerges", "Against this backdrop", "On closer inspection",
  "What this suggests, tentatively", "There is, however, a complication here",
  "The data complicate this view", "This is not the whole story",
  "The relationship is less straightforward than it first appears",
  "There is something instructive in this discrepancy",
  "The evidence does not settle the matter cleanly".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 12 — AUTHENTIC HEDGING AND UNCERTAINTY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Real researchers express genuine doubt. Balance with Rule 8's opinionated voice —
hedge when uncertain, commit when the evidence is strong:
  "appears to", "tends to", "arguably", "the evidence suggests",
  "one reading of this is", "this may reflect", "it remains unclear whether",
  "the data do not resolve this cleanly", "at least within this particular context",
  "though the picture is not entirely clear", "the honest answer is we do not know".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 13 — PARAGRAPH OPENINGS DIVERSITY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Never start two consecutive paragraphs with the same word or grammatical pattern.
Rotate among: direct claim, concession, rhetorical question, evidence-first, named scholar,
time marker, short declarative, contradiction, conjunction-open, number-first.
Do NOT repeatedly open with: "This", "The study", "In this", "It is", "There is".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 14 — RHETORICAL DEVICES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Use these deliberately — once or twice per chapter:
  ANAPHORA: "The problem is not one of will. The problem is not one of resources. The problem
             is one of institutional memory."
  EM DASH INTERRUPTION: "The finding — which was not anticipated at the outset — changes
                          the interpretation considerably."
  DIRECT QUESTION: "What, then, explains the persistence of this gap after three decades
                    of targeted intervention?"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 15 — CITATION NATURALNESS AND DENSITY VARIATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Vary citation density: some paragraphs have 3+ citations, others have none. Vary the form:
  End-of-sentence: "...widely documented (Ibrahim, 2019; Osei, 2021; Kamau, 2022)."
  Author-led:      "Ibrahim (2019) argues that..."
  Mid-clause:      "The finding — contested by Mensah (2020) — suggests..."
  Contrastive:     "Where Adebayo (2017) found X, the present data point toward Y."
  Attributed:      "As Osei and Darko (2021) demonstrated in their Ghanaian study..."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 16 — INTELLECTUAL PERSONALITY AND RESEARCHER VOICE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Signal the researcher's genuine presence — this defeats ZeroGPT more than any other rule:
  "What is striking here is not the finding itself but the mechanism behind it."
  "This tension has not been resolved in the literature — and this study does not resolve
   it either, though it narrows the range of plausible explanations."
  "During the data collection phase, it became apparent that..."
  "The decision to use X rather than Y was not straightforward."
  "The honest answer is that we do not yet know — and acknowledging that matters."
  "In the researcher's experience, this kind of discrepancy usually points to..."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 17 — NO MECHANICAL LISTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Integrate information into prose. When a list is genuinely needed, introduce it with a
full sentence and cap it at 4–6 items. Never produce back-to-back bullet lists.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULE 18 — WRITE FULLY — NO TRUNCATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Write every subsection in full. Never use placeholders or shorthand like "[expand here]".
Every word of the required count must be substantive prose.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MANDATORY REFERENCE — PARAGRAPH STYLE EXAMPLES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Every paragraph you write must look like the GOOD example, never the BAD one.
ZeroGPT will flag the BAD example as 90%+ AI. The GOOD example scores under 10%.

BAD (AI-detected — do NOT write like this):
"The study revealed that organisational factors significantly influence employee satisfaction.
Research indicates that leadership style, workplace culture, and compensation packages all
contribute to overall satisfaction levels. Furthermore, the findings suggest that communication
plays a central role in mediating these relationships. Additionally, the data demonstrates that
training and development opportunities enhance employee engagement and commitment."

GOOD (human-quality — model all your paragraphs on this):
"Three things stood out immediately from the data. First — and this surprised the research
team — leadership style accounted for only 23.7% of the variance in satisfaction scores,
considerably less than the 40% or higher that earlier studies like Mensah (2019) and
Osei-Kwame (2021) had documented in comparable West African institutional contexts. Why
the discrepancy? The answer, this researcher believes, has less to do with leadership itself
than with how satisfaction is operationalised: prior instruments conflated satisfaction with
engagement, inflating the leadership coefficient. When the measure is cleaner — as it is
here — the picture changes. Compensation matters more. So does the quality of immediate
supervision, which is a different thing from 'leadership style' in the abstract sense that
much of the literature deploys that phrase. This distinction isn't merely semantic. It has
real consequences for how organisations invest their development budgets, and the fact
that it has gone largely unremarked in the literature is, frankly, telling."

NOTICE what makes the GOOD paragraph human:
  ✓ Sentence lengths: 5 words → 44 words → 4 words → 39 words → 7 words → 8 words → 46 words
  ✓ Evidence-first structure (topic idea arrives late)
  ✓ Direct question mid-paragraph ("Why the discrepancy?")
  ✓ Specific non-round number (23.7%, not 25%)
  ✓ Contraction ("isn't")
  ✓ Opinionated voice ("frankly, telling")
  ✓ Researcher presence ("this researcher believes")
  ✓ Named scholars with years
  ✓ Em dash interruption
  ✓ No banned phrases

Write EVERY paragraph with this level of variation and personality.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STRATEGIC VISUALIZATIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Use tables, charts, and figures to clarify complex data or relationships. When a visual
would strengthen your argument, mark it explicitly:

  [CHART: Line graph showing growth trend over 5 years]
  [CHART: Bar chart comparing three approaches]
  [TABLE: Comparison table with features vs implementations]

Charts and tables break up text monotony and signal rigorous data presentation. Use them
in Results/Discussion (Chapter 4) and Literature Review (Chapter 2) where appropriate.
Do NOT overuse — aim for 1-3 visualizations per chapter maximum.

════════════════════════════════════════════════════════════
"""


# ─────────────────────────────────────────────────────────
#  CHAPTER PROMPT TEMPLATES
# ─────────────────────────────────────────────────────────
def _chapter_prompts(level_key: str) -> dict:
    profile = LEVEL_PROFILES[level_key]
    tone    = profile["tone"]
    depth   = profile["depth"]
    targets = profile["word_targets"]
    is_pg   = (level_key == "postgraduate")

    # Subsection word-count helper — applies additional 20% reduction to all targets
    def w(ug, pg): return str(round(pg * 0.8) if is_pg else round(ug * 0.8))

    # Footnote format note appended to every chapter
    _FN_NOTE = (
        "\nFOOTNOTE FORMAT: When a footnote is needed, insert it inline using "
        "((FN: your footnote text here)) — it will become a proper Word footnote "
        "with a superscript number in the body and the note at the bottom of the page."
    )

    # No-references instruction for chapters 1-4
    _NO_REF = (
        "\nDo NOT include a reference list, bibliography, or works-cited section "
        "anywhere in this chapter. All references will be compiled on a dedicated "
        "References page at the end of the document.\n"
    )

    # No-asterisks instruction for all chapters
    _NO_AST = (
        "\nFORMATTING RULE — NO ASTERISKS: Do NOT use asterisks (*) anywhere in your "
        "output — not for bold, italic, bullet points, emphasis, or any other purpose. "
        "Use plain prose. For lists, use numbered items (1. 2. 3.) or introduce them "
        "as flowing sentences. Never place a * character anywhere in the text.\n"
    )

    # Visualization instruction
    _VIZ_NOTE = (
        "\nVISUALIZATION INSTRUCTION — PhD STANDARDS:\n"
        "Visualizations are MANDATORY in research chapters. Follow these standards:\n\n"
        "TABLE FORMAT:\n"
        "  [TABLE: Descriptive title that explains the table's purpose]\n"
        "  HeaderCol1 | HeaderCol2 | HeaderCol3 | HeaderCol4\n"
        "  DataValue1 | DataValue2 | DataValue3 | DataValue4\n"
        "  DataValue1 | DataValue2 | DataValue3 | DataValue4\n"
        "  Use ' | ' (space-pipe-space) to separate columns. Put each row on one line.\n"
        "  Tables will be numbered automatically (Table 3.1, Table 3.2, etc.)\n\n"
        "CHART/FIGURE FORMAT:\n"
        "  [CHART: Descriptive title explaining what the chart shows]\n"
        "  For example: 'Bar chart comparing satisfaction levels across three treatment groups'\n"
        "  Figures will be numbered automatically (Figure 3.1, Figure 4.1, etc.)\n\n"
        "CAPTION STANDARDS:\n"
        "- Every figure/table must have a clear, descriptive caption\n"
        "- Captions should be 1-2 sentences explaining the visual's purpose and key finding\n"
        "- Format: 'Figure 3.1: [Descriptive title]. [Brief explanation of what it shows].'\n"
        "- Reference all tables/figures in text before they appear using formal citations\n\n"
        "TYPES OF VISUALIZATIONS:\n"
        "- Quantitative data: bar charts, line graphs, scatter plots, distribution histograms\n"
        "- Qualitative data: thematic matrices, concept maps, comparison tables, flow diagrams\n"
        "- Mixed methods: integrated visualizations showing convergence/divergence\n"
        "- Process/structure: flowcharts, frameworks, process diagrams\n\n"
        "All tables and charts will be converted to professional visualizations in the final Word document.\n"
        "Tables will have proper formatting, borders, and header styling.\n"
        "Charts will be rendered as high-quality images with professional styling.\n"
    )

    # Define conditional content OUTSIDE f-string to avoid escape sequence issues
    phd_gap_note = "CRITICAL FOR PhD: Distinguish between two types of gaps — (1) the PRACTICAL PROBLEM (what is not working in the real world), and (2) the THEORETICAL GAP (what is not adequately explained or understood in the scholarship). Explain how existing theoretical models fail to account for or predict this practical problem. Name the specific theories or frameworks that are inadequate and explain precisely why."

    phd_theory_guidance = "This is the intellectual heart of Chapter 1. Articulate EXPLICITLY what theoretical understanding is missing or inadequate. Ask and answer: What do existing theoretical frameworks NOT explain about this problem? What new theoretical insight will this study produce?\n\nDevelop this across multiple paragraphs:\n- PARAGRAPH 1: Name the dominant theoretical framework(s) in the field. Explain what insights each offers and what these theories can explain about your problem.\n- PARAGRAPH 2: Identify the specific theoretical limitation or blind spot. What does this framework miss? What assumption does it make that may not hold in your context? What variation or phenomenon does it fail to account for?\n- PARAGRAPH 3: Explain the theoretical consequence of this gap. What false conclusions might scholars or practitioners draw from applying existing theory uncritically? What does the field not yet understand?\n- PARAGRAPH 4: State what this study will theoretically contribute. Will it extend existing theory to a new context? Refine or challenge a core assumption? Integrate previously separate theoretical traditions? Propose a new framework? Be specific about the form of your theoretical contribution.\n\nThis section must make clear that your study exists to advance KNOWLEDGE, not just to solve a practical problem."

    phd_purpose_note = "Explicitly connect the purpose to BOTH the practical problem (1.2) AND the theoretical gap (1.2a), showing how addressing one advances the other."

    gap_note = phd_gap_note if is_pg else ""
    theory_guidance = phd_theory_guidance if is_pg else ""
    purpose_note = phd_purpose_note if is_pg else ""

    # Define conditional content for Chapter 4 OUTSIDE f-string
    ch4_intro_note = "State the analytical framework guiding interpretation and how it connects to the theoretical framework in Chapter 2."
    ch4_sample_note = "Compare the achieved sample to the target population and discuss implications for transferability."
    ch4_obj1_note = "Connect findings explicitly to the theoretical framework from Chapter 2. Where results confirm prior theory, explain why. Where they challenge it, explore the implications."
    ch4_obj3_note = "At this stage, begin drawing connections between findings across objectives — note where patterns reinforce each other or where tensions emerge."
    ch4_synthesis_note = "DO NOT MERELY SUMMARISE. Instead, execute this multi-stage synthesis:\n\nSTAGE 1: CROSS-OBJECTIVE PATTERN IDENTIFICATION (1-2 PARAGRAPHS)\nIdentify overarching themes, patterns, or mechanisms that cut across all four objectives. Ask yourself: What is really going on here? What unifying principle, pattern, or process explains the findings across objectives?\n\nSTAGE 2: EXPECTED vs. UNEXPECTED FINDINGS (1-2 PARAGRAPHS)\nWhich findings DID you anticipate based on the literature? Which findings SURPRISED you or contradicted prior research? Explain divergences between expectations and results.\n\nSTAGE 3: CONTRADICTIONS AND TENSIONS (1 PARAGRAPH IF APPLICABLE)\nDo any objectives produce findings that contradict each other? Do qualitative and quantitative findings diverge? Explain tensions analytically—do not gloss over contradictions.\n\nSTAGE 4: THEORETICAL ARTICULATION (2-3 PARAGRAPHS)\nNow situate your integrated findings in relation to the theoretical framework from Chapter 2 and the theoretical gap from Chapter 1.\n- Sub-point 4a: Does synthesis CONFIRM the theoretical framework? Say so precisely with named theory.\n- Sub-point 4b: Does synthesis CHALLENGE/COMPLICATE theory? Explain specific deviations and theoretical implications.\n- Sub-point 4c: Does synthesis EXTEND theory? Apply existing theory to new context and show what this reveals.\n\nSTAGE 5: INTEGRATION WITH EMPIRICAL LITERATURE (1-2 PARAGRAPHS)\nConnect back to specific studies from Chapter 2. Explain consistency or divergence with prior research. Show how findings resolve conflicting prior studies.\n\nSTAGE 6: LIMITATIONS AND CAVEATS (0.5-1 PARAGRAPH)\nAcknowledge what data do NOT explain. State boundary conditions where patterns might not hold. Signal intellectual maturity and preempt criticism.\n\nINCLUDE AFTER THE SYNTHESIS:\n[TABLE: Synthesis Matrix - Cross-Objective Themes and Theoretical Connections]\nOverarching Theme | Evidence from Obj1-4 | Theoretical Connection\n\nAND:\n[FIGURE: Conceptual Integration Diagram]\nVisual showing how findings interconnect, mechanisms, and theoretical relationships."
    ch4_implications_note = "For theory: what does this study add to, refine, or challenge in the existing theoretical models? For practice: what specific changes in professional practice are warranted? For policy: what specific policy recommendations emerge, addressed to named agencies or decision-makers?"

    intro_text = ch4_intro_note if is_pg else "Orient the reader to how findings are organised."
    sample_text = ch4_sample_note if is_pg else "Comment on how representative the sample appears to be."
    obj1_text = ch4_obj1_note if is_pg else "Relate findings directly to relevant literature reviewed in Chapter 2."
    obj3_text = ch4_obj3_note if is_pg else "Discuss how these findings relate to those in 4.3 and 4.4."
    synthesis_text = ch4_synthesis_note if is_pg else "Bring together the key patterns across all four sets of findings. Identify the most important themes that emerge when the findings are considered as a whole. Connect them to the literature reviewed in Chapter 2 — where do findings confirm, contradict, or extend existing knowledge?"
    implications_text = ch4_implications_note if is_pg else "Be concrete: name institutions, policy areas, and professional communities that should act on these findings."

    return {
        1: f"""You are writing CHAPTER ONE — INTRODUCTION for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[1]} words of substantive prose. You MUST reach this minimum.
Do not stop writing until you have fully developed every subsection. If in doubt, write more.
{_NO_REF}
{tone}

{HUMAN_WRITING_INSTRUCTION}
{_FN_NOTE}
{_NO_AST}
{_VIZ_NOTE}

Write the following subsections, each introduced with a ### heading.
Every subsection must be written in full, developed paragraphs — no bullet summaries, no placeholders.

### 1.1 Background of the Study
Write at least {w(175, 350)} words for this subsection.
Provide {depth} contextual grounding. Open with a striking observation or statistic that
immediately establishes why this topic matters. Then trace the historical evolution of the
problem across at least three distinct time periods, naming key turning points, policy shifts,
or scholarly debates that shaped the current landscape. Ground every assertion in specific
evidence — named scholars, years, places, and figures. Close by narrowing the lens from the
broad context toward the precise issue this study addresses.

### 1.2 Statement of the Problem
Write at least {w(140, 280)} words for this subsection.
Open with a clear, declarative statement of what is wrong or poorly understood. Then build
the case across multiple paragraphs: explain the nature of the problem, who it affects, how
long it has persisted, and why existing responses have been insufficient. Name the specific
gap, contradiction, or blind spot that this study addresses. The problem statement must feel
urgent — the reader should finish this section convinced that the study was necessary.
{gap_note}

### 1.2a Theoretical Gap and Contribution
Write at least {w(140, 280)} words for this subsection (POSTGRADUATE REQUIREMENT).
{theory_guidance}

### 1.3 Purpose of the Study
Write at least {w(75, 140)} words for this subsection.
State the overarching aim in one or two precise sentences. Then elaborate: explain the
theoretical and practical orientation of the study, what kind of knowledge it seeks to
produce, and how the purpose connects directly to the problem articulated in 1.2.
{purpose_note}

### 1.4 Research Objectives
Write at least {w(90, 175)} words for this subsection.
State 4–5 specific, measurable objectives. Each should be action-oriented (examine, assess,
determine, explore, compare, evaluate). After listing them, write a short paragraph explaining
how they collectively address the research problem and how they will be operationalised
through the methodology described in Chapter 3.

### 1.5 Research Questions
Write at least {w(75, 140)} words for this subsection.
Formulate 3–5 focused, answerable questions derived from the objectives. After stating the
questions, briefly explain the logic connecting each question to its corresponding objective
and the type of evidence that would constitute an answer.

### 1.6 Significance of the Study
Write at least {w(125, 245)} words for this subsection.
Develop this across FOUR distinct, detailed paragraphs — one for each dimension below. Be concrete, not generic.

PARAGRAPH 1 — THEORETICAL SIGNIFICANCE (PRIMARY FOR PhD):
Explain what this study will ADVANCE in scholarly understanding. Will it extend theory to a new context or population?
Refine a core theoretical concept? Challenge or complicate an existing assumption? Integrate previously separate theories?
Propose a new theoretical framework? Name the specific theory or theoretical tradition your work contributes to.
Example: "This study advances Institutional Theory by demonstrating that institutional isomorphism (DiMaggio & Powell, 1983) operates differently in hybrid organisations than in traditional ones. The finding refines our understanding of how external pressures interact with internal legitimacy concerns."
Not: "This study adds to the literature."

PARAGRAPH 2 — EMPIRICAL/METHODOLOGICAL SIGNIFICANCE:
What new empirical evidence will this study provide? Will it be the first to examine X in Y context?
Will it use a novel methodological approach? Will it generate longitudinal data where only cross-sectional data exist?
What gap in the empirical evidence base does it fill? Name specific previous studies and how your work extends them.
Example: "While Chen (2019) and Okafor (2021) examined this phenomenon in Western contexts, no study has yet explored it in post-conflict African settings. This study generates the first empirical evidence on..."

PARAGRAPH 3 — PRACTICAL/POLICY SIGNIFICANCE:
What will practitioners, policymakers, or service providers DO differently based on these findings?
Name specific organisations, policy areas, or professional communities. Be concrete about the practical changes warranted.
Example: "These findings have direct implications for primary healthcare policy in East Africa. If the results confirm our hypothesis, the WHO will need to revise its guidelines for community health worker training, potentially affecting 450,000 workers across the region."

PARAGRAPH 4 — BENEFICIARY GROUPS:
Who specifically benefits and how? Name academics (which subdisciplines), practitioners (which professions),
policymakers (which agencies), communities (which populations). Make clear the pathway from evidence to impact.
Example: "The primary beneficiaries are (1) development researchers studying adaptive capacity in agriculture, (2) extension agents in smallholder farming who need evidence on which interventions work, (3) agricultural policymakers designing rural development programmes, and (4) smallholder farmers themselves, who..."

### 1.7 Scope and Delimitations
Write at least {w(100, 196)} words for this subsection.
Define the geographic, temporal, and thematic boundaries with precision. For each
boundary, explain not just what is excluded but why the exclusion is methodologically
justified rather than a limitation of convenience. Acknowledge the trade-offs involved.

### 1.8 Limitations of the Study
Write at least {w(100, 196)} words for this subsection.
Identify at least four genuine constraints — methodological, practical, or contextual.
For each, explain what the limitation is, how it arose, and what steps were taken to
minimise its impact on the validity and transferability of findings. Be candid: real
researchers acknowledge imperfection.

### 1.9 Definition of Key Terms
Write at least {w(110, 210)} words for this subsection.
Define 6–8 terms that carry specific technical or conceptual meanings in this study.
For each term: provide a working definition grounded in at least one cited scholar,
explain how this study's usage compares to or departs from common usage, and note any
definitional controversies relevant to the research.

### 1.10 Organisation of the Study
Write at least {w(65, 126)} words for this subsection.
Describe what each chapter covers in two to three sentences per chapter — not a list,
but short, flowing paragraphs. Explain the logical progression from chapter to chapter.

Do NOT write a chapter title heading at the very top — begin directly with section ### 1.1.""",

        2: f"""You are writing CHAPTER TWO — LITERATURE REVIEW for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[2]} words of substantive prose. You MUST reach this minimum.
The literature review is the longest and most intellectually demanding chapter. Write with depth.
{_NO_REF}
{tone}

{HUMAN_WRITING_INSTRUCTION}
{_FN_NOTE}
{_NO_AST}
{_VIZ_NOTE}

Write the following subsections in full. Every subsection demands extended, analytical prose.

### 2.1 Introduction to the Chapter
Write at least {w(100, 210)} words.
Open by situating the literature review within the study's broader purpose. Explain how
this chapter is organised and why that organisational logic was chosen. Describe the scope
of literature reviewed — databases, date range, inclusion criteria — without being mechanical.
End with a statement of what the review reveals and how it sets up the research gap.

### 2.2 Conceptual Review
Write at least {w(210, 420)} words.
Identify the 4–6 central concepts of this study. For each concept: trace its intellectual
history (who coined or defined it, when, and in what context), map the range of definitions
across the literature (noting where scholars converge and diverge), and state explicitly
which conceptualisation this study adopts and why. Write this as connected analytical prose,
not as a series of dictionary definitions.
{"Engage with conceptual tensions and competing paradigms — do not smooth them over." if is_pg else ""}

### 2.3 Theoretical Framework
Write at least {w(240, 490)} words.
Identify 2–3 theories or models that directly inform this study. For each theory, develop
a full sub-argument across multiple paragraphs: name the originator and intellectual
tradition, describe the core propositions, trace how it has been applied and tested in
empirical research over the past decade, and make explicit how it will guide this study's
analytical framework. {"Critically evaluate each theory — identify its explanatory strengths, its known limitations, and how scholars have critiqued or refined it." if is_pg else "Explain how each theory applies to the specific context of this study."}

### 2.4 Empirical Review
Write at least {w(325, 630)} words.
Critically review at least {"10-12" if is_pg else "6-8"} prior studies. Organise the
review thematically rather than chronologically. For each thematic cluster: identify the
key studies, summarise their findings and methodological approaches, note where results
converge, flag contradictions or anomalies in the evidence base, and comment on methodological
quality. {"Evaluate sample sizes, research designs, and contextual applicability." if is_pg else ""}
This section must read as a genuine scholarly conversation, not a descriptive catalogue.

### 2.5 Review of Related Studies
Write at least {w(210, 420)} words.
Focus specifically on studies conducted in comparable contexts or addressing analogous
sub-questions. For each study reviewed: explain what it investigated, summarise its
principal findings, assess what it contributes to this study's conceptual or empirical
foundations, and — critically — identify precisely where it falls short relative to the
present study's aims. This section should make the research gap feel inevitable.

### 2.6 Research Gap
Write at least {w(140, 280)} words.
Do not simply assert that a gap exists — argue for it. Draw together the evidence from the
preceding sections to show exactly what has been studied, what remains unstudied, why the
existing studies are insufficient for this particular problem, and why this gap matters.
{"Distinguish between empirical gaps (what data are missing), theoretical gaps (what explanatory frameworks have not been tested here), and methodological gaps (how prior studies' designs could be improved)." if is_pg else "Make clear why filling this gap produces knowledge that is both novel and useful."}

### 2.7 Chapter Summary
Write at least {w(125, 245)} words.
Do not list what was covered. Instead, synthesise: identify the 2–3 most important
intellectual threads that emerge from the review, explain how they relate to each other,
and show explicitly how they set up the methodological choices and analytical framework
of Chapter 3. End with a sentence or two that creates a bridge forward.

Do NOT write a chapter title heading at the very top — begin directly with section ### 2.1.""",

        3: f"""You are writing CHAPTER THREE — RESEARCH METHODOLOGY for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[3]} words of substantive prose. You MUST reach this minimum.
The methodology chapter must be precise, justified, and replicable. Write with rigour.
VISUALIZATIONS ARE MANDATORY — include 6-8 figures/tables throughout this chapter.
{_NO_REF}
{tone}

{HUMAN_WRITING_INSTRUCTION}
{_FN_NOTE}
{_NO_AST}

FIGURE AND TABLE STANDARDS FOR CHAPTER 3:
Use EXACTLY these formats — ALL visualizations must be tables or charts:

REQUIRED VISUALIZATIONS (use [TABLE:...] or [CHART:...]):
1. [TABLE: Research Design Framework - shows Paradigm | Design Type | Connection to RQs | Methodological Justification]
2. [TABLE: Analytical Process Flowchart - shows Stage | Activity | Input Data | Output | Next Stage]
3. [TABLE: Sampling Strategy Breakdown with Target Population | Total N | Inclusion Criteria | Sample Size (n) | Sampling Method | Justification]
4. [TABLE: Data Collection Timeline with Week/Phase | Activities | Responsible Party | Expected Outputs | Duration]
5. [TABLE: Data Collection Instruments Matrix with Instrument Name | Type | Purpose | Number of Items | Reliability Method | Sample Items]
6. [TABLE: Validity and Reliability Framework - shows Quality Measure | Definition | How Achieved | Evidence]
7. [TABLE: Ethical Considerations Checklist with Ethical Dimension | Consideration | How Operationalised | Approval Status]
8. [CHART: Methodology Integration Diagram - visual showing Research Question → Design → Sample → Instruments → Analysis]

Format: All tables use pipe-separated columns (| header | header |). Provide actual data rows, not just headers.
Captions: After each [TABLE:...] or [CHART:...] marker, include a 1-2 sentence caption explaining purpose and key insight.
References: Cite all tables/figures in text BEFORE they appear (e.g., 'As presented in Table 3.1, the sampling strategy...')
All visualizations will be converted to professional Word tables with borders, headers, and alternating row colors.

Write the following subsections in full.

### 3.1 Introduction to the Chapter
Write at least {w(90, 175)} words.
Orient the reader to the chapter's purpose and structure. Explain the epistemological logic
that connects the research questions to the design choices made. {"State the researcher's ontological and epistemological position upfront and explain how it shapes the chapter's approach to the treatment of evidence and knowledge claims." if is_pg else "Explain how the methodology flows from the research questions and problem."}
After this section, include:
[FIGURE: Research Design Framework showing the paradigm → design → connection to research questions]

### 3.2 Research Design
Write at least {w(160, 315)} words.
Describe the overall research strategy and justify the choice of qualitative, quantitative,
or mixed-methods design by reference to the nature of the research questions. Cite at least
three methodologists who support this design choice. Explain what this design can and cannot
do — including what it sacrifices — and defend the choice against obvious alternatives.
{"Connect the design explicitly to the epistemological position stated in 3.1." if is_pg else ""}

### 3.3 Research Philosophy and Paradigm
Write at least {w(160, 350)} words.
{"Develop the philosophical grounding in detail. Discuss the ontological position (what the researcher believes about the nature of reality — is it singular and knowable, or multiple and constructed?), the epistemological position (what counts as valid knowledge, and how it can be acquired), and how these positions connect to the chosen methodology. Distinguish between positivism, interpretivism, constructivism, pragmatism, and critical realism with enough precision that the reader understands which stance is adopted here and why." if is_pg else "Identify the research paradigm (e.g., interpretivist, positivist, pragmatist) and explain in clear terms how it shapes the study's approach to data, evidence, and knowledge. Draw on at least two methodologists to justify the paradigmatic choice."}

### 3.4 Research Approach
Write at least {w(100, 196)} words.
{"Specify whether the study uses inductive, deductive, or abductive reasoning. Justify this choice by reference to the research questions and the nature of the evidence being collected. Explain how the approach shapes the analytical process in Chapter 4." if is_pg else "Specify the reasoning approach (inductive or deductive) and explain how it guides data analysis. Connect this to the research design."}
After this section, include:
[FIGURE: Analytical Process Flowchart showing data collection → initial coding → analytical refinement → interpretation steps]

### 3.5 Study Area and Setting
Write at least {w(115, 224)} words.
Describe the physical, institutional, or organisational setting with enough specificity that
the reader can visualise it. Explain why this setting was chosen — what makes it appropriate
for answering these research questions. Discuss access, gatekeeping, and any contextual
factors (political, cultural, institutional) that shaped the fieldwork.

### 3.6 Target Population
Write at least {w(100, 196)} words.
Define the population with precision — who qualifies, why they qualify, and how large the
total population is (with a source if applicable). Explain the relevance of this population
to the research questions. Address any challenges in defining or accessing the population.

### 3.7 Sample Size and Sampling Technique
Write at least {w(135, 266)} words.
Specify the sample size and justify it — cite at least two sources on sample size adequacy
for the chosen design. Describe the sampling technique in precise operational terms: exactly
how participants were identified, approached, screened, and recruited. {"Discuss how the technique addresses issues of representativeness (quantitative) or theoretical saturation and transferability (qualitative)." if is_pg else "Explain how the sample is representative of the population."}
Address any non-response and how it was handled.
After this section, include:
[TABLE: Sampling Strategy Breakdown]
Target Population | Total N | Inclusion Criteria | Sample Size (n) | Sampling Method | Justification
[Provide specific numbers and explanation for how sample was derived from target population]

### 3.8 Data Collection Instruments
Write at least {w(140, 280)} words.
Describe each instrument used (questionnaire, semi-structured interview guide, observation
protocol, document analysis schedule). For each instrument: explain the rationale for its
design, describe its structure (sections, item types, scale formats), explain the piloting
process and any revisions made, and justify its appropriateness for collecting the data
required by each research objective.
After this section, include:
[TABLE: Data Collection Instruments Matrix]
Instrument Name | Purpose (Which RQ?) | Structure (Sections/Items) | Response Format | Justification

### 3.9 Validity and Reliability
Write at least {w(135, 266)} words.
{"Address validity and reliability using the criteria appropriate to the paradigm. For quantitative work: construct validity, criterion validity, internal consistency (Cronbach's alpha), and test-retest reliability. For qualitative work: credibility (member-checking, triangulation), transferability (thick description), dependability (audit trail), and confirmability (reflexivity) — drawing on Lincoln and Guba (1985). Explain specifically how each criterion was operationalised in this study." if is_pg else "Explain what steps were taken to ensure the instruments measure what they intend to measure and produce consistent results. Discuss any piloting and revision process. Address both internal validity and reliability."}
After this section, include:
[FIGURE: Validity and Reliability Framework]
Showing paradigm-appropriate quality measures and how each was operationalised in this study.

### 3.10 Data Collection Procedure
Write at least {w(125, 245)} words.
Describe the data collection process as a step-by-step chronological narrative: ethics
clearance, participant recruitment, informed consent, instrument administration, data
recording, and quality checks. Include time frames and quantities (how many interviews
conducted over how many weeks, response rate for questionnaires). Be specific enough
that a researcher could replicate this procedure.
After this section, include:
[TABLE: Data Collection Timeline - 7-week schedule]
Week | Primary Activities | Responsible Party | Expected Outputs

### 3.11 Data Analysis Methods
Write at least {w(140, 280)} words.
Explain the analytical approach in enough detail for replication. {"Name the specific software used (SPSS, NVivo, Atlas.ti, R, Python) and justify the choice. Describe the analytical procedures step by step: coding (open, axial, selective), thematic analysis phases, statistical tests applied and their assumptions, regression models and their specification. Connect each analytical step to the specific research questions it addresses." if is_pg else "Describe the analytical approach clearly: how data were organised, coded, and interpreted. Name any software used and explain how it was applied. Connect the analysis to the research questions."}

### 3.12 Ethical Considerations
Write at least {w(110, 210)} words.
Address at least six ethical dimensions: informed consent (what participants were told and
how consent was obtained), anonymity and confidentiality (how data were anonymised and
protected), right to withdraw (how this was communicated and facilitated), data storage
and security (how data are stored and for how long), institutional ethics approval
(institution and reference number if applicable), and researcher positionality
(how the researcher's background may have influenced data collection and interpretation).
After this section, include:
[TABLE: Ethical Considerations Operationalisation]
Ethical Dimension | Description | How Operationalised in This Study | Evidence

### 3.13 Chapter Summary
Write at least {w(100, 196)} words.
Synthesise the methodological choices made in this chapter as a coherent whole. Explain
how design, philosophy, sampling, instruments, and analysis hang together as a unified
approach to answering the research questions. {"Address how the methodology addresses the research gap identified in Chapter 2 and positions the study within its paradigmatic tradition." if is_pg else "Show how the methodology directly serves the research objectives stated in Chapter 1."}
After this section, include:
[FIGURE: Methodology Integration Diagram]
Showing how research design, philosophy, sampling strategy, data collection instruments, and analytical approach connect as a coherent system.

Do NOT write a chapter title heading at the very top — begin directly with section ### 3.1.""",

        4: f"""You are writing CHAPTER FOUR — RESULTS AND DISCUSSION for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[4]} words of substantive prose. You MUST reach this minimum.
Present rich, specific, interpreted findings. This chapter must demonstrate analytical depth.
VISUALIZATIONS ARE CRITICAL — include 8-12 figures/tables throughout this chapter to present data professionally.
{_NO_REF}
{tone}

{HUMAN_WRITING_INSTRUCTION}
{_FN_NOTE}
{_NO_AST}

FIGURE AND TABLE STANDARDS FOR CHAPTER 4:
This chapter requires extensive data visualization to meet PhD standards:
- Table 4.1: Sample Demographics (Age | Gender | Education | Experience | Location with frequencies/%)
- Table 4.2: Response Rate and Non-Response Analysis (if quantitative)
- Figures 4.1-4.3: Data visualizations for Objective 1 (charts, graphs, or qualitative matrices)
- Figures 4.4-4.6: Data visualizations for Objectives 2-3 (charts, graphs, comparative tables)
- Figure 4.7: Synthesis diagram or matrix showing relationships across all objectives
- Table 4.3: Finding Summary Matrix (Objective | Key Finding | Supporting Evidence | Theoretical Connection)
- Figures 4.8-4.10: Additional visualizations for complex patterns, comparisons, or theoretical connections
- Figure 4.11: Implications framework (theoretical, practical, policy dimensions)

VISUALIZATION REQUIREMENTS:
- For quantitative findings: use bar charts, line graphs, distribution plots for each major result
- For qualitative findings: use thematic matrices, concept maps, or comparison tables
- For mixed-methods: use integrated visualizations showing convergence/divergence
- All figures must have numbered captions (Figure 4.1, Figure 4.2, etc.) with descriptive titles
- All tables must be professionally formatted with clear headers and logical grouping
- Every table/figure MUST be referenced in the text before it appears ("As shown in Table 4.1...")

Write the following subsections in full.

### 4.1 Introduction to the Chapter
Write at least {w(90, 175)} words.
Explain how the chapter is structured and why. Briefly recap the research objectives so
the reader knows what findings will address. {intro_text}

### 4.2 Sample / Response Rate Overview
Write at least {w(110, 210)} words.
Present the demographic and descriptive profile of the sample across multiple characteristics
(age, gender, education, experience, geographic distribution — as relevant). Discuss the
response rate if applicable and explain patterns in non-response. {sample_text}
Include immediately after this section:
[TABLE: Sample Demographics Breakdown]
Demographic Variable | Categories | Frequency (n) | Percentage (%)
[Include all relevant characteristics with actual numbers]

For quantitative studies, also include:
[CHART: Sample demographics visualization]
Showing age distribution, gender distribution, or other key characteristics visually.

### 4.3 Findings Related to Objective 1
Write at least {w(180, 350)} words.
Present specific, detailed findings for the first research objective. Use plausible
quantitative values (percentages, means, frequencies) or qualitative themes with
representative illustrative evidence. Interpret the findings rather than just reporting
them: explain what patterns emerge, what they mean, and what accounts for them.
{obj1_text}
Include professional data visualization immediately after:
[CHART: Bar chart/line graph/distribution showing Objective 1 findings with specific values]
OR
[TABLE: Objective 1 Thematic Analysis Matrix | Theme | Frequency | Representative Quote/Evidence | Theoretical Connection]

### 4.4 Findings Related to Objective 2
Write at least {w(180, 350)} words.
Apply the same approach as 4.3 to the second research objective. Ensure this section has
its own narrative arc — do not simply replicate the structure of 4.3. Introduce any
unexpected or contradictory findings and engage with them analytically.
Include data visualization:
[CHART: Bar chart/line graph showing Objective 2 findings]
OR
[TABLE: Objective 2 Findings Summary | Finding | Evidence | Significance]

### 4.5 Findings Related to Objective 3
Write at least {w(180, 350)} words.
Apply the same approach to the third objective. {obj3_text}
Include data visualization:
[CHART: Chart showing Objective 3 findings]
OR
[TABLE: Objective 3 Findings Detailed Breakdown]

### 4.6 Findings Related to Objective 4
Write at least {w(160, 315)} words.
Present findings for the fourth objective with the same analytical rigour. By the end of
this section, all major findings should be on the table, setting up the synthesis in 4.7.
Include data visualization:
[CHART: Chart comparing Objective 4 findings]
OR
[TABLE: Objective 4 Key Findings with Evidence]

### 4.7 Synthesis and Discussion of Major Findings
Write at least {w(210, 420)} words.
THIS IS THE INTELLECTUAL HEART OF THE CHAPTER — your opportunity to demonstrate meta-analytical thinking across all four objectives.

{synthesis_text}

### 4.8 Implications of the Findings
Write at least {w(140, 280)} words.
Discuss implications for theory, practice, and policy separately across dedicated paragraphs.
Name specific stakeholders and explain precisely what each set of findings means for them.
{implications_text}
Include immediately after:
[TABLE: Implications Framework]
Implication Domain | Specific Implication | Target Stakeholder | Actionable Consequence

### 4.9 Chapter Summary
Write at least {w(100, 196)} words.
Distil the most important results and analytical insights in two to three substantive
paragraphs. Do not list findings — synthesise. End with a transition that sets up the
conclusions and recommendations in Chapter 5.

Do NOT write a chapter title heading at the very top — begin directly with section ### 4.1.""",

        5: f"""You are writing CHAPTER FIVE — CONCLUSIONS AND RECOMMENDATIONS for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[5]} words of substantive prose. You MUST reach this minimum.
This chapter must deliver a satisfying intellectual conclusion — not a mechanical recap.

{tone}

{HUMAN_WRITING_INSTRUCTION}
{_FN_NOTE}
{_NO_AST}
{_VIZ_NOTE}

Write the following subsections in full.

### 5.1 Introduction to the Chapter
Write at least {w(75, 140)} words.
Orient the reader to the chapter's purpose and structure. Briefly explain how this chapter
brings the entire study to a close and what it aims to deliver beyond simply summarising
earlier chapters.

### 5.2 Summary of the Study
Write at least {w(160, 315)} words.
Recount the entire research journey in a flowing, synthesised narrative across at least
four substantive paragraphs: the problem and its context, the objectives and theoretical
framework, the methodology and its justification, and the principal findings. Do not
quote verbatim from earlier chapters — reframe and integrate. A reader encountering this
study for the first time through this section should understand its full arc.

### 5.3 Conclusions
Write at least {w(180, 350)} words.
Draw one specific, argued conclusion per research objective — each conclusion in its own
paragraph. Each conclusion must: state what the study found, explain what this finding
means in context, and connect it to the evidence from Chapter 4. {"Where conclusions are tentative or conditional, say so and explain the conditions under which the conclusion holds. Where they challenge prior theory, develop that challenge explicitly." if is_pg else "State conclusions with appropriate confidence — neither overclaiming nor underselling what the data support."}

### 5.4 Contribution to Knowledge
Write at least {w(135, 266)} words.
{"Articulate the study's contribution across at least three dimensions: theoretical (how it extends, refines, or challenges existing theoretical models), empirical (what new data or patterns it adds to the evidence base), and methodological (whether it demonstrates a novel application of method in this context). Be precise — 'this study contributes to the literature' is not a contribution; naming exactly what it adds is." if is_pg else "Explain in concrete terms what is new or valuable about what this study found. How does it advance understanding beyond what was known before? What practical problems does it help solve?"}

### 5.5 Recommendations
Write at least {w(160, 315)} words.
Provide 5–6 specific, actionable, evidence-grounded recommendations. Write each as a
full paragraph rather than a bullet point: name the recommendation, identify the specific
finding that supports it, name the stakeholder or institution it is directed at, and
describe what implementing it would look like in practice. Recommendations must flow
directly from the findings — no recommendation should appear without a grounding in
Chapter 4.

### 5.6 Recommendations for Future Research
Write at least {w(125, 245)} words.
Propose 3–4 specific research directions that arise from this study's limitations or from
questions it raised but could not answer. Each recommendation for future research should:
identify the gap or question, explain why it matters, suggest an appropriate methodological
approach, and state what such research would contribute. {"For postgraduate work, these should point toward theoretical refinement, comparative cross-context studies, or longitudinal designs." if is_pg else ""}

### 5.7 Chapter Summary
Write at least {w(80, 154)} words.
A dignified, forward-looking closing that does not merely repeat the conclusions. Reflect
on what the study set out to do and what it achieved. End with a final paragraph that
gestures toward the broader significance of the work — without overreaching.

---

After 5.7, write the following two sections:

## REFERENCES
List at least {w(9, 18)} academic references in APA 7th edition format.
References must be plausible, field-relevant, diverse, and correctly formatted.
Include: journal articles (majority), books, book chapters, institutional/government
reports, and conference papers. Span at least 2005–2023. Mix foundational texts with
recent scholarship (at least 6 references from 2018 onwards).
Format each exactly as:
  Author, A. A., & Author, B. B. (Year). Title of article. Journal Name, Volume(Issue), pages. https://doi.org/xxxxx

## APPENDICES

### Appendix A: Research Instrument
Provide a complete {"interview guide (14+ open and semi-structured questions across thematic sections)" if is_pg else "questionnaire (10+ items using Likert scales, multiple choice, and open-ended questions)"}, appropriate to the research design described in Chapter 3. Include an introduction/preamble and section headings.

### Appendix B: Data Collection Timeline
A structured {"7" if is_pg else "4"}-week timeline table for the data collection phase, with activities, responsible parties, and expected outputs for each week.

### Appendix C: Ethical Clearance Template
A sample informed consent form that would be used with participants in this study,
including all required elements (study description, risks, rights, confidentiality, contact details).

Do NOT write a chapter title heading at the very top — begin directly with section ### 5.1.""",
    }


# ─────────────────────────────────────────────────────────
#  WORD FOOTNOTE SUPPORT
# ─────────────────────────────────────────────────────────

_FOOTNOTES_REL  = ('http://schemas.openxmlformats.org/officeDocument/'
                   '2006/relationships/footnotes')
_FOOTNOTES_CT   = ('application/vnd.openxmlformats-officedocument'
                   '.wordprocessingml.footnotes+xml')
_FN_INLINE_RE   = re.compile(r'\(\(FN:\s*(.+?)\)\)', re.DOTALL)


class FootnoteManager:
    """
    Manages Word footnotes using a post-save zip-injection strategy.

    Strategy:
        1. add_footnote() writes the <w:footnoteReference> superscript into the
           body paragraph (standard OxmlElement — no internal API required).
        2. The footnote text is queued in self._footnotes.
        3. After doc.save(path), call fn_mgr.inject(path) to inject
           word/footnotes.xml plus the required relationships/content-types
           directly into the .docx zip file.

    Result: clicking the superscript in Word jumps to the footnote at the
    bottom of the page, and clicking the footnote number jumps back.
    """

    def __init__(self, doc):
        self.doc        = doc
        self._next_id   = 1
        self._footnotes = []   # list of (id, text) pairs

    def add_footnote(self, paragraph, footnote_text: str) -> int:
        """
        Insert a superscript footnote reference into *paragraph*.
        Returns the footnote id.  Call inject() after doc.save().
        """
        fn_id = self._next_id
        self._next_id += 1

        # Superscript reference in body text
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rStyle = OxmlElement('w:rStyle')
        rStyle.set(qn('w:val'), 'FootnoteReference')
        rPr.append(rStyle)
        r.append(rPr)
        ref = OxmlElement('w:footnoteReference')
        ref.set(qn('w:id'), str(fn_id))
        r.append(ref)
        paragraph._p.append(r)

        self._footnotes.append((fn_id, footnote_text.strip()))
        return fn_id

    def inject(self, docx_path: str):
        """
        Post-process the saved .docx zip to insert word/footnotes.xml
        and the required relationship + content-type entries.
        Must be called AFTER doc.save(docx_path).
        """
        if not self._footnotes:
            return

        import zipfile
        from lxml import etree

        W       = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        XML_NS  = 'http://www.w3.org/XML/1998/namespace'
        FN_REL  = ('http://schemas.openxmlformats.org/officeDocument/'
                   '2006/relationships/footnotes')
        FN_CT   = ('application/vnd.openxmlformats-officedocument'
                   '.wordprocessingml.footnotes+xml')

        # ── Build footnotes.xml ──────────────────────────
        root = etree.Element(f'{{{W}}}footnotes', nsmap={'w': W})

        for sep_id, sep_tag in [('-1', 'separator'), ('0', 'continuationSeparator')]:
            fn = etree.SubElement(root, f'{{{W}}}footnote')
            fn.set(f'{{{W}}}type', sep_tag)
            fn.set(f'{{{W}}}id', sep_id)
            p = etree.SubElement(fn, f'{{{W}}}p')
            r = etree.SubElement(p, f'{{{W}}}r')
            etree.SubElement(r, f'{{{W}}}{sep_tag}')

        for fn_id, fn_text in self._footnotes:
            fn = etree.SubElement(root, f'{{{W}}}footnote')
            fn.set(f'{{{W}}}id', str(fn_id))
            p = etree.SubElement(fn, f'{{{W}}}p')

            r_num = etree.SubElement(p, f'{{{W}}}r')
            rPr   = etree.SubElement(r_num, f'{{{W}}}rPr')
            rs    = etree.SubElement(rPr, f'{{{W}}}rStyle')
            rs.set(f'{{{W}}}val', 'FootnoteReference')
            etree.SubElement(r_num, f'{{{W}}}footnoteRef')

            r_txt = etree.SubElement(p, f'{{{W}}}r')
            t     = etree.SubElement(r_txt, f'{{{W}}}t')
            t.set(f'{{{XML_NS}}}space', 'preserve')
            t.text = ' ' + fn_text

        footnotes_xml = etree.tostring(
            root, xml_declaration=True, encoding='UTF-8', standalone=True
        )

        # ── Rewrite the .docx zip ────────────────────────
        tmp = docx_path + '.fn.tmp'
        try:
            with zipfile.ZipFile(docx_path, 'r') as zin, \
                 zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:

                has_footnotes_xml = False
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename == 'word/_rels/document.xml.rels':
                        rel_entry = (
                            f'<Relationship Id="rFnotes1" '
                            f'Type="{FN_REL}" Target="footnotes.xml"/>'
                        ).encode()
                        data = data.replace(b'</Relationships>',
                                            rel_entry + b'</Relationships>')

                    elif item.filename == '[Content_Types].xml':
                        ct_entry = (
                            f'<Override PartName="/word/footnotes.xml" '
                            f'ContentType="{FN_CT}"/>'
                        ).encode()
                        data = data.replace(b'</Types>', ct_entry + b'</Types>')

                    elif item.filename == 'word/footnotes.xml':
                        has_footnotes_xml = True
                        data = footnotes_xml   # replace existing

                    zout.writestr(item, data)

                if not has_footnotes_xml:
                    zout.writestr('word/footnotes.xml', footnotes_xml)

            os.replace(tmp, docx_path)

        except Exception:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise


# ─────────────────────────────────────────────────────────
#  REFERENCES PAGE BUILDER
# ─────────────────────────────────────────────────────────

def build_references_page(doc, references_text: str):
    """Render the ## REFERENCES block on its own page."""
    add_page_break(doc)
    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run("REFERENCES")
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    hdr.paragraph_format.space_after = Pt(4)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Strip leading ## REFERENCES heading if Claude included it
    body = re.sub(r'(?im)^##\s*REFERENCES\s*$', '', references_text).strip()

    # Each non-empty line is one reference entry
    for line in body.splitlines():
        line = line.strip()
        if not line:
            continue
        p = doc.add_paragraph()
        p.add_run(line)
        p.paragraph_format.left_indent   = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)   # hanging indent
        p.paragraph_format.space_after   = Pt(4)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


# ─────────────────────────────────────────────────────────
#  DOCUMENT STYLING HELPERS
# ─────────────────────────────────────────────────────────

def add_page_break(doc):
    from docx.enum.text import WD_BREAK
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)


def add_horizontal_rule(doc, color="2E74B5", thickness="8"):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    thickness)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)


def add_chapter_header(doc, chapter_num, custom_subtitle=None):
    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    label = doc.add_paragraph()
    label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = label.add_run(CHAPTER_TITLES[chapter_num])
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    label.paragraph_format.space_after = Pt(4)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Use custom subtitle if provided, otherwise use default
    subtitle_text = custom_subtitle if custom_subtitle else CHAPTER_SUBTITLES[chapter_num]
    r2 = subtitle.add_run(subtitle_text)
    r2.font.size  = Pt(16)
    r2.font.bold  = True
    r2.font.color.rgb = MID_BLUE
    subtitle.paragraph_format.space_before = Pt(4)
    subtitle.paragraph_format.space_after  = Pt(20)


def _style_section_heading(paragraph, level):
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    if level == 2:
        run.font.size    = Pt(12)
        run.font.bold    = True
        run.font.color.rgb = MID_BLUE
        paragraph.paragraph_format.space_before = Pt(14)
        paragraph.paragraph_format.space_after  = Pt(5)
    elif level == 3:
        run.font.size    = Pt(11)
        run.font.bold    = True
        run.font.italic  = True
        run.font.color.rgb = GREY
        paragraph.paragraph_format.space_before = Pt(10)
        paragraph.paragraph_format.space_after  = Pt(4)


def _strip_stray_asterisks(text: str) -> str:
    """Remove any remaining lone asterisks that aren't part of bold/italic markup."""
    # Remove *** bold-italic markers (convert to plain text)
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)
    # After bold/italic has been parsed, strip any leftover bare asterisks
    # (ones not forming complete **…** or *…* pairs)
    text = re.sub(r'\*+', '', text)
    return text


def _add_inline_formatting(paragraph, text, fn_mgr=None):
    """Add text to paragraph with bold/italic and optional footnote support."""
    # First split on footnote markers if a manager is supplied
    if fn_mgr and _FN_INLINE_RE.search(text):
        segments = _FN_INLINE_RE.split(text)
        # split with one capture group gives [text, fn, text, fn, ...]
        for idx, seg in enumerate(segments):
            if idx % 2 == 0:
                # Regular text — apply bold/italic
                _add_inline_formatting(paragraph, seg, fn_mgr=None)
            else:
                # Footnote text — insert real Word footnote
                fn_mgr.add_footnote(paragraph, seg)
        return

    # Bold/italic processing — then strip any leftover asterisks
    parts = re.split(r"(\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*)", text)
    for part in parts:
        if part.startswith("***") and part.endswith("***"):
            r = paragraph.add_run(part[3:-3])
            r.bold = True
            r.italic = True
        elif part.startswith("**") and part.endswith("**"):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*"):
            paragraph.add_run(part[1:-1]).italic = True
        else:
            # Strip any residual bare asterisks from non-markup fragments
            paragraph.add_run(part.replace("*", ""))


def _render_table(doc, table_lines):
    """
    Render a list of markdown pipe-table lines as a professional Word Table.

    Handles:
      | Col A | Col B |        ← header row
      |---|---|                ← separator row (skipped)
      | data  | data  |        ← body rows

    Features:
      - Proper column width distribution (6.5 inches available)
      - Professional borders and shading
      - Text wrapping for long content
      - Alternating row colors for readability
      - Centered headers with bold formatting
      - Proper cell padding and margins
    """
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # Parse each row into a list of cell strings
    def parse_row(line):
        line = line.strip().strip("|")
        return [cell.strip() for cell in line.split("|")]

    # Filter out pure separator rows (only dashes/colons/pipes/spaces)
    rows = []
    for line in table_lines:
        if re.match(r'^[\s|:\-]+$', line):
            continue           # skip separator rows
        cells = parse_row(line)
        if any(cells):
            rows.append(cells)

    if not rows:
        return

    # Normalise column count
    col_count = max(len(r) for r in rows)
    rows = [r + [''] * (col_count - len(r)) for r in rows]

    # Create table
    tbl = doc.add_table(rows=len(rows), cols=col_count)
    tbl.autofit = False
    try:
        tbl.style = 'Table Grid'
    except Exception:
        pass

    # Calculate optimal column widths (6.5 inches available page width)
    total_width = Inches(6.5)
    col_width = Inches(6.5 / col_count) if col_count > 0 else Inches(1.0)

    # Apply consistent column widths
    for col_idx in range(col_count):
        for row in tbl.rows:
            row.cells[col_idx].width = col_width

    # Header and body row styling
    header_color = 'D3D3D3'      # Light grey for headers
    alt_row_color = 'F5F5F5'     # Very light grey for alternating rows

    for r_idx, row_data in enumerate(rows):
        row = tbl.rows[r_idx]
        is_header = (r_idx == 0)

        # Set row height
        row.height = Inches(0.35) if is_header else Inches(0.45)

        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]

            # Set cell background color
            shading_elm = OxmlElement('w:shd')
            if is_header:
                shading_elm.set(qn('w:fill'), header_color)
            elif r_idx % 2 == 0:  # Alternate row colors
                shading_elm.set(qn('w:fill'), alt_row_color)
            else:
                shading_elm.set(qn('w:fill'), 'FFFFFF')  # White for other rows
            cell._element.get_or_add_tcPr().append(shading_elm)

            # Set cell borders
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '12')  # Border width
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), 'CCCCCC')  # Light grey border
                tcBorders.append(border)
            tcPr.append(tcBorders)

            # Set vertical alignment (top for better readability)
            cell.vertical_alignment = 0  # 0 = top

            # Configure paragraph and text
            for p in cell.paragraphs:
                p.clear()
            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()

            # Alignment: center for headers, left for body
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.LEFT

            # Text formatting
            run = p.add_run(cell_text)
            if is_header:
                run.bold = True
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 0, 0)  # Black text
            else:
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(32, 32, 32)  # Dark grey text

            # Paragraph formatting (padding and margins)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.left_indent = Pt(8)
            p.paragraph_format.right_indent = Pt(8)
            p.paragraph_format.line_spacing = 1.1

            # Enable text wrapping
            cell.width = col_width

    # Add spacing after table
    space_para = doc.add_paragraph()
    space_para.paragraph_format.space_after = Pt(8)
    space_para.paragraph_format.space_before = Pt(4)


def _create_sample_chart(description: str):
    """
    Create a sample chart image based on description.
    Returns path to saved PNG image.

    Strategy:
    1. Try matplotlib (best quality)
    2. Fallback to PIL/Pillow (simple charts)
    3. Fallback to text-based placeholder
    """

    # ─── TRY MATPLOTLIB FIRST ───────────────────────────────────
    if VISUALIZATION_AVAILABLE:
        try:
            fig, ax = plt.subplots(figsize=(8, 5), dpi=100)

            # Generate sample data based on description keywords
            if 'line' in description.lower() or 'trend' in description.lower():
                x = np.arange(1, 6)
                y = np.array([20, 35, 48, 62, 78])
                ax.plot(x, y, marker='o', linewidth=2, markersize=8, color='#2E74B5')
                ax.set_xlabel('Period', fontsize=11, fontweight='bold')
                ax.set_ylabel('Value', fontsize=11, fontweight='bold')
                ax.grid(True, alpha=0.3)
            elif 'bar' in description.lower():
                categories = ['Category A', 'Category B', 'Category C', 'Category D']
                values = [45, 62, 38, 71]
                ax.bar(categories, values, color='#2E74B5', alpha=0.8, edgecolor='black')
                ax.set_ylabel('Value', fontsize=11, fontweight='bold')
                plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            elif 'pie' in description.lower():
                labels = ['Group A', 'Group B', 'Group C', 'Group D']
                sizes = [30, 25, 25, 20]
                colors = ['#2E74B5', '#4F90C3', '#A9C8E1', '#D9E5F0']
                ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
                ax.set_title('Distribution', fontweight='bold', fontsize=12)
            else:
                # Default: simple bar chart
                categories = ['Item 1', 'Item 2', 'Item 3']
                values = [55, 68, 42]
                ax.bar(categories, values, color='#2E74B5', alpha=0.8, edgecolor='black')
                ax.set_ylabel('Value', fontsize=11, fontweight='bold')

            ax.set_title(description, fontweight='bold', fontsize=13, pad=15)
            plt.tight_layout()

            # Save to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                fig.savefig(tmp.name, dpi=100, bbox_inches='tight')
                plt.close(fig)
                return tmp.name
        except Exception as e:
            plt.close('all')
            pass  # Fall through to PIL fallback

    # ─── FALLBACK: PIL/PILLOW ───────────────────────────────────
    try:
        from PIL import Image, ImageDraw, ImageFont

        # Create image
        width, height = 800, 500
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)

        # Draw border
        draw.rectangle([10, 10, width-10, height-10], outline='#2E74B5', width=3)

        # Draw title
        title_y = 30
        draw.text((20, title_y), description, fill='#1F497D')

        # Determine chart type and draw accordingly
        desc_lower = description.lower()

        if 'line' in desc_lower or 'trend' in desc_lower:
            # Draw line chart
            points = [(100, 350), (220, 280), (340, 200), (460, 150), (580, 100)]
            draw.line(points, fill='#2E74B5', width=3)
            for point in points:
                draw.ellipse([point[0]-5, point[1]-5, point[0]+5, point[1]+5], fill='#2E74B5')
        elif 'pie' in desc_lower:
            # Draw pie chart representation
            center_x, center_y = width // 2, height // 2 - 30
            radius = 80
            draw.ellipse([center_x-radius, center_y-radius, center_x+radius, center_y+radius],
                        outline='#2E74B5', width=2)
            # Draw pie segments
            angles = [0, 108, 198, 288, 360]
            colors = ['#2E74B5', '#4F90C3', '#A9C8E1', '#D9E5F0']
            for i in range(len(colors)):
                draw.pieslice([center_x-radius, center_y-radius, center_x+radius, center_y+radius],
                            angles[i], angles[i+1], fill=colors[i], outline='#1F497D')
        else:
            # Draw bar chart (default)
            bar_width = 60
            bars_x = [120, 240, 360, 480]
            bars_height = [280, 380, 220, 340]
            for x, h in zip(bars_x, bars_height):
                draw.rectangle([x, h, x+bar_width, 400], fill='#2E74B5', outline='#1F497D')

        # Save image
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            img.save(tmp.name, 'PNG')
            return tmp.name
    except Exception as e:
        pass  # Fall through to text placeholder

    # ─── FALLBACK: TEXT-BASED PLACEHOLDER ───────────────────────
    try:
        from PIL import Image, ImageDraw

        width, height = 800, 500
        img = Image.new('RGB', (width, height), color='#F5F5F5')
        draw = ImageDraw.Draw(img)

        # Draw border
        draw.rectangle([20, 20, width-20, height-20], outline='#CCCCCC', width=2)

        # Draw placeholder text
        text = f"[Chart: {description[:60]}...]"
        draw.text((60, height//2 - 30), text, fill='#999999')
        draw.text((60, height//2 + 20), "Chart generation requires matplotlib or PIL", fill='#CCCCCC')

        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            img.save(tmp.name, 'PNG')
            return tmp.name
    except Exception:
        pass

    # If all else fails, return None (but this is rare)
    return None


def parse_chapter_content(doc, content, fn_mgr=None):
    """
    Render markdown-ish chapter content into the Word document.
    fn_mgr — a FootnoteManager instance; if supplied, ((FN: text)) markers
              become real Word footnotes with page-bottom hyperlinks.
    """
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue

        # ── Table marker [TABLE: description] ────────────────
        if line.startswith("[TABLE:"):
            # Match [TABLE: ...] with optional content after closing bracket
            match = re.match(r"\[TABLE:\s*(.+?)\](.*)", line)
            if match:
                table_title = match.group(1).strip()
                table_content = match.group(2).strip()

                # If no content on same line, collect from next lines
                collected_lines = []
                if not table_content:
                    i += 1
                    while i < len(lines):
                        next_line = lines[i].rstrip()
                        if not next_line.strip():
                            i += 1
                            continue
                        if re.match(r"^#{1,3} ", next_line) or next_line.startswith("["):
                            break
                        if next_line.count("|") >= 3:  # Table-like line
                            collected_lines.append(next_line)
                            i += 1
                        else:
                            break

                # Parse table content if we have any
                table_lines = []

                if collected_lines:
                    # Process collected lines (each line is a row with | separators)
                    for line in collected_lines:
                        cells = [cell.strip() for cell in line.split("|")]
                        cells = [c for c in cells if c]  # Remove empty cells
                        if cells:
                            pipe_line = "| " + " | ".join(cells) + " |"
                            table_lines.append(pipe_line)

                    # Add separator after first row (header)
                    if len(table_lines) > 1 and collected_lines:
                        col_count = len([c for c in collected_lines[0].split("|") if c.strip()])
                        separator = "| " + " | ".join(["---"] * col_count) + " |"
                        table_lines.insert(1, separator)

                elif table_content and ' | ' in table_content:
                    # Inline content (all on one line with | separators)
                    all_cells = [cell.strip() for cell in table_content.split(' | ')]

                    if all_cells and len(all_cells) >= 4:
                        # For inline content, infer column count
                        # Prefer common table sizes (4, 5, 6 columns)
                        col_count = None
                        for preferred_cols in [4, 5, 6, 3, 7]:
                            if len(all_cells) % preferred_cols == 0:
                                col_count = preferred_cols
                                break

                        # Fallback
                        if col_count is None:
                            for test_cols in range(2, min(10, len(all_cells) // 2 + 1)):
                                if len(all_cells) % test_cols == 0:
                                    col_count = test_cols
                                    break

                        if col_count is None:
                            col_count = max(2, len(all_cells) // 3)

                        # Group cells into rows by column count
                        rows = []
                        for idx in range(0, len(all_cells), col_count):
                            row = all_cells[idx:idx + col_count]
                            if len(row) == col_count:
                                rows.append(row)

                        # Convert to pipe-table format
                        if rows:
                            for row in rows:
                                pipe_line = "| " + " | ".join(row) + " |"
                                table_lines.append(pipe_line)

                            # Add separator after header
                            if len(table_lines) > 1:
                                separator = "| " + " | ".join(["---"] * col_count) + " |"
                                table_lines.insert(1, separator)

                # Render the table if we have any lines
                if table_lines:
                    _render_table(doc, table_lines)

            i += 1
            continue

        # ── Chart marker [CHART: description] ─────────────────
        if line.startswith("[CHART:"):
            match = re.match(r"\[CHART:\s*(.+?)\]", line)
            if match:
                description = match.group(1).strip()
                chart_path = _create_sample_chart(description)
                if chart_path:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(chart_path, width=Inches(5.5))
                    cap = doc.add_paragraph()
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap_run = cap.add_run(description)
                    cap_run.italic = True
                    cap_run.font.size = Pt(10)
                    cap.paragraph_format.space_after = Pt(6)
                    try:
                        os.remove(chart_path)
                    except:
                        pass
            i += 1
            continue

        # ── Markdown pipe table ──────────────────────────────
        if line.lstrip().startswith("|"):
            table_lines = []
            while i < len(lines):
                l = lines[i].rstrip()
                if not l.lstrip().startswith("|"):
                    break
                table_lines.append(l)
                i += 1
            _render_table(doc, table_lines)
            continue

        if re.match(r"^#{1,2} ", line) and not line.startswith("###"):
            text  = re.sub(r"^#{1,2} ", "", line).strip()
            p     = doc.add_heading(text, level=2)
            _style_section_heading(p, 2)
            i += 1

        elif line.startswith("### "):
            text = line[4:].strip()
            p    = doc.add_heading(text, level=3)
            _style_section_heading(p, 3)
            i += 1

        elif re.match(r"^[\-\*] ", line):
            text = line[2:].strip()
            p    = doc.add_paragraph(style="List Bullet")
            _add_inline_formatting(p, text, fn_mgr)
            p.paragraph_format.left_indent  = Inches(0.3)
            p.paragraph_format.space_after  = Pt(3)
            i += 1

        elif re.match(r"^\d+\. ", line):
            text = re.sub(r"^\d+\. ", "", line).strip()
            p    = doc.add_paragraph(style="List Number")
            _add_inline_formatting(p, text, fn_mgr)
            p.paragraph_format.left_indent  = Inches(0.3)
            p.paragraph_format.space_after  = Pt(3)
            i += 1

        else:
            # ── Check for inline table (line with many | separators) ──
            if line.count("|") >= 4:
                # This looks like a table embedded in a line
                # Try to parse it as a table
                cells = [cell.strip() for cell in line.split("|")]
                cells = [c for c in cells if c]  # remove empty cells

                if len(cells) >= 4:
                    # Collect all consecutive table-like lines
                    table_lines = []
                    table_lines.append(line)
                    i += 1

                    # Collect more table rows
                    while i < len(lines):
                        next_line = lines[i].rstrip()
                        if not next_line.strip() or next_line.count("|") < 4:
                            break
                        table_lines.append(next_line)
                        i += 1

                    # Parse collected table lines
                    if table_lines:
                        parsed_table = []
                        col_count = None

                        for tbl_line in table_lines:
                            row_cells = [cell.strip() for cell in tbl_line.split("|")]
                            row_cells = [c for c in row_cells if c]

                            if row_cells:
                                if col_count is None:
                                    col_count = len(row_cells)

                                # Pad or trim to match column count
                                while len(row_cells) < col_count:
                                    row_cells.append("")
                                row_cells = row_cells[:col_count]

                                pipe_line = "| " + " | ".join(row_cells) + " |"
                                parsed_table.append(pipe_line)

                        # Add separator after header if multiple rows
                        if len(parsed_table) > 1 and col_count:
                            separator = "| " + " | ".join(["---"] * col_count) + " |"
                            parsed_table.insert(1, separator)

                        if parsed_table:
                            _render_table(doc, parsed_table)
                    continue

            # ── Regular paragraph ──
            para_lines = []
            while i < len(lines):
                l = lines[i].rstrip()
                if (not l.strip()
                        or re.match(r"^#{1,3} ", l)
                        or re.match(r"^[\-\*] ", l)
                        or re.match(r"^\d+\. ", l)
                        or l.lstrip().startswith("|")
                        or l.count("|") >= 4):  # Skip inline tables
                    break
                para_lines.append(l)
                i += 1
            text = " ".join(para_lines).strip()
            if text:
                p = doc.add_paragraph()
                _add_inline_formatting(p, text, fn_mgr)
                p.alignment                          = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.space_after       = Pt(6)
                p.paragraph_format.first_line_indent = Inches(0.3)


# ─────────────────────────────────────────────────────────
#  DOCUMENT SECTION BUILDERS
# ─────────────────────────────────────────────────────────

def set_document_defaults(doc):
    style      = doc.styles["Normal"]
    style.font.name  = "Times New Roman"
    style.font.size  = Pt(13)
    pf = style.paragraph_format
    pf.alignment          = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing_rule  = WD_LINE_SPACING.DOUBLE   # true double spacing (2× font height)
    pf.space_after        = Pt(0)


def build_title_page(doc, topic, research_level):
    sec = doc.sections[0]
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin   = Inches(1.5)
    sec.right_margin  = Inches(1.0)

    level_label = LEVEL_PROFILES[research_level]["label"]

    for _ in range(5):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run(topic.upper())
    r.font.size  = Pt(16)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    title.paragraph_format.space_after = Pt(16)


def build_front_matter_page(doc, front_matter_text):
    """Render each ## section of front matter on its own page."""
    # Split the raw text into (heading, body) pairs at every ## marker
    parts = re.split(r"(?m)^(## .+)$", front_matter_text)
    # parts[0] is any text before the first ##; skip it
    # Remaining items come in pairs: heading, body
    sections = []
    i = 1
    while i + 1 < len(parts):
        sections.append((parts[i].strip(), parts[i + 1].strip()))
        i += 2

    for heading, body in sections:
        add_page_break(doc)
        # Render the section heading
        heading_text = re.sub(r"^## ", "", heading).strip()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(heading_text.upper())
        r.font.size  = Pt(14)
        r.font.bold  = True
        r.font.color.rgb = DARK_BLUE
        p.paragraph_format.space_after = Pt(12)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        # Render the section body (strip stray ## heading lines already consumed)
        parse_chapter_content(doc, body)


def extract_chapter_titles_from_custom_toc(custom_toc: str) -> dict:
    """
    Parse custom TOC text and extract chapter titles for chapters 1-5.

    Expected format (flexible):
      CHAPTER ONE: INTRODUCTION
        1.1 Background
      CHAPTER THREE: SYSTEM DESIGN
        3.1 Architecture
      etc.

    Returns dict: {1: "INTRODUCTION", 3: "SYSTEM DESIGN", ...}
    """
    chapter_titles = {}
    if not custom_toc or not custom_toc.strip():
        return chapter_titles

    lines = custom_toc.strip().splitlines()
    for line in lines:
        line = line.strip()
        # Match patterns like: "CHAPTER ONE: INTRODUCTION", "CHAPTER 3: SYSTEM DESIGN"
        match = re.search(r'CHAPTER\s+(?:ONE|1)[:\s–-]+(.+?)(?:\s*$|\s*[:\d])', line, re.IGNORECASE)
        if match:
            title = match.group(1).strip()
            chapter_titles[1] = title.upper()
            continue

        match = re.search(r'CHAPTER\s+(?:TWO|2)[:\s–-]+(.+?)(?:\s*$|\s*[:\d])', line, re.IGNORECASE)
        if match:
            title = match.group(1).strip()
            chapter_titles[2] = title.upper()
            continue

        match = re.search(r'CHAPTER\s+(?:THREE|3)[:\s–-]+(.+?)(?:\s*$|\s*[:\d])', line, re.IGNORECASE)
        if match:
            title = match.group(1).strip()
            chapter_titles[3] = title.upper()
            continue

        match = re.search(r'CHAPTER\s+(?:FOUR|4)[:\s–-]+(.+?)(?:\s*$|\s*[:\d])', line, re.IGNORECASE)
        if match:
            title = match.group(1).strip()
            chapter_titles[4] = title.upper()
            continue

        match = re.search(r'CHAPTER\s+(?:FIVE|5)[:\s–-]+(.+?)(?:\s*$|\s*[:\d])', line, re.IGNORECASE)
        if match:
            title = match.group(1).strip()
            chapter_titles[5] = title.upper()
            continue

    return chapter_titles


def build_toc_page(doc, research_level, chapters_list=None, custom_toc=None,
                   front_matter_sections=None):
    """
    Render the Table of Contents page.

    chapters_list         : list of chapter numbers to include, e.g. [1,3,4,5].
                            Defaults to [1,2,3,4,5].
    custom_toc            : if provided (non-empty string), render the caller-supplied
                            TOC text instead of the auto-generated one.
    front_matter_sections : list of optional sections included in the document,
                            e.g. ["declaration","acknowledgements"].
                            None → all three included.
    """
    chapters_list = chapters_list or list(range(1, 6))
    optional_all  = ["declaration", "dedication", "acknowledgements"]
    if front_matter_sections is None:
        fm_include = optional_all
    else:
        fm_include = [s.lower().strip() for s in front_matter_sections
                      if s.lower().strip() in optional_all]

    add_page_break(doc)

    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run("TABLE OF CONTENTS")
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    hdr.paragraph_format.space_after = Pt(4)
    add_horizontal_rule(doc, color="1F497D", thickness="8")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ── Custom TOC supplied by the user ───────────────────
    if custom_toc and custom_toc.strip():
        for line in custom_toc.strip().splitlines():
            line = line.rstrip()
            if not line:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(2)
                continue
            p  = doc.add_paragraph()
            p.paragraph_format.space_after  = Pt(1)
            p.paragraph_format.line_spacing = Pt(18)
            rn = p.add_run(line)
            rn.font.size      = Pt(11)
            rn.font.color.rgb = RGBColor(0x20, 0x20, 0x20)
        return

    # ── Auto-generated TOC ────────────────────────────────
    # Front matter — only list sections actually included in the document
    entries = []
    if "declaration"      in fm_include: entries.append(("Declaration",      True))
    if "dedication"       in fm_include: entries.append(("Dedication",        True))
    if "acknowledgements" in fm_include: entries.append(("Acknowledgements",  True))
    entries += [
        ("Abstract",               True),
        ("Table of Contents",      True),
        ("List of Tables",         True),
        ("List of Abbreviations",  True),
        ("", False),
    ]

    # Per-chapter TOC groups — only included up to max_chapters
    _chapter_groups = {
        1: [
            ("CHAPTER ONE: INTRODUCTION", True),
            ("  1.1  Background of the Study", False),
            ("  1.2  Statement of the Problem", False),
            ("  1.3  Purpose of the Study", False),
            ("  1.4  Research Objectives", False),
            ("  1.5  Research Questions", False),
            ("  1.6  Significance of the Study", False),
            ("  1.7  Scope and Delimitations", False),
            ("  1.8  Limitations of the Study", False),
            ("  1.9  Definition of Key Terms", False),
            ("  1.10 Organisation of the Study", False),
        ],
        2: [
            ("CHAPTER TWO: LITERATURE REVIEW", True),
            ("  2.1  Introduction to the Chapter", False),
            ("  2.2  Conceptual Review", False),
            ("  2.3  Theoretical Framework", False),
            ("  2.4  Empirical Review", False),
            ("  2.5  Review of Related Studies", False),
            ("  2.6  Research Gap", False),
            ("  2.7  Chapter Summary", False),
        ],
        3: [
            ("CHAPTER THREE: RESEARCH METHODOLOGY", True),
            ("  3.1  Introduction to the Chapter", False),
            ("  3.2  Research Design", False),
            ("  3.3  Research Philosophy and Paradigm", False),
            ("  3.4  Research Approach", False),
            ("  3.5  Study Area and Setting", False),
            ("  3.6  Target Population", False),
            ("  3.7  Sample Size and Sampling Technique", False),
            ("  3.8  Data Collection Instruments", False),
            ("  3.9  Validity and Reliability", False),
            ("  3.10 Data Collection Procedure", False),
            ("  3.11 Data Analysis Methods", False),
            ("  3.12 Ethical Considerations", False),
            ("  3.13 Chapter Summary", False),
        ],
        4: [
            ("CHAPTER FOUR: RESULTS AND DISCUSSION", True),
            ("  4.1  Introduction to the Chapter", False),
            ("  4.2  Sample / Response Rate Overview", False),
            ("  4.3  Findings Related to Objective 1", False),
            ("  4.4  Findings Related to Objective 2", False),
            ("  4.5  Findings Related to Objective 3", False),
            ("  4.6  Findings Related to Objective 4", False),
            ("  4.7  Synthesis and Discussion of Major Findings", False),
            ("  4.8  Implications of the Findings", False),
            ("  4.9  Chapter Summary", False),
        ],
        5: [
            ("CHAPTER FIVE: CONCLUSIONS AND RECOMMENDATIONS", True),
            ("  5.1  Introduction to the Chapter", False),
            ("  5.2  Summary of the Study", False),
            ("  5.3  Conclusions", False),
            ("  5.4  Contribution to Knowledge", False),
            ("  5.5  Recommendations", False),
            ("  5.6  Recommendations for Future Research", False),
            ("  5.7  Chapter Summary", False),
        ],
    }

    for n in chapters_list:
        entries.extend(_chapter_groups[n])
        entries.append(("", False))

    if 5 in chapters_list:
        entries += [("References", True), ("Appendices", True)]

    for text, is_bold in entries:
        if not text.strip():
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            continue
        p  = doc.add_paragraph()
        p.paragraph_format.space_after   = Pt(1)
        p.paragraph_format.line_spacing  = Pt(18)
        rn = p.add_run(text)
        rn.font.size  = Pt(11)
        rn.bold       = is_bold
        rn.font.color.rgb = DARK_BLUE if is_bold else RGBColor(0x20, 0x20, 0x20)
        dots = max(2, 65 - len(text))
        p.add_run(" " + ("." * dots) + " ")
        p.add_run("____").font.size = Pt(10)


def build_list_of_tables_page(doc):
    """Build a List of Tables page."""
    add_page_break(doc)

    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run("LIST OF TABLES")
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    hdr.paragraph_format.space_after = Pt(4)
    add_horizontal_rule(doc, color="1F497D", thickness="8")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Count tables in the document
    table_count = len(doc.tables)

    if table_count == 0:
        # No tables found
        p = doc.add_paragraph("No tables in this document.")
        p.paragraph_format.space_after = Pt(6)
        return

    # Generate list of tables
    for table_idx, table in enumerate(doc.tables, 1):
        # Create table entry with number and description
        # Try to extract a meaningful description from the first row
        if table.rows:
            first_row_text = " | ".join([cell.text[:15] for cell in table.rows[0].cells])
            caption = f"Table {table_idx}: {first_row_text}..."
        else:
            caption = f"Table {table_idx}"

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = Pt(18)
        p.paragraph_format.left_indent = Inches(0.3)

        rn = p.add_run(caption)
        rn.font.size = Pt(11)
        rn.font.color.rgb = RGBColor(0x20, 0x20, 0x20)

        # Add dots and page number placeholder
        dots = max(2, 55 - len(caption))
        p.add_run(" " + ("." * dots) + " ")
        p.add_run("____").font.size = Pt(10)


def build_list_of_figures_page(doc):
    """Build a List of Figures page (includes charts and images)."""
    add_page_break(doc)

    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run("LIST OF FIGURES")
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    hdr.paragraph_format.space_after = Pt(4)
    add_horizontal_rule(doc, color="1F497D", thickness="8")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Count drawing elements (charts/images/figures) in the document
    figure_count = 0
    for element in doc.element.body:
        # Count paragraphs that contain drawing elements
        if element.tag.endswith('}p'):
            for child in element.iter():
                if 'drawing' in child.tag.lower():
                    figure_count += 1
                    break

    if figure_count == 0:
        # No figures found
        p = doc.add_paragraph("No figures in this document.")
        p.paragraph_format.space_after = Pt(6)
        return

    # Generate list of figures
    figure_idx = 0
    for para in doc.paragraphs:
        # Check if paragraph contains images/drawings
        has_drawing = False
        for run in para.runs:
            if run._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing') is not None:
                has_drawing = True
                break

        if has_drawing:
            figure_idx += 1
            # Try to get caption from next paragraph if it's italicized
            caption = f"Figure {figure_idx}"

            # Look for italicized caption in following paragraphs
            try:
                para_idx = doc.paragraphs.index(para)
                if para_idx + 1 < len(doc.paragraphs):
                    next_para = doc.paragraphs[para_idx + 1]
                    if next_para.runs and next_para.runs[0].italic:
                        caption = f"Figure {figure_idx}: {next_para.text[:50]}"
            except:
                pass

            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = Pt(18)
            p.paragraph_format.left_indent = Inches(0.3)

            rn = p.add_run(caption)
            rn.font.size = Pt(11)
            rn.font.color.rgb = RGBColor(0x20, 0x20, 0x20)

            # Add dots and page number placeholder
            dots = max(2, 55 - len(caption))
            p.add_run(" " + ("." * dots) + " ")
            p.add_run("____").font.size = Pt(10)


def build_abbreviations_page(doc):
    add_page_break(doc)
    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run("LIST OF ABBREVIATIONS")
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    hdr.paragraph_format.space_after = Pt(4)
    add_horizontal_rule(doc, color="1F497D", thickness="8")

    intro = doc.add_paragraph()
    intro.add_run("The following abbreviations are used throughout this document:").italic = True
    intro.paragraph_format.space_after = Pt(8)

    for abbr, meaning in [
        ("APA",    "American Psychological Association"),
        ("GDP",    "Gross Domestic Product"),
        ("ICT",    "Information and Communication Technology"),
        ("NGO",    "Non-Governmental Organisation"),
        ("UN",     "United Nations"),
        ("WHO",    "World Health Organization"),
        ("SPSS",   "Statistical Package for the Social Sciences"),
        ("e.g.",   "For example (Latin: exempli gratia)"),
        ("i.e.",   "That is (Latin: id est)"),
        ("et al.", "And others (Latin: et alii)"),
    ]:
        p  = doc.add_paragraph(style="List Bullet")
        rb = p.add_run(f"{abbr}:  ")
        rb.bold = True
        p.add_run(meaning)
        p.paragraph_format.space_after = Pt(3)


def build_chapter_page(doc, chapter_num, chapter_content, fn_mgr=None, custom_subtitle=None):
    add_page_break(doc)
    add_chapter_header(doc, chapter_num, custom_subtitle=custom_subtitle)
    parse_chapter_content(doc, chapter_content, fn_mgr=fn_mgr)


def build_document(topic: str, research_level: str,
                   front_matter: str, chapters: dict,
                   output_dir: str) -> str:
    doc = Document()
    set_document_defaults(doc)

    build_title_page(doc, topic, research_level)
    build_front_matter_page(doc, front_matter)
    build_toc_page(doc, research_level)
    build_abbreviations_page(doc)

    # Build all chapters (this creates the tables)
    for num in range(1, 6):
        build_chapter_page(doc, num, chapters[num])

    # Build List of Figures after all chapters (so we can count all figures/charts)
    build_list_of_figures_page(doc)

    # Build List of Tables after all chapters (so we can count all tables)
    build_list_of_tables_page(doc)

    # Sanitize topic for Windows-safe filename
    # Remove or replace invalid Windows filename characters: < > : " / \ | ? *
    safe = topic
    for char in '<>:"/\\|?*':
        safe = safe.replace(char, '_')
    # Keep only alphanumeric, spaces, underscores, and hyphens
    safe = re.sub(r'[^\w\s\-]', '', safe).strip().replace(" ", "_")
    # Limit length and ensure non-empty
    safe = safe[:50] if safe else "Document"

    filename = f"Research_{safe}.docx"
    path     = os.path.join(output_dir, filename)
    doc.save(path)
    return path


# ─────────────────────────────────────────────────────────
#  CLAUDE API — CONTENT GENERATION
# ─────────────────────────────────────────────────────────

def _stream_content(client, system: str, prompt: str,
                    model: str, max_tokens: int) -> str:
    use_thinking = model in ("claude-opus-4-6", "claude-sonnet-4-6")

    THINKING_BUDGET = 8000   # tokens reserved for Claude's internal reasoning
    MIN_OUTPUT      = 12000  # minimum tokens guaranteed for actual text output

    if use_thinking:
        # max_tokens must cover thinking budget + output budget
        # Never let adaptive thinking swallow the output capacity
        actual_max = max(max_tokens, THINKING_BUDGET + MIN_OUTPUT)
        kwargs = dict(
            model=model,
            max_tokens=actual_max,
            thinking={"type": "enabled", "budget_tokens": THINKING_BUDGET},
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )
    else:
        kwargs = dict(
            model=model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )

    with client.messages.stream(**kwargs) as stream:
        return "".join(stream.text_stream)


def generate_front_matter(client, topic: str, research_level: str,
                           model: str = None,
                           front_matter_sections: list = None,
                           custom_instructions: str = None) -> str:
    """
    Generate front matter pages.

    front_matter_sections : list of optional sections to include.
        Allowed values: "declaration", "dedication", "acknowledgements"
        Abstract is ALWAYS included regardless of this list.
        Default (None) → all three optional sections included.
    """
    model   = model or config.MODEL
    profile = LEVEL_PROFILES[research_level]

    # Resolve which optional sections to include
    optional_all = ["declaration", "dedication", "acknowledgements"]
    if front_matter_sections is None:
        include = optional_all
    else:
        include = [s.lower().strip() for s in front_matter_sections
                   if s.lower().strip() in optional_all]

    system = (
        f"You are a senior academic writer working at {profile['label']} level. "
        "Write formal, substantive academic front matter. Use ## for each section heading. "
        "Write with authenticity and warmth where appropriate — these pages are read by "
        "examiners and set the tone for the entire document. "
        + HUMAN_WRITING_INSTRUCTION
    )

    # Build the section list dynamically
    section_blocks = []

    if "declaration" in include:
        section_blocks.append(
            "## DECLARATION\n"
            "Write a formal, original academic declaration of authorship (120–160 words). "
            "Include: a statement that the work is the researcher's own, that all sources "
            "have been properly cited, that the work has not been submitted elsewhere for "
            "examination, and acknowledgement that plagiarism constitutes academic misconduct. "
            "The tone should be formal and confident, not mechanical."
        )
    if "dedication" in include:
        section_blocks.append(
            "## DEDICATION\n"
            "Write a heartfelt, personal dedication (60–90 words). Dedicate to specific "
            "named people (family members, mentors, or a community). The dedication should "
            "feel genuine — avoid generic phrases. It should be warm, brief, and memorable."
        )
    if "acknowledgements" in include:
        section_blocks.append(
            "## ACKNOWLEDGEMENTS\n"
            "Write genuine, specific acknowledgements (250–320 words). Thank in separate "
            "sentences or short paragraphs: the academic supervisor (by title and role), "
            "the institution and department, research participants (without naming them), "
            "colleagues or peers who provided feedback, family and close supporters. "
            "The tone should be warm and personal, not formulaic. Avoid generic praise — "
            "be specific about what each person contributed."
        )

    abstract_word_min = profile['front_words'] // 2
    abstract_word_max = abstract_word_min + 80
    pg_note = (
        "For postgraduate level: include a sentence on epistemological positioning, "
        "the theoretical framework used, and the study's contribution to theory. "
        if research_level == "postgraduate" else ""
    )
    section_blocks.append(
        f"## ABSTRACT\n"
        f"Write a structured abstract of {abstract_word_min}–{abstract_word_max} words covering: "
        f"(1) background and problem statement, (2) research objectives, "
        f"(3) methodology and data collection approach, (4) principal findings, "
        f"(5) conclusions and recommendations. {pg_note}"
        f"End with: Keywords: [5 relevant academic keywords separated by semicolons]."
    )

    prompt = (
        f"For an academic research project titled: **{topic}**\n"
        f"Research level: {profile['label']}\n\n"
        "Write the following front matter sections. Each section must be fully developed "
        "to the word counts specified. Use ## to introduce each section heading.\n\n"
        + "\n\n".join(section_blocks)
        + "\n\nStart directly with the first ## heading. Do not add any preamble."
    )

    if custom_instructions and custom_instructions.strip():
        prompt += (
            f"\n\n--- ADDITIONAL INSTRUCTIONS ---\n"
            f"{custom_instructions.strip()}\n"
            f"--- END ADDITIONAL INSTRUCTIONS ---"
        )

    print("  [Front Matter] generating...", end=" ", flush=True)
    text = _stream_content(client, system, prompt, model, 5000)
    print(f"done ({len(text):,} chars)")
    return text


def generate_chapter(client, topic: str, chapter_num: int,
                     research_level: str, model: str = None,
                     custom_instructions: str = None) -> str:
    model    = model or config.MODEL
    prompts  = _chapter_prompts(research_level)
    prompt   = prompts[chapter_num].format(topic=topic)
    profile  = LEVEL_PROFILES[research_level]
    target   = profile["word_targets"][chapter_num]

    system = (
        f"You are a highly experienced human academic researcher writing at {profile['label']} level. "
        "You have spent years publishing in peer-reviewed journals and supervising postgraduate students. "
        "Your writing is authoritative, specific, and unmistakably human — characterised by varied rhythm, "
        "genuine intellectual engagement, specific citations, occasional hedging, and a distinct scholarly voice. "
        "You never produce AI-sounding text. You write fully developed, substantive prose and never truncate, "
        "summarise, or leave placeholders. You meet every word count target without compromise."
    )

    if custom_instructions and custom_instructions.strip():
        prompt += (
            f"\n\n--- ADDITIONAL INSTRUCTIONS ---\n"
            f"{custom_instructions.strip()}\n"
            f"--- END ADDITIONAL INSTRUCTIONS ---"
        )

    print(f"  [Ch {chapter_num}] {CHAPTER_SUBTITLES[chapter_num]}...", end=" ", flush=True)
    # Allow generous token budget: ~1.4 tokens/word, plus headroom for instructions output
    text = _stream_content(client, system, prompt, model, max(10000, int(target * 3.5)))
    print(f"done ({len(text):,} chars)")
    return text


# ─────────────────────────────────────────────────────────
#  CLI ENTRY POINT
# ─────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("   ACADEMIC RESEARCH WRITEUP AGENT  —  5-CHAPTER FORMAT")
    print("=" * 60)

    topic = input("\nResearch topic:\n> ").strip()
    level = input("Research level (undergraduate / postgraduate):\n> ").strip().lower()
    if level not in LEVEL_PROFILES:
        print("Defaulting to undergraduate.")
        level = "undergraduate"

    os.environ["ANTHROPIC_API_KEY"] = config.ANTHROPIC_API_KEY
    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)

    front    = generate_front_matter(client, topic, level)
    chapters = {n: generate_chapter(client, topic, n, level) for n in range(1, 6)}

    out = build_document(topic, level, front, chapters, os.getcwd())
    print(f"\nDocument saved: {out}")


if __name__ == "__main__":
    main()
