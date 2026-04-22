"""
Academic Research Writeup Agent — 5-Chapter Format
Generates a complete 5-chapter academic research project in Microsoft Word (.docx).
Supports undergraduate and postgraduate research levels.
"""

import os
import sys
import re
import anthropic

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
            "Each main subsection should be at least 300-400 words of substantive prose."
        ),
        "depth":        "substantive but accessible",
        "word_targets": {1: 2400, 2: 3400, 3: 3800, 4: 2600, 5: 2400},
        "front_words":  600,
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
            "Each main subsection should be at least 450-600 words of dense, substantive prose."
        ),
        "depth":        "critical, theoretically sophisticated, and reflexive",
        "word_targets": {1: 3200, 2: 4600, 3: 5200, 4: 3800, 5: 3200},
        "front_words":  850,
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

    return {
        1: f"""You are writing CHAPTER ONE — INTRODUCTION for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[1]} words of substantive prose. You MUST reach this minimum.
Do not stop writing until you have fully developed every subsection. If in doubt, write more.

{tone}

{HUMAN_WRITING_INSTRUCTION}

Write the following subsections, each introduced with a ### heading.
Every subsection must be written in full, developed paragraphs — no bullet summaries, no placeholders.

### 1.1 Background of the Study
Write at least {"500" if is_pg else "350"} words for this subsection.
Provide {depth} contextual grounding. Open with a striking observation or statistic that
immediately establishes why this topic matters. Then trace the historical evolution of the
problem across at least three distinct time periods, naming key turning points, policy shifts,
or scholarly debates that shaped the current landscape. Ground every assertion in specific
evidence — named scholars, years, places, and figures. Close by narrowing the lens from the
broad context toward the precise issue this study addresses.

### 1.2 Statement of the Problem
Write at least {"400" if is_pg else "280"} words for this subsection.
Open with a clear, declarative statement of what is wrong or poorly understood. Then build
the case across multiple paragraphs: explain the nature of the problem, who it affects, how
long it has persisted, and why existing responses have been insufficient. Name the specific
gap, contradiction, or blind spot that this study addresses. The problem statement must feel
urgent — the reader should finish this section convinced that the study was necessary.

### 1.3 Purpose of the Study
Write at least {"200" if is_pg else "150"} words for this subsection.
State the overarching aim in one or two precise sentences. Then elaborate: explain the
theoretical and practical orientation of the study, what kind of knowledge it seeks to
produce, and how the purpose connects directly to the problem articulated in 1.2.

### 1.4 Research Objectives
Write at least {"250" if is_pg else "180"} words for this subsection.
State 4–5 specific, measurable objectives. Each should be action-oriented (examine, assess,
determine, explore, compare, evaluate). After listing them, write a short paragraph explaining
how they collectively address the research problem and how they will be operationalised
through the methodology described in Chapter 3.

### 1.5 Research Questions
Write at least {"200" if is_pg else "150"} words for this subsection.
Formulate 3–5 focused, answerable questions derived from the objectives. After stating the
questions, briefly explain the logic connecting each question to its corresponding objective
and the type of evidence that would constitute an answer.

### 1.6 Significance of the Study
Write at least {"350" if is_pg else "250"} words for this subsection.
Develop this across multiple distinct paragraphs — at minimum one paragraph each for:
theoretical significance (what this study adds to scholarly debates), practical significance
(what changes in practice, policy, or service delivery), and beneficiary groups (academics,
practitioners, policymakers, communities). Be concrete rather than generic — name the
specific journals, policy areas, or organisations that would benefit.

### 1.7 Scope and Delimitations
Write at least {"280" if is_pg else "200"} words for this subsection.
Define the geographic, temporal, and thematic boundaries with precision. For each
boundary, explain not just what is excluded but why the exclusion is methodologically
justified rather than a limitation of convenience. Acknowledge the trade-offs involved.

### 1.8 Limitations of the Study
Write at least {"280" if is_pg else "200"} words for this subsection.
Identify at least four genuine constraints — methodological, practical, or contextual.
For each, explain what the limitation is, how it arose, and what steps were taken to
minimise its impact on the validity and transferability of findings. Be candid: real
researchers acknowledge imperfection.

### 1.9 Definition of Key Terms
Write at least {"300" if is_pg else "220"} words for this subsection.
Define 6–8 terms that carry specific technical or conceptual meanings in this study.
For each term: provide a working definition grounded in at least one cited scholar,
explain how this study's usage compares to or departs from common usage, and note any
definitional controversies relevant to the research.

### 1.10 Organisation of the Study
Write at least {"180" if is_pg else "130"} words for this subsection.
Describe what each chapter covers in two to three sentences per chapter — not a list,
but short, flowing paragraphs. Explain the logical progression from chapter to chapter.

Do NOT write a chapter title heading at the very top — begin directly with section ### 1.1.""",

        2: f"""You are writing CHAPTER TWO — LITERATURE REVIEW for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[2]} words of substantive prose. You MUST reach this minimum.
The literature review is the longest and most intellectually demanding chapter. Write with depth.

{tone}

{HUMAN_WRITING_INSTRUCTION}

Write the following subsections in full. Every subsection demands extended, analytical prose.

### 2.1 Introduction to the Chapter
Write at least {"300" if is_pg else "200"} words.
Open by situating the literature review within the study's broader purpose. Explain how
this chapter is organised and why that organisational logic was chosen. Describe the scope
of literature reviewed — databases, date range, inclusion criteria — without being mechanical.
End with a statement of what the review reveals and how it sets up the research gap.

### 2.2 Conceptual Review
Write at least {"600" if is_pg else "420"} words.
Identify the 4–6 central concepts of this study. For each concept: trace its intellectual
history (who coined or defined it, when, and in what context), map the range of definitions
across the literature (noting where scholars converge and diverge), and state explicitly
which conceptualisation this study adopts and why. Write this as connected analytical prose,
not as a series of dictionary definitions.
{"Engage with conceptual tensions and competing paradigms — do not smooth them over." if is_pg else ""}

### 2.3 Theoretical Framework
Write at least {"700" if is_pg else "480"} words.
Identify 2–3 theories or models that directly inform this study. For each theory, develop
a full sub-argument across multiple paragraphs: name the originator and intellectual
tradition, describe the core propositions, trace how it has been applied and tested in
empirical research over the past decade, and make explicit how it will guide this study's
analytical framework. {"Critically evaluate each theory — identify its explanatory strengths, its known limitations, and how scholars have critiqued or refined it." if is_pg else "Explain how each theory applies to the specific context of this study."}

### 2.4 Empirical Review
Write at least {"900" if is_pg else "650"} words.
Critically review at least {"12-15" if is_pg else "8-10"} prior studies. Organise the
review thematically rather than chronologically. For each thematic cluster: identify the
key studies, summarise their findings and methodological approaches, note where results
converge, flag contradictions or anomalies in the evidence base, and comment on methodological
quality. {"Evaluate sample sizes, research designs, and contextual applicability." if is_pg else ""}
This section must read as a genuine scholarly conversation, not a descriptive catalogue.

### 2.5 Review of Related Studies
Write at least {"600" if is_pg else "420"} words.
Focus specifically on studies conducted in comparable contexts or addressing analogous
sub-questions. For each study reviewed: explain what it investigated, summarise its
principal findings, assess what it contributes to this study's conceptual or empirical
foundations, and — critically — identify precisely where it falls short relative to the
present study's aims. This section should make the research gap feel inevitable.

### 2.6 Research Gap
Write at least {"400" if is_pg else "280"} words.
Do not simply assert that a gap exists — argue for it. Draw together the evidence from the
preceding sections to show exactly what has been studied, what remains unstudied, why the
existing studies are insufficient for this particular problem, and why this gap matters.
{"Distinguish between empirical gaps (what data are missing), theoretical gaps (what explanatory frameworks have not been tested here), and methodological gaps (how prior studies' designs could be improved)." if is_pg else "Make clear why filling this gap produces knowledge that is both novel and useful."}

### 2.7 Chapter Summary
Write at least {"350" if is_pg else "250"} words.
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

{tone}

{HUMAN_WRITING_INSTRUCTION}

Write the following subsections in full.

### 3.1 Introduction to the Chapter
Write at least {"250" if is_pg else "180"} words.
Orient the reader to the chapter's purpose and structure. Explain the epistemological logic
that connects the research questions to the design choices made. {"State the researcher's ontological and epistemological position upfront and explain how it shapes the chapter's approach to the treatment of evidence and knowledge claims." if is_pg else "Explain how the methodology flows from the research questions and problem."}

### 3.2 Research Design
Write at least {"450" if is_pg else "320"} words.
Describe the overall research strategy and justify the choice of qualitative, quantitative,
or mixed-methods design by reference to the nature of the research questions. Cite at least
three methodologists who support this design choice. Explain what this design can and cannot
do — including what it sacrifices — and defend the choice against obvious alternatives.
{"Connect the design explicitly to the epistemological position stated in 3.1." if is_pg else ""}

### 3.3 Research Philosophy and Paradigm
Write at least {"500" if is_pg else "320"} words.
{"Develop the philosophical grounding in detail. Discuss the ontological position (what the researcher believes about the nature of reality — is it singular and knowable, or multiple and constructed?), the epistemological position (what counts as valid knowledge, and how it can be acquired), and how these positions connect to the chosen methodology. Distinguish between positivism, interpretivism, constructivism, pragmatism, and critical realism with enough precision that the reader understands which stance is adopted here and why." if is_pg else "Identify the research paradigm (e.g., interpretivist, positivist, pragmatist) and explain in clear terms how it shapes the study's approach to data, evidence, and knowledge. Draw on at least two methodologists to justify the paradigmatic choice."}

### 3.4 Research Approach
Write at least {"280" if is_pg else "200"} words.
{"Specify whether the study uses inductive, deductive, or abductive reasoning. Justify this choice by reference to the research questions and the nature of the evidence being collected. Explain how the approach shapes the analytical process in Chapter 4." if is_pg else "Specify the reasoning approach (inductive or deductive) and explain how it guides data analysis. Connect this to the research design."}

### 3.5 Study Area and Setting
Write at least {"320" if is_pg else "230"} words.
Describe the physical, institutional, or organisational setting with enough specificity that
the reader can visualise it. Explain why this setting was chosen — what makes it appropriate
for answering these research questions. Discuss access, gatekeeping, and any contextual
factors (political, cultural, institutional) that shaped the fieldwork.

### 3.6 Target Population
Write at least {"280" if is_pg else "200"} words.
Define the population with precision — who qualifies, why they qualify, and how large the
total population is (with a source if applicable). Explain the relevance of this population
to the research questions. Address any challenges in defining or accessing the population.

### 3.7 Sample Size and Sampling Technique
Write at least {"380" if is_pg else "270"} words.
Specify the sample size and justify it — cite at least two sources on sample size adequacy
for the chosen design. Describe the sampling technique in precise operational terms: exactly
how participants were identified, approached, screened, and recruited. {"Discuss how the technique addresses issues of representativeness (quantitative) or theoretical saturation and transferability (qualitative)." if is_pg else "Explain how the sample is representative of the population."} Address any non-response and how it was handled.

### 3.8 Data Collection Instruments
Write at least {"400" if is_pg else "280"} words.
Describe each instrument used (questionnaire, semi-structured interview guide, observation
protocol, document analysis schedule). For each instrument: explain the rationale for its
design, describe its structure (sections, item types, scale formats), explain the piloting
process and any revisions made, and justify its appropriateness for collecting the data
required by each research objective.

### 3.9 Validity and Reliability
Write at least {"380" if is_pg else "270"} words.
{"Address validity and reliability using the criteria appropriate to the paradigm. For quantitative work: construct validity, criterion validity, internal consistency (Cronbach's alpha), and test-retest reliability. For qualitative work: credibility (member-checking, triangulation), transferability (thick description), dependability (audit trail), and confirmability (reflexivity) — drawing on Lincoln and Guba (1985). Explain specifically how each criterion was operationalised in this study." if is_pg else "Explain what steps were taken to ensure the instruments measure what they intend to measure and produce consistent results. Discuss any piloting and revision process. Address both internal validity and reliability."}

### 3.10 Data Collection Procedure
Write at least {"350" if is_pg else "250"} words.
Describe the data collection process as a step-by-step chronological narrative: ethics
clearance, participant recruitment, informed consent, instrument administration, data
recording, and quality checks. Include time frames and quantities (how many interviews
conducted over how many weeks, response rate for questionnaires). Be specific enough
that a researcher could replicate this procedure.

### 3.11 Data Analysis Methods
Write at least {"400" if is_pg else "280"} words.
Explain the analytical approach in enough detail for replication. {"Name the specific software used (SPSS, NVivo, Atlas.ti, R, Python) and justify the choice. Describe the analytical procedures step by step: coding (open, axial, selective), thematic analysis phases, statistical tests applied and their assumptions, regression models and their specification. Connect each analytical step to the specific research questions it addresses." if is_pg else "Describe the analytical approach clearly: how data were organised, coded, and interpreted. Name any software used and explain how it was applied. Connect the analysis to the research questions."}

### 3.12 Ethical Considerations
Write at least {"300" if is_pg else "220"} words.
Address at least six ethical dimensions: informed consent (what participants were told and
how consent was obtained), anonymity and confidentiality (how data were anonymised and
protected), right to withdraw (how this was communicated and facilitated), data storage
and security (how data are stored and for how long), institutional ethics approval
(institution and reference number if applicable), and researcher positionality
(how the researcher's background may have influenced data collection and interpretation).

### 3.13 Chapter Summary
Write at least {"280" if is_pg else "200"} words.
Synthesise the methodological choices made in this chapter as a coherent whole. Explain
how design, philosophy, sampling, instruments, and analysis hang together as a unified
approach to answering the research questions. {"Address how the methodology addresses the research gap identified in Chapter 2 and positions the study within its paradigmatic tradition." if is_pg else "Show how the methodology directly serves the research objectives stated in Chapter 1."}

Do NOT write a chapter title heading at the very top — begin directly with section ### 3.1.""",

        4: f"""You are writing CHAPTER FOUR — RESULTS AND DISCUSSION for an academic research project.
Topic: {{topic}}
Research level: {profile['label']}
MINIMUM word count: {targets[4]} words of substantive prose. You MUST reach this minimum.
Present rich, specific, interpreted findings. This chapter must demonstrate analytical depth.

{tone}

{HUMAN_WRITING_INSTRUCTION}

Write the following subsections in full.

### 4.1 Introduction to the Chapter
Write at least {"250" if is_pg else "180"} words.
Explain how the chapter is structured and why. Briefly recap the research objectives so
the reader knows what findings will address. {"State the analytical framework guiding interpretation and how it connects to the theoretical framework in Chapter 2." if is_pg else "Orient the reader to how findings are organised."}

### 4.2 Sample / Response Rate Overview
Write at least {"300" if is_pg else "220"} words.
Present the demographic and descriptive profile of the sample across multiple characteristics
(age, gender, education, experience, geographic distribution — as relevant). Discuss the
response rate if applicable and explain patterns in non-response. {"Compare the achieved sample to the target population and discuss implications for transferability." if is_pg else "Comment on how representative the sample appears to be."}

### 4.3 Findings Related to Objective 1
Write at least {"500" if is_pg else "360"} words.
Present specific, detailed findings for the first research objective. Use plausible
quantitative values (percentages, means, frequencies) or qualitative themes with
representative illustrative evidence. Interpret the findings rather than just reporting
them: explain what patterns emerge, what they mean, and what accounts for them.
{"Connect findings explicitly to the theoretical framework from Chapter 2. Where results confirm prior theory, explain why. Where they challenge it, explore the implications." if is_pg else "Relate findings directly to relevant literature reviewed in Chapter 2."}

### 4.4 Findings Related to Objective 2
Write at least {"500" if is_pg else "360"} words.
Apply the same approach as 4.3 to the second research objective. Ensure this section has
its own narrative arc — do not simply replicate the structure of 4.3. Introduce any
unexpected or contradictory findings and engage with them analytically.

### 4.5 Findings Related to Objective 3
Write at least {"500" if is_pg else "360"} words.
Apply the same approach to the third objective. {"At this stage, begin drawing connections between findings across objectives — note where patterns reinforce each other or where tensions emerge." if is_pg else "Discuss how these findings relate to those in 4.3 and 4.4."}

### 4.6 Findings Related to Objective 4
Write at least {"450" if is_pg else "320"} words.
Present findings for the fourth objective with the same analytical rigour. By the end of
this section, all major findings should be on the table, setting up the synthesis in 4.7.

### 4.7 Synthesis and Discussion of Major Findings
Write at least {"600" if is_pg else "420"} words.
{"This is the intellectual heart of the chapter. Do not merely summarise the preceding sections. Instead, synthesise: identify overarching themes that cut across the four sets of findings, explore unexpected results and what they suggest, address contradictions between data sources, and situate the findings in relation to the theoretical framework and empirical literature from Chapter 2. Where findings confirm prior scholarship, say so with precision. Where they challenge or extend it, develop that argument fully." if is_pg else "Bring together the key patterns across all four sets of findings. Identify the most important themes that emerge when the findings are considered as a whole. Connect them to the literature reviewed in Chapter 2 — where do findings confirm, contradict, or extend existing knowledge?"}

### 4.8 Implications of the Findings
Write at least {"400" if is_pg else "280"} words.
Discuss implications for theory, practice, and policy separately across dedicated paragraphs.
Name specific stakeholders and explain precisely what each set of findings means for them.
{"For theory: what does this study add to, refine, or challenge in the existing theoretical models? For practice: what specific changes in professional practice are warranted? For policy: what specific policy recommendations emerge, addressed to named agencies or decision-makers?" if is_pg else "Be concrete: name institutions, policy areas, and professional communities that should act on these findings."}

### 4.9 Chapter Summary
Write at least {"280" if is_pg else "200"} words.
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

Write the following subsections in full.

### 5.1 Introduction to the Chapter
Write at least {"200" if is_pg else "150"} words.
Orient the reader to the chapter's purpose and structure. Briefly explain how this chapter
brings the entire study to a close and what it aims to deliver beyond simply summarising
earlier chapters.

### 5.2 Summary of the Study
Write at least {"450" if is_pg else "320"} words.
Recount the entire research journey in a flowing, synthesised narrative across at least
four substantive paragraphs: the problem and its context, the objectives and theoretical
framework, the methodology and its justification, and the principal findings. Do not
quote verbatim from earlier chapters — reframe and integrate. A reader encountering this
study for the first time through this section should understand its full arc.

### 5.3 Conclusions
Write at least {"500" if is_pg else "360"} words.
Draw one specific, argued conclusion per research objective — each conclusion in its own
paragraph. Each conclusion must: state what the study found, explain what this finding
means in context, and connect it to the evidence from Chapter 4. {"Where conclusions are tentative or conditional, say so and explain the conditions under which the conclusion holds. Where they challenge prior theory, develop that challenge explicitly." if is_pg else "State conclusions with appropriate confidence — neither overclaiming nor underselling what the data support."}

### 5.4 Contribution to Knowledge
Write at least {"380" if is_pg else "270"} words.
{"Articulate the study's contribution across at least three dimensions: theoretical (how it extends, refines, or challenges existing theoretical models), empirical (what new data or patterns it adds to the evidence base), and methodological (whether it demonstrates a novel application of method in this context). Be precise — 'this study contributes to the literature' is not a contribution; naming exactly what it adds is." if is_pg else "Explain in concrete terms what is new or valuable about what this study found. How does it advance understanding beyond what was known before? What practical problems does it help solve?"}

### 5.5 Recommendations
Write at least {"450" if is_pg else "320"} words.
Provide 6–8 specific, actionable, evidence-grounded recommendations. Write each as a
full paragraph rather than a bullet point: name the recommendation, identify the specific
finding that supports it, name the stakeholder or institution it is directed at, and
describe what implementing it would look like in practice. Recommendations must flow
directly from the findings — no recommendation should appear without a grounding in
Chapter 4.

### 5.6 Recommendations for Future Research
Write at least {"350" if is_pg else "250"} words.
Propose 4–5 specific research directions that arise from this study's limitations or from
questions it raised but could not answer. Each recommendation for future research should:
identify the gap or question, explain why it matters, suggest an appropriate methodological
approach, and state what such research would contribute. {"For postgraduate work, these should point toward theoretical refinement, comparative cross-context studies, or longitudinal designs." if is_pg else ""}

### 5.7 Chapter Summary
Write at least {"220" if is_pg else "160"} words.
A dignified, forward-looking closing that does not merely repeat the conclusions. Reflect
on what the study set out to do and what it achieved. End with a final paragraph that
gestures toward the broader significance of the work — without overreaching.

---

After 5.7, write the following two sections:

## REFERENCES
List at least {"25" if is_pg else "18"} academic references in APA 7th edition format.
References must be plausible, field-relevant, diverse, and correctly formatted.
Include: journal articles (majority), books, book chapters, institutional/government
reports, and conference papers. Span at least 2005–2023. Mix foundational texts with
recent scholarship (at least 8 references from 2018 onwards).
Format each exactly as:
  Author, A. A., & Author, B. B. (Year). Title of article. Journal Name, Volume(Issue), pages. https://doi.org/xxxxx

## APPENDICES

### Appendix A: Research Instrument
Provide a complete {"interview guide (20+ open and semi-structured questions across thematic sections)" if is_pg else "questionnaire (20+ items using Likert scales, multiple choice, and open-ended questions)"}, appropriate to the research design described in Chapter 3. Include an introduction/preamble and section headings.

### Appendix B: Data Collection Timeline
A structured {"10" if is_pg else "8"}-week timeline table for the data collection phase, with activities, responsible parties, and expected outputs for each week.

### Appendix C: Ethical Clearance Template
A sample informed consent form that would be used with participants in this study,
including all required elements (study description, risks, rights, confidentiality, contact details).

Do NOT write a chapter title heading at the very top — begin directly with section ### 5.1.""",
    }


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


def add_chapter_header(doc, chapter_num):
    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    label = doc.add_paragraph()
    label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = label.add_run(CHAPTER_TITLES[chapter_num])
    r.font.size  = Pt(14)
    r.font.bold  = True
    r.font.color.rgb = DARK_BLUE
    label.paragraph_format.space_after = Pt(4)

    add_horizontal_rule(doc, color="1F497D", thickness="12")

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = subtitle.add_run(CHAPTER_SUBTITLES[chapter_num])
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


def _add_inline_formatting(paragraph, text):
    parts = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*"):
            paragraph.add_run(part[1:-1]).italic = True
        else:
            paragraph.add_run(part)


def parse_chapter_content(doc, content):
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue

        if re.match(r"^#{1,2} ", line) and not line.startswith("###"):
            level = 2 if line.startswith("## ") else 2
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
            _add_inline_formatting(p, text)
            p.paragraph_format.left_indent  = Inches(0.3)
            p.paragraph_format.space_after  = Pt(3)
            i += 1

        elif re.match(r"^\d+\. ", line):
            text = re.sub(r"^\d+\. ", "", line).strip()
            p    = doc.add_paragraph(style="List Number")
            _add_inline_formatting(p, text)
            p.paragraph_format.left_indent  = Inches(0.3)
            p.paragraph_format.space_after  = Pt(3)
            i += 1

        else:
            para_lines = []
            while i < len(lines):
                l = lines[i].rstrip()
                if (not l.strip()
                        or re.match(r"^#{1,3} ", l)
                        or re.match(r"^[\-\*] ", l)
                        or re.match(r"^\d+\. ", l)):
                    break
                para_lines.append(l)
                i += 1
            text = " ".join(para_lines).strip()
            if text:
                p = doc.add_paragraph()
                _add_inline_formatting(p, text)
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
    pf.alignment     = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing  = Pt(24)
    pf.space_after   = Pt(0)


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

    add_horizontal_rule(doc, color="1F497D", thickness="8")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run(
        f"A {level_label} Research Project Submitted in Partial Fulfilment\n"
        "of the Requirements for the Award of a Degree"
    )
    r2.font.size   = Pt(12)
    r2.font.italic = True
    r2.font.color.rgb = GREY
    sub.paragraph_format.space_after = Pt(28)

    for label, value in [
        ("Submitted by:",  "_______________________________"),
        ("Student ID:",    "_______________________________"),
        ("Supervisor:",    "_______________________________"),
        ("Institution:",   "_______________________________"),
        ("Department:",    "_______________________________"),
        ("Date:",          "_______________________________"),
    ]:
        row = doc.add_paragraph()
        row.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rb = row.add_run(f"{label}  ")
        rb.bold      = True
        rb.font.size = Pt(11)
        row.add_run(value).font.size = Pt(11)
        row.paragraph_format.space_after = Pt(6)


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
        add_horizontal_rule(doc, color="1F497D", thickness="8")
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        # Render the section body (strip stray ## heading lines already consumed)
        parse_chapter_content(doc, body)


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


def build_chapter_page(doc, chapter_num, chapter_content):
    add_page_break(doc)
    add_chapter_header(doc, chapter_num)
    parse_chapter_content(doc, chapter_content)


def build_document(topic: str, research_level: str,
                   front_matter: str, chapters: dict,
                   output_dir: str) -> str:
    doc = Document()
    set_document_defaults(doc)

    build_title_page(doc, topic, research_level)
    build_front_matter_page(doc, front_matter)
    build_toc_page(doc, research_level)
    build_abbreviations_page(doc)
    for num in range(1, 6):
        build_chapter_page(doc, num, chapters[num])

    safe     = re.sub(r"[^\w\s-]", "", topic).strip().replace(" ", "_")[:50]
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
    kwargs = dict(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": prompt}],
    )
    if use_thinking:
        kwargs["thinking"] = {"type": "adaptive"}

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
    text = _stream_content(client, system, prompt, model, 3000)
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
