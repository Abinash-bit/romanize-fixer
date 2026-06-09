"""
suggest.py
----------
Offline, dependency-free romanization-correction engine.

It LEARNS character-level rewrite rules from the universal dictionary
(AI Version -> PlanetRead Version) and applies them to *unseen* words so the
tool can catch misspellings that are not literally in the dictionary.

Nothing here calls the internet or any API. Pure pattern learning.

Public API
==========
    engine = SuggestEngine(pairs)          # pairs: list[(wrong, right)]
    result = engine.suggest(word)          # -> Suggestion | None

A Suggestion has: original, correction, confidence (0..1), reason, source.
"""

from __future__ import annotations
import difflib
import logging
import time
from collections import defaultdict, Counter
from dataclasses import dataclass

log = logging.getLogger("romanize.engine")

BOUNDARY = "#"          # marks start/end of a word so rules can be position-aware
CTX = 2                 # how many chars of left/right context a rule remembers

# Ultra-common correct words the engine should never try to "fix" (cuts review
# noise). Verified NOT to be correction targets in the dictionary. Edit freely.
# Any word that later becomes a real correction target is auto-skipped anyway,
# because exact-dictionary words are removed from this set at load time.
STOPWORDS = {
    "hain", "hai", "nahi", "kya", "aur", "mein", "hum", "tum", "yeh", "kar",
    "raha", "rahe", "rahi", "gaya", "gayi", "liye", "sab", "bhi", "toh", "jab",
    "tab", "kab", "koi", "kuch", "isse", "uska", "mera", "tera", "apna",
}


# ──────────────────────────────────────────────────────────────
# small helpers
# ──────────────────────────────────────────────────────────────
def _aug(w: str) -> str:
    return BOUNDARY + w + BOUNDARY


def _strip(w: str) -> str:
    return w.replace(BOUNDARY, "")


def levenshtein(a: str, b: str, max_dist: int | None = None) -> int:
    """Edit distance. If max_dist is given, bail out early (returning max_dist+1)
    as soon as no cell in a row is within the cap — much faster when we only care
    about 'is this close?' rather than the exact distance."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    if max_dist is not None and abs(len(a) - len(b)) > max_dist:
        return max_dist + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        row_min = i
        for j, cb in enumerate(b, 1):
            v = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb))
            cur.append(v)
            if v < row_min:
                row_min = v
        if max_dist is not None and row_min > max_dist:
            return max_dist + 1
        prev = cur
    return prev[-1]


@dataclass
class Suggestion:
    original: str
    correction: str
    confidence: float
    reason: str
    source: str            # "fuzzy" | "rules"


# ──────────────────────────────────────────────────────────────
# the engine
# ──────────────────────────────────────────────────────────────
class SuggestEngine:
    def __init__(self, pairs, ctx: int = CTX):
        t0 = time.perf_counter()
        self.ctx = ctx
        # normalise to lowercase, keep first variant of each side for learning
        self.pairs = []
        for a, b in pairs:
            a1 = a.split("/")[0].strip().lower()
            b1 = b.split("/")[0].strip().lower()
            if a1 and b1:
                self.pairs.append((a1, b1))

        self.wrong_set = {a for a, _ in self.pairs}          # known AI-wrong words
        self.right_set = {b for _, b in self.pairs}          # known correct words
        # never "fix" a stopword — but if it IS a real correction target, drop it
        self.stopwords = {w for w in STOPWORDS if w not in self.wrong_set}
        # bucket wrong words by length so fuzzy lookups only compare similar lengths
        self._wrong_by_len = defaultdict(list)
        for w in self.wrong_set:
            self._wrong_by_len[len(w)].append(w)
        # exact map (all variants) used to skip words we already handle exactly
        self.exact = {}
        for a, b in pairs:
            for v in a.split("/"):
                v = v.strip().lower()
                if v:
                    self.exact[v] = b
        log.info("Engine: %d pairs, %d wrong-words, %d exact variants",
                 len(self.pairs), len(self.wrong_set), len(self.exact))

        t1 = time.perf_counter()
        self.rules = self._mine_rules()
        self._build_index()
        t2 = time.perf_counter()
        self._score_rules()
        t3 = time.perf_counter()
        self._cache = {}
        log.info("Engine ready: %d rules | mine=%.2fs score=%.2fs total=%.2fs",
                 len(self.rules), t2 - t1, t3 - t2, t3 - t0)

    # ----- 1. learn rules --------------------------------------------------
    def _mine_rules(self):
        """rule key = (lctx, frm, rctx) -> Counter(to). Mined from gold alignments."""
        raw = defaultdict(Counter)
        for a, b in self.pairs:
            if a == b:
                continue
            aa, bb = _aug(a), _aug(b)
            sm = difflib.SequenceMatcher(None, aa, bb, autojunk=False)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == "equal":
                    continue
                frm = aa[i1:i2]
                to = bb[j1:j2]
                lctx = aa[max(0, i1 - self.ctx):i1]
                rctx = aa[i2:i2 + self.ctx]
                raw[(lctx, frm, rctx)][to] += 1
        rules = {}
        for key, c in raw.items():
            to, sup = c.most_common(1)[0]
            rules[key] = {"to": to, "support": sup, "obs": sum(c.values()), "prec": 0.0}
        return rules

    def _build_index(self):
        """Index rules by their 'frm' string so we never scan rules that can't match."""
        self._by_frm = defaultdict(list)   # frm -> [(lctx, rctx, info), ...]
        for (lctx, frm, rctx), info in self.rules.items():
            self._by_frm[frm].append((lctx, rctx, info))

    def _find_firings(self, aug):
        """All places a rule matches in `aug`. Uses the frm index for speed.

        Returns list of (start, end, to, support, prec, key).
        """
        firings = []
        n = len(aug)
        for frm, variants in self._by_frm.items():
            F = len(frm)
            start = 0
            while True:
                idx = aug.find(frm, start)
                if idx < 0:
                    break
                start = idx + 1
                for lctx, rctx, info in variants:
                    L, R = len(lctx), len(rctx)
                    if idx - L < 0 or aug[idx - L:idx] != lctx:
                        continue
                    if idx + F + R > n or aug[idx + F:idx + F + R] != rctx:
                        continue
                    firings.append((idx, idx + F, info["to"], info["support"],
                                    info["prec"], (lctx, frm, rctx), info))
        return firings

    # ----- 2. validate each rule on the training data ----------------------
    def _score_rules(self):
        """
        Strict, POSITION-AWARE precision: a rule firing on a training wrong-word
        counts as a hit ONLY if the gold alignment edits that exact span the same
        way. This correctly punishes generic rules (like 'a'->'') that match many
        positions but are only right in a few — the real cause of over-deletion.
        """
        for key in self.rules:
            self.rules[key]["hits"] = 0
            self.rules[key]["fires"] = 0

        for wrong, right in self.pairs:
            aug, augr = _aug(wrong), _aug(right)
            # gold edits as a set of (i1, i2, to) spans on the augmented wrong word
            gold_spans = set()
            sm = difflib.SequenceMatcher(None, aug, augr, autojunk=False)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag != "equal":
                    gold_spans.add((i1, i2, augr[j1:j2]))

            for s, e, to, sup, prec, key, info in self._find_firings(aug):
                info["fires"] += 1
                if (s, e, to) in gold_spans:
                    info["hits"] += 1
        for info in self.rules.values():
            info["prec"] = info["hits"] / info["fires"] if info["fires"] else 0.0

    # ----- 3. apply rules to an arbitrary word -----------------------------
    def _apply_rules(self, word: str):
        aug = _aug(word)
        # firing tuple: (start, end, to, support, prec, key, info)
        firings = self._find_firings(aug)
        # greedy non-overlapping. Prefer LONGER matches first (more specific, e.g.
        # 'aa'->'a' beats two separate 'a'->'' deletions), then precision*support.
        firings.sort(key=lambda f: (f[1] - f[0], f[4] * f[3], f[3]), reverse=True)
        used = [False] * len(aug)
        chosen = []
        for f in firings:
            s, e = f[0], f[1]
            if any(used[s:e]):
                continue
            for i in range(s, e):
                used[i] = True
            chosen.append(f)
        if not chosen:
            return None
        chosen.sort(key=lambda f: f[0])
        out, pos = [], 0
        for s, e, to, *_ in chosen:
            out.append(aug[pos:s])
            out.append(to)
            pos = e
        out.append(aug[pos:])
        cand = _strip("".join(out))
        return cand, chosen

    # ----- 4. public: suggest a correction for one word --------------------
    def suggest(self, word: str):
        w = word.lower()
        if w in self._cache:
            return self._cache[w]
        result = self._suggest(w)
        self._cache[w] = result
        return result

    def _suggest(self, w: str):
        if not w or not w.isalpha():
            return None
        if w in self.exact:                 # handled by the exact dictionary already
            return None
        if w in self.right_set:             # already a known-correct word — leave it
            return None
        if w in self.stopwords:             # ultra-common correct word — leave it
            return None

        best = None

        # --- source A: fuzzy distance to a known WRONG word -----------------
        # if w is 1-2 edits from a word we KNOW is wrong, that's strong evidence
        nearest, ndist = None, 99
        for dl in (0, -1, 1, -2, 2):            # only compare similar-length words
            for k in self._wrong_by_len.get(len(w) + dl, ()):
                d = levenshtein(w, k, max_dist=2)   # we only care about d <= 2
                if d < ndist:
                    nearest, ndist = k, d
            if ndist == 1:                       # can't do better than 1 here
                break

        # --- source B: rule-based generation -------------------------------
        ruled = self._apply_rules(w)

        if ruled:
            cand, chosen = ruled
            if cand != w:
                precs = [c[4] for c in chosen]   # c = (s,e,to,support,prec,key,info)
                sups = [c[3] for c in chosen]
                conf = min(precs)                         # weakest link (strict precision)
                # thinly-attested rules are risky -> cap below the auto-fix bar
                if min(sups) < 3:
                    conf = min(conf, 0.70)
                # stacking many edits at once is risky -> cap unless corroborated
                if len(chosen) >= 3 and cand not in self.right_set:
                    conf = min(conf, 0.70)
                # the produced word is itself a known-correct word -> very strong
                if cand in self.right_set:
                    conf = min(1.0, conf + 0.10)
                # near a known wrong word as well -> corroboration
                if nearest is not None and ndist == 1:
                    conf = min(1.0, conf + 0.05)
                # short words are ambiguous -> penalise
                if len(w) <= 3:
                    conf -= 0.30
                elif len(w) == 4:
                    conf -= 0.12
                conf = max(0.0, min(1.0, conf))
                rule_desc = ", ".join(
                    f"{_strip(c[5][1]) or '∅'}→{_strip(c[2]) or '∅'}" for c in chosen
                )
                best = Suggestion(w, cand, conf, f"rule(s): {rule_desc}", "rules")

        # NOTE: there used to be a "1 edit from a known error" fuzzy fallback here
        # that copied a *different* word's correction wholesale. It produced badly
        # wrong guesses (kaali->dali, waahab->sahab, Jhoola->bhula) because a word
        # one letter away is a DIFFERENT word with a different fix. Removed on
        # purpose — only learned rewrite rules generate suggestions now.
        return best
