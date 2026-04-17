"""
SpendBot — Rule-based conversational assistant for SpendSight.

No external AI/LLM APIs.  Uses only: datetime, re, and standard library.
"""

import re
from datetime import date, datetime, timedelta
from difflib import get_close_matches, SequenceMatcher

# ---------------------------------------------------------------------------
# Default subcategory corpus (mirrors DEFAULT_CATEGORIES in app.py)
# ---------------------------------------------------------------------------
DEFAULT_CATEGORIES = {
    "Groceries":     ["Milk", "Dal", "Rice", "Sugar", "Atta", "Oil", "Soap",
                      "Shampoo", "Vegetables", "Fruits", "Eggs", "Pulses",
                      "Ghee", "Butter", "Tea"],
    "Fast Food":     ["McDonald's", "Wada Pav", "Pani Puri", "Samosa", "Pizza",
                      "Chai", "Biryani", "Thali", "Dosa", "Idli", "Burger", "Noodles"],
    "Fuel":          ["Petrol", "Diesel", "CNG"],
    "Utilities":     ["Electricity", "Water", "Gas Cylinder", "Internet",
                      "Mobile Recharge", "DTH"],
    "Entertainment": ["Netflix", "Amazon Prime", "Hotstar", "Movies", "Games",
                      "Events", "Spotify"],
    "Health":        ["Medicine", "Doctor", "Gym", "Supplements", "Lab Test", "Dental"],
    "Transport":     ["Ola/Uber", "Local Train", "Bus", "Auto", "Metro", "Parking"],
    "Shopping":      ["Clothing", "Electronics", "Household", "Personal Care",
                      "Accessories", "Books"],
    "EMI / Finance": ["Card EMI", "Loan EMI", "Insurance Premium", "SIP", "Rent"],
    "Other":         ["Miscellaneous", "Gifts", "Donations", "Fees"],
}

# ---------------------------------------------------------------------------
# Month-name → number mapping (full + 3-letter abbreviations)
# ---------------------------------------------------------------------------
MONTH_NAMES = {
    "january": 1,  "jan": 1,
    "february": 2, "feb": 2,
    "march": 3,    "mar": 3,
    "april": 4,    "apr": 4,
    "may": 5,
    "june": 6,     "jun": 6,
    "july": 7,     "jul": 7,
    "august": 8,   "aug": 8,
    "september": 9,"sep": 9,  "sept": 9,
    "october": 10, "oct": 10,
    "november": 11,"nov": 11,
    "december": 12,"dec": 12,
}

MONTH_DISPLAY = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------
# "March 2026", "Jan 2025", etc.
_RE_MONTH_YEAR = re.compile(
    r"\b(january|february|march|april|may|june|july|august|september|october|november|december"
    r"|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)"
    r"(?:\s+(\d{4}))?\b",
    re.IGNORECASE,
)

_RE_THIS_MONTH  = re.compile(r"\bthis\s+month\b", re.IGNORECASE)
_RE_LAST_MONTH  = re.compile(r"\blast\s+month\b", re.IGNORECASE)
_RE_YTD         = re.compile(r"\b(year\s+to\s+date|ytd)\b", re.IGNORECASE)
_RE_NEXT_MONTH  = re.compile(r"\bnext\s+month\b", re.IGNORECASE)

# Enhancement 2: Date Shortcuts
_RE_LAST_WEEK   = re.compile(r"\blast\s+week\b", re.IGNORECASE)
_RE_PAST_7_DAYS = re.compile(r"\b(past|last)\s+7\s+days\b", re.IGNORECASE)
_RE_YESTERDAY   = re.compile(r"\byesterday\b", re.IGNORECASE)
_RE_LAST_30_DAYS= re.compile(r"\b(past|last)\s+30\s+days\b", re.IGNORECASE)
_RE_THIS_WEEK   = re.compile(r"\bthis\s+week\b", re.IGNORECASE)
_RE_LAST_YEAR   = re.compile(r"\blast\s+year\b", re.IGNORECASE)
_RE_LAST_X_MONTHS = re.compile(r"\b(last|past|previous)\s+(\d+)\s+months?\b", re.IGNORECASE)

# Enhancement 5: Insights
_RE_INSIGHTS    = re.compile(r"\b(most|top|biggest|largest|insights|summary)\b", re.IGNORECASE)
_RE_TOP_N       = re.compile(r"\btop\s+(\d+)\b", re.IGNORECASE)

# Enhancement 2 & 3: Rank & Continuation
_RE_RANK        = re.compile(r"\b(first|1st|second|2nd|third|3rd|fourth|4th|fifth|5th|next|then|after that|highest|lowest)\b", re.IGNORECASE)
_RE_RANK_NUM    = {
    "first": 1, "1st": 1, "highest": 1,
    "second": 2, "2nd": 2, "two": 2,
    "third": 3, "3rd": 3, "three": 3,
    "fourth": 4, "4th": 4, "four": 4,
    "fifth": 5, "5th": 5, "five": 5
}
_RE_CONTINUATION = re.compile(r"^(and\s+)?(what\s+about|how\s+about|same\s+for|now\s+for|and\s+this|and\s+last|what\s+about|and)\s+", re.IGNORECASE)

# Intent: total / overall
_RE_TOTAL_INTENT = re.compile(
    r"\b(total|overall|all|everything|sum|how much (did i |have i )?spend(t)?|"
    r"expense[s]?|spending)\b",
    re.IGNORECASE,
)

# Intent: EMI / Loan
_RE_EMI_TOTAL   = re.compile(r"\b((emi\w*|loan\w*)\s+total|total\s+(emi\w*|loan\w*)|how much total\s+(emi\w*|loan\w*))\b", re.IGNORECASE)
_RE_EMI_END     = re.compile(r"\b((emi\w*)\s+(is\s+)?coming\s+to\s+an?\s+end|(emi\w*)\s+coming\s+to\s+end|(emi\w*) ending soon|(emi\w*) about to end|which\s+(emi\w*)\s+ends?)\b", re.IGNORECASE)
_RE_EMI_WHEN    = re.compile(r"\bwhen (does|will) (.+?) (end|finish|complete)\b", re.IGNORECASE)
_RE_EMI_NEXT    = re.compile(r"\b(next\s+emi\w*|(emi\w*)\s+due|upcoming\s+emi\w*)\b", re.IGNORECASE)
_RE_EMI_COUNT   = re.compile(r"\b(how many|how may|number of|count)\b.*\b(emi\w*|loan\w*|recurring|fixed)\b|\b(emi\w*|loan\w*|recurring|fixed)\b.*\b(how many|how may|number of|count)\b", re.IGNORECASE)
_RE_EMI_LATEST  = re.compile(r"\b(latest|newest|most recent)\b.*\b(recurring|emi\w*|loan\w*|fixed)\b|\b(recurring|emi\w*|loan\w*|fixed)\b.*\b(latest|newest|most recent)\b", re.IGNORECASE)

# Gibberish guard: if the message has fewer than 2 word-like tokens, treat as gibberish
_RE_WORD = re.compile(r"[a-zA-Z]{2,}")

# ---------------------------------------------------------------------------
# Fuzzy / normalised matching helpers
# ---------------------------------------------------------------------------

# Simple suffix rules for plural → singular (applied to normalised lowercase)
_PLURAL_RULES = [
    (r"groceries$", "grocery"), (r"berries$", "berry"),
    (r"ies$",       "y"),       (r"ves$",     "f"),
    (r"ses$",       "s"),       (r"s$",       ""),
]

# Alias map: normalised user input → normalised canonical key
_COMMON_ALIASES: dict[str, str] = {
    "panipuri":      "pani puri",
    "wadapav":       "wada pav",
    "fastfood":      "fast food",
    "fual":          "fuel",
    "fuell":         "fuel",
    "mcdonalds":     "mcdonald's",
    "mobilerecharge":"mobile recharge",
    "gascylinder":   "gas cylinder",
    "labtest":       "lab test",
    "amazomprime":   "amazon prime",
    "amazone":       "amazon prime",
    "olauber":       "ola/uber",
    "cardemi":       "card EMI",
    "loanemi":       "loan EMI",
}

# Threshold constants
_AUTO_MATCH_THRESHOLD  = 0.82   # ≥ this → auto-match with a correction note
_SUGGEST_THRESHOLD     = 0.70   # ≥ this (but < AUTO) → ask "Did you mean?"


def _normalize(text: str) -> str:
    """Lowercase; remove spaces, hyphens, slashes."""
    return re.sub(r"[\s\-/]+", "", text.lower())


def _singularize(word: str) -> str:
    """Best-effort plural → singular on a *normalised* (space-free) lowercase word."""
    for pattern, repl in _PLURAL_RULES:
        if re.search(pattern, word):
            candidate = re.sub(pattern, repl, word)
            if len(candidate) >= 2:          # don't collapse to empty
                return candidate
    return word


def detect_item_with_suggestions(
    user_message: str,
    expenses: list,
    default_categories: dict,
) -> tuple:
    """
    Fuzzy / normalised item detection — no external APIs.

    Priority order
    --------------
    1. Exact match (case-insensitive)
    2. Normalised match  (strip spaces/hyphens + alias map)
    3. Singular-form match
    4. Fuzzy via difflib  (SequenceMatcher score ≥ 0.85 → auto; 0.70–0.85 → suggest)
    5. Substring match

    Parameters
    ----------
    user_message       : raw user text
    expenses           : list of expense dicts (to pick up custom subcategories)
    default_categories : {category: [subcategory, …]}

    Returns
    -------
    matched_item    : str | None          canonical name, or None
    match_type      : str                 'exact'|'normalized'|'fuzzy'|'substring'|'suggestion'|'none'
    confidence      : float               0.0–1.0
    suggestion_text : str                 human-readable note / question
    alternatives    : list[str]           up to 8 candidate names for "none" / "suggestion" cases
    """
    # Build combined term dict: lower → canonical
    all_terms: dict[str, str] = {}
    for cat, subs in default_categories.items():
        all_terms[cat.lower()] = cat
        for sub in subs:
            all_terms[sub.lower()] = sub
    for exp in expenses:
        for key in ("category", "subcategory"):
            val = (exp.get(key) or "").strip()
            if val:
                all_terms[val.lower()] = val

    # Extract candidate token (word/phrase after "on / about / spent on")
    m = re.search(
        r"\b(?:on|about|spent\s+on)\s+([a-z][a-z\s]{0,25}?)"
        r"(?:\s+(?:this|last|in|for|january|february|march|april|may|june|july"
        r"|august|september|october|november|december|\d{4})|[?,.!]|$)",
        user_message.lower(),
    )
    token = m.group(1).strip() if m else user_message.lower().strip()
    if not token:
        alts = sorted({v for v in all_terms.values()}, key=str.lower)[:8]
        return None, "none", 0.0, "", alts

    all_canonicals = sorted({v for v in all_terms.values()}, key=str.lower)

    # ── 1. Exact match ────────────────────────────────────────────────────
    if token in all_terms:
        return all_terms[token], "exact", 1.0, "", []

    # ── 2. Normalised match + alias ───────────────────────────────────────
    norm_token = _normalize(token)
    resolved   = _COMMON_ALIASES.get(norm_token, norm_token)
    norm_map: dict[str, str] = {_normalize(k): v for k, v in all_terms.items()}

    for candidate in (resolved, norm_token):
        if candidate in norm_map:
            return norm_map[candidate], "normalized", 0.95, f"(you typed '{token}')", []

    # ── 3. Singular-form match ────────────────────────────────────────────
    singular = _singularize(norm_token)
    for nk, canon in norm_map.items():
        if _singularize(nk) == singular or nk == singular:
            return canon, "normalized", 0.92, f"(you typed '{token}')", []

    # ── 4. Fuzzy via difflib ──────────────────────────────────────────────
    norm_keys = list(norm_map.keys())
    close = get_close_matches(norm_token, norm_keys, n=3, cutoff=_SUGGEST_THRESHOLD)
    if close:
        best  = close[0]
        score = SequenceMatcher(None, norm_token, best).ratio()
        canon = norm_map[best]
        if score >= _AUTO_MATCH_THRESHOLD:
            return canon, "fuzzy", score, f"(you typed '{token}')", []
        else:
            alts = [norm_map[c] for c in close[1:3]]
            return (canon, "suggestion", score,
                    f"Did you mean '{canon}'? (from '{token}')", alts)

    # ── 5. Substring match ────────────────────────────────────────────────
    for lower_key, canon in all_terms.items():
        if token in lower_key or (len(token) >= 3 and lower_key.startswith(token)):
            return canon, "substring", 0.75, f"(matched '{canon}' from '{token}')", []

    # ── No match ──────────────────────────────────────────────────────────
    return None, "none", 0.0, f"I don't recognize '{token}'.", all_canonicals[:8]


# ---------------------------------------------------------------------------
# SpendBot
# ---------------------------------------------------------------------------

class SpendBot:
    """
    Rule-based spending assistant.

    Parameters
    ----------
    expenses : list[dict]
        List of expense records with keys: amount, category, subcategory, date
    custom_categories : dict, optional
        Extra {category: [subcategory, …]} from user data, merged with defaults
    currency_symbol : str
        Symbol to use in replies (default "₹")
    """

    def __init__(self, expenses: list, custom_categories: dict = None,
                 recurring_payments: list = None, currency_symbol: str = "₹"):
        self.expenses           = expenses
        self.recurring_payments = recurring_payments or []
        self.currency_symbol    = currency_symbol

        # Build merged category map and flat subcategory list
        self._categories: dict[str, list[str]] = {}
        for cat, subs in DEFAULT_CATEGORIES.items():
            self._categories[cat] = list(subs)
        for cat, subs in (custom_categories or {}).items():
            if cat in self._categories:
                existing = {s.lower() for s in self._categories[cat]}
                self._categories[cat] += [s for s in subs if s.lower() not in existing]
            else:
                self._categories[cat] = list(subs)

        # Flat list of all subcategories (lower-cased) → canonical form
        self._sub_to_canonical: dict[str, str] = {}
        self._cat_to_canonical: dict[str, str] = {}
        for cat, subs in self._categories.items():
            self._cat_to_canonical[cat.lower()] = cat
            for sub in subs:
                self._sub_to_canonical[sub.lower()] = sub

    # ------------------------------------------------------------------
    # Public entry-point
    # ------------------------------------------------------------------

    def reply(self, message: str, context: dict = None) -> dict:
        """
        Process a user message with session context and return a response dict.
        """
        msg = message.strip()
        msg = re.sub(r"\bhow may\b", "how many", msg, flags=re.IGNORECASE)
        msg = re.sub(r"\bemin\b", "emi", msg, flags=re.IGNORECASE)
        msg = re.sub(r"\bemis\b", "emi", msg, flags=re.IGNORECASE)
        today = date.today()
        context = context or {}

        # 0. Handle EMI / Loan Specific Queries
        if _RE_EMI_TOTAL.search(msg):
            return self._handle_emi_total(msg)
        if _RE_EMI_END.search(msg):
            return self._handle_emi_end()
        match_when = _RE_EMI_WHEN.search(msg)
        if match_when:
            return self._handle_emi_when(match_when.group(2))
        if _RE_EMI_NEXT.search(msg):
            return self._handle_emi_next()
        if _RE_EMI_COUNT.search(msg):
            return self._handle_emi_count(msg)
        if _RE_EMI_LATEST.search(msg):
            return self._handle_emi_latest(msg)

        # Guard: empty or gibberish
        words = _RE_WORD.findall(msg)
        if not msg or (len(words) < 2 and not _RE_MONTH_YEAR.search(msg) and self._detect_item(msg) is None):
            return self._clarify("Sorry, I didn't understand. Try 'How much on milk this month?'")

        # 1. Detect item FIRST (Priority: Explicit over Context)
        explicit_item_res = self._detect_item(msg)
        fuzzy_item_name = None
        if not explicit_item_res:
            # Try fuzzy but only extract name, don't return full reply yet
            f_res = detect_item_with_suggestions(msg, self.expenses, self._categories)
            if f_res[0] and f_res[1] not in ("suggestion", "none"):
                fuzzy_item_name = f_res[0]

        # Determine which item we are talking about
        final_item = None
        if explicit_item_res:
            final_item = explicit_item_res[0]
        elif fuzzy_item_name:
            final_item = fuzzy_item_name
        elif _RE_CONTINUATION.search(msg) and context.get("last_queried_item"):
            final_item = context.get("last_queried_item")

        # 2. Detect Timeframe
        tf_res = self._detect_timeframe(msg)
        start, end, label, tf_defaulted = tf_res if tf_res else (None, None, None, True)

        # Timeframe Inheritance: If item found in msg but no timeframe, use session timeframe
        if (explicit_item_res or fuzzy_item_name) and tf_defaulted:
            if context.get("context_start") and context.get("context_end"):
                try:
                    start = date.fromisoformat(context["context_start"])
                    end   = date.fromisoformat(context["context_end"])
                    label = context.get("context_timeframe", "the selected period")
                    tf_defaulted = False
                except: pass

        # 3. Handle Insights (Enhancement 5)
        if _RE_INSIGHTS.search(msg):
            res = self._get_insights(msg, start, end, label, tf_defaulted)
            if res.get("item") in ("insights", "top_spending"):
                matching = [e for e in self.expenses if start <= self._parse_date(e.get("date", "")) <= end]
                by_sub = {}
                for e in matching:
                    sub = e.get("subcategory") or e.get("category")
                    by_sub[sub] = by_sub.get(sub, 0.0) + e["amount"]
                sorted_subs = sorted(by_sub.items(), key=lambda x: x[1], reverse=True)
                res["context_sorted_expenses"] = sorted_subs
                res["context_timeframe"] = label
                res["context_start"] = start.isoformat()
                res["context_end"] = end.isoformat()
            return res

        # 4. Check for Ranked Queries (Enhancement 2)
        rank_info = self._get_rank_info(msg, context)
        if rank_info:
            res = self._ranked_query(rank_info, start, end, label, context)
            if res:
                res["context_start"] = start.isoformat()
                res["context_end"] = end.isoformat()
                return res

        # 5. Handle Multiple Items (Enhancement 3)
        items = self._split_items(msg)
        if items:
            responses = []
            grand_total = 0.0
            found_any = False
            for it in items:
                res = self._detect_item(it)
                if not res:
                    fuz = self._detect_item_fuzzy(it, start, end, label, tf_defaulted)
                    if fuz and fuz.get("match_type") not in ("suggestion", "none"):
                        res = (fuz["item_canonical"], False)
                if res:
                    name, _ = res
                    matching = self._filter_expenses(name, start, end)
                    total = sum(e["amount"] for e in matching)
                    grand_total += total
                    responses.append(f"{name}: {self._fmt(total)}")
                    found_any = True
            if found_any:
                suffix = " (defaulting to this month)" if tf_defaulted else ""
                reply = ", ".join(responses) + f" | Total: {self._fmt(grand_total)} for {label}{suffix}"
                return {
                    "reply": reply, "total": round(grand_total, 2), "item": "multiple", "timeframe": label,
                    "last_queried_item": "multiple", "context_start": start.isoformat(), "context_end": end.isoformat(), "context_timeframe": label
                }

        # 6. Execute Item Query
        if final_item:
            # Timeframe Clarification (if still defaulted and no TF in context)
            tf_keywords = ["this", "last", "past", "yesterday", "month", "year", "week", "ytd"] + list(MONTH_NAMES.keys())
            has_tf_explicit = any(k in msg.lower() for k in tf_keywords)
            
            if tf_defaulted and not has_tf_explicit and not context.get("context_start"):
                return {
                    "reply": f"Which timeframe for '{final_item}'? Reply: 'this month', 'last month', or 'all time'",
                    "needs_timeframe": True, "pending_item": final_item, "total": None, "item": final_item, "timeframe": None,
                    "last_queried_item": final_item
                }

            res = self._item_query(final_item, start, end, label, tf_defaulted)
            if res.get("total") == 0:
                res = self._get_smart_empty_response(final_item, start, end, label, tf_defaulted)
            
            res["last_queried_item"] = final_item
            res["context_start"] = start.isoformat()
            res["context_end"] = end.isoformat()
            res["context_timeframe"] = label
            return res

        # 7. Fallback to total query
        if bool(_RE_TOTAL_INTENT.search(msg)) or not tf_defaulted:
            res = self._total_query(start, end, label, tf_defaulted)
            res["context_start"] = start.isoformat()
            res["context_end"] = end.isoformat()
            res["context_timeframe"] = label
            return res

        return self._clarify("Sorry, I didn't quite catch that. Try 'How much on milk this month?'")

    def _split_items(self, msg: str) -> list:
        """Split message into potential multiple items using 'and', ',', '&'."""
        # Clean msg of timeframe/intent words to avoid them being treated as items
        clean = _RE_TOTAL_INTENT.sub("", msg)
        clean = _RE_THIS_MONTH.sub("", clean)
        clean = _RE_LAST_MONTH.sub("", clean)
        # ... could add more but let's keep it simple
        
        parts = re.split(r",| and | & ", clean, flags=re.IGNORECASE)
        parts = [p.strip() for p in parts if p.strip()]
        return parts if len(parts) > 1 else []

    def _get_insights(self, msg: str, start: date, end: date, label: str, defaulted: bool) -> dict:
        """Handle Enhancement 5: Top spending and insights."""
        suffix = " (defaulting to this month)" if defaulted else ""
        
        # Filter expenses in timeframe
        matching = [e for e in self.expenses if start <= self._parse_date(e.get("date", "")) <= end]
        
        if not matching:
            return self._clarify(f"No expenses found for {label}{suffix} to provide insights.")

        if "insights" in msg.lower() or "summary" in msg.lower():
            total = sum(e["amount"] for e in matching)
            # Group by subcategory
            by_sub = {}
            for e in matching:
                sub = e.get("subcategory") or e.get("category")
                by_sub[sub] = by_sub.get(sub, 0.0) + e["amount"]
            top_item = max(by_sub, key=by_sub.get)
            freq = {}
            for e in matching:
                sub = e.get("subcategory") or e.get("category")
                freq[sub] = freq.get(sub, 0) + 1
            most_freq = max(freq, key=freq.get)
            
            reply = (f"📊 {label}: Total {self._fmt(total)} | "
                     f"Top: {top_item} ({self._fmt(by_sub[top_item])}) | "
                     f"Most frequent: {most_freq} ({freq[most_freq]}x)")
            return {"reply": reply, "total": round(total, 2), "item": "insights", "timeframe": label}

        # Top N logic
        limit = 1
        m_top = _RE_TOP_N.search(msg.lower())
        if m_top:
            limit = int(m_top.group(1))
        elif any(k in msg.lower() for k in ["most", "biggest", "largest"]):
            limit = 1

        # Sort all individual expenses or grouped? Requirement says "what did I spend most on"
        # Usually implies subcategory grouping or category grouping.
        # "You spent ₹245.00 on Groceries in March 2026 (includes Milk, Vegetables, Bread)"
        # Let's group by subcategory for "most"
        by_sub = {}
        for e in matching:
            sub = e.get("subcategory") or e.get("category")
            by_sub[sub] = by_sub.get(sub, 0.0) + e["amount"]
        
        sorted_subs = sorted(by_sub.items(), key=lambda x: x[1], reverse=True)
        top_list = sorted_subs[:limit]
        
        if limit == 1:
            item, amt = top_list[0]
            reply = f"Your highest expense {label} was {item}: {self._fmt(amt)}{suffix}"
        else:
            lines = [f"{i+1}. {item}: {self._fmt(amt)}" for i, (item, amt) in enumerate(top_list)]
            reply = f"Top {len(top_list)} expenses for {label}: " + " | ".join(lines) + suffix
            
        return {"reply": reply, "total": None, "item": "top_spending", "timeframe": label}

    # ------------------------------------------------------------------
    # Timeframe detection
    # ------------------------------------------------------------------

    def _detect_timeframe(self, msg: str):
        """
        Returns (start_date, end_date, label, defaulted) or None for invalid.

        defaulted=True when we silently fell back to this month.
        """
        today = date.today()

        # "next month" → invalid
        if _RE_NEXT_MONTH.search(msg):
            return None

        # Enhancement: Last X Months (Complete months only, exclude current partial month)
        m_x = _RE_LAST_X_MONTHS.search(msg)
        if m_x:
            X = int(m_x.group(2))
            if X <= 0: return *self._month_range(today.year, today.month), f"{MONTH_DISPLAY[today.month]} {today.year}", True
            
            # End = last day of previous month
            em, ey = today.month - 1, today.year
            if em == 0: em, ey = 12, ey - 1
            _, end = self._month_range(ey, em)
            
            # Start = 1st day of month X months ago
            # e.g. X=1, sm=em. X=3, sm=em-2.
            sm, sy = em - X + 1, ey
            while sm <= 0:
                sm += 12
                sy -= 1
            start = date(sy, sm, 1)
            
            label_pref = m_x.group(1).capitalize()
            label = f"{label_pref} {X} Month{'s' if X>1 else ''} ({MONTH_DISPLAY[sm][:3]}-{MONTH_DISPLAY[em][:3]} {ey})"
            if sy != ey:
                label = f"{label_pref} {X} Month{'s' if X>1 else ''} ({MONTH_DISPLAY[sm][:3]} {sy}-{MONTH_DISPLAY[em][:3]} {ey})"
            
            return start, end, label, False

        # Enhancement 2: Relative Timeframes
        if _RE_YESTERDAY.search(msg):
            d = today - timedelta(days=1)
            return d, d, "yesterday", False
        if _RE_LAST_WEEK.search(msg):
            return today - timedelta(days=7), today - timedelta(days=1), "last week", False
        if _RE_PAST_7_DAYS.search(msg):
            return today - timedelta(days=6), today, "the past 7 days", False
        if _RE_LAST_30_DAYS.search(msg):
            return today - timedelta(days=29), today, "the past 30 days", False
        if _RE_THIS_WEEK.search(msg):
            # Monday to Sunday of current week
            start = today - timedelta(days=today.weekday())
            end = start + timedelta(days=6)
            return start, end, "this week", False
        if _RE_LAST_YEAR.search(msg):
            return date(today.year - 1, 1, 1), date(today.year - 1, 12, 31), str(today.year - 1), False

        # "this month"
        if _RE_THIS_MONTH.search(msg):
            return *self._month_range(today.year, today.month), \
                   f"{MONTH_DISPLAY[today.month]} {today.year}", False

        # "last month"
        if _RE_LAST_MONTH.search(msg):
            m, y = today.month - 1, today.year
            if m == 0:
                m, y = 12, y - 1
            return *self._month_range(y, m), f"{MONTH_DISPLAY[m]} {y}", False

        # Year-to-date
        if _RE_YTD.search(msg):
            start = date(today.year, 1, 1)
            label = f"Jan–{MONTH_DISPLAY[today.month]} {today.year}"
            return start, today, label, False

        # Named month (+ optional year)
        m = _RE_MONTH_YEAR.search(msg)
        if m:
            month_num = MONTH_NAMES[m.group(1).lower()]
            year      = int(m.group(2)) if m.group(2) else today.year

            # Enhancement 2: If no year specified and month is in future, assume last year
            if not m.group(2) and date(year, month_num, 1) > today.replace(day=1):
                year -= 1

            # Reject future months (if year was specified or even after adjustment)
            target = date(year, month_num, 1)
            if target > today.replace(day=1):
                return None

            label = f"{MONTH_DISPLAY[month_num]} {year}"
            return *self._month_range(year, month_num), label, False

        # Default → this month
        return (*self._month_range(today.year, today.month),
                f"{MONTH_DISPLAY[today.month]} {today.year}", True)

    @staticmethod
    def _month_range(year: int, month: int):
        from calendar import monthrange as _mr
        last_day = _mr(year, month)[1]
        return date(year, month, 1), date(year, month, last_day)

    # ------------------------------------------------------------------
    # Item detection
    # ------------------------------------------------------------------

    def _detect_item(self, msg: str):
        """
        Scan message for known subcategories / categories.
        Returns (canonical_name, is_category_level) or None.

        Strategy:
        - Longest-match wins (avoids "burger" matching inside "burger king").
        - Case-insensitive.
        - Matches both subcategories and top-level categories.
        """
        msg_lower = msg.lower()

        # Build a sorted list (longest first) so multi-word items take precedence
        candidates = []
        for sub_lower, canonical in self._sub_to_canonical.items():
            candidates.append((sub_lower, canonical, False))
        for cat_lower, canonical in self._cat_to_canonical.items():
            candidates.append((cat_lower, canonical, True))

        candidates.sort(key=lambda x: len(x[0]), reverse=True)

        for term, canonical, is_cat in candidates:
            # Word-boundary aware match: term must appear as a whole word / phrase
            pattern = r"\b" + re.escape(term) + r"\b"
            if re.search(pattern, msg_lower):
                return canonical, is_cat

        return None

    def _detect_item_fuzzy(
        self,
        msg: str,
        start: date,
        end: date,
        tf_label: str,
        tf_defaulted: bool,
    ) -> dict | None:
        """
        Call the module-level detect_item_with_suggestions() and convert the
        result into a full reply dict, or return None if nothing was found.

        match_type routing
        ------------------
        exact / normalized / fuzzy / substring  → auto-match with an optional note
        suggestion (score 0.70–0.85)            → ask "Did you mean …?" and return
                                                  extra fields so app.py can store
                                                  the pending suggestion in session
        none                                    → return None (caller decides)
        """
        matched, match_type, confidence, suggestion_text, alts = \
            detect_item_with_suggestions(msg, self.expenses, self._categories)

        if matched is None or match_type == "none":
            return None

        if match_type == "suggestion":
            # Medium confidence — ask the user to confirm before running the query
            r = self._clarify(
                f"Did you mean '{matched}'? "
                f"Reply 'yes' to confirm, or retype the item name."
            )
            # Extra fields consumed by /api/chat to persist the pending state
            r["match_type"]    = "suggestion"
            r["suggested_item"] = matched
            r["_start"]        = start.isoformat()
            r["_end"]          = end.isoformat()
            r["_tf_label"]     = tf_label
            r["_tf_defaulted"] = tf_defaulted
            return r

        # Auto-match: run the normal item query, prepend a correction note if needed
        r    = self._item_query(matched, start, end, tf_label, tf_defaulted)
        note = suggestion_text   # e.g. "(you typed 'vegitable')"
        if note:
            r["reply"] = f"Showing results for '{matched}' {note}. {r['reply']}"
        r["match_type"] = match_type
        r["item_canonical"] = matched
        return r

    def _detect_unknown_item(self, msg: str) -> str | None:
        """
        Look for a noun-like word after common question prepositions
        ('on', 'for', 'about', 'spent on') that isn't in our vocabulary.
        Returns the unknown token, or None.
        """
        # e.g. "what about coffee", "spent on coffee", "how much on coffee"
        pattern = (
            r"\b(?:on|about|spent on)\s+([a-z][a-z\s]{1,20}?)"
            r"(?:\s+(?:this|last|in|for|january|february|march|april|may|june|july"
            r"|august|september|october|november|december|\d{4})|[?,.]|$)"
        )
        m = re.search(pattern, msg.lower())
        if m:
            candidate = m.group(1).strip()
            # Exclude pure timeframe words and month names
            timeframe_words = {
                "month", "year", "date", "today", "week", "yesterday",
            } | set(MONTH_NAMES.keys())
            if candidate not in timeframe_words:
                return candidate
        return None

    def _available_items(self) -> list[str]:
        """Flat sorted list of all known subcategory names."""
        seen = set()
        items = []
        for subs in self._categories.values():
            for s in subs:
                sl = s.lower()
                if sl not in seen:
                    seen.add(sl)
                    items.append(s)
        return sorted(items, key=str.lower)

    # ------------------------------------------------------------------
    # Query handlers
    # ------------------------------------------------------------------

    def _item_query(self, item_name: str, start: date, end: date,
                    tf_label: str, defaulted: bool) -> dict:
        """Return spending on a specific item within the date range."""
        matching = self._filter_expenses(item_name, start, end)
        total    = sum(e["amount"] for e in matching)

        suffix = " (defaulting to this month)" if defaulted else ""

        if total == 0:
            reply = (
                f"I couldn't find any expenses for {item_name} in {tf_label}."
            )
            return {"reply": reply, "total": 0.0,
                    "item": item_name.lower(), "timeframe": tf_label}

        # Enhancement 4: Category-Level Queries
        is_cat = item_name.lower() in [c.lower() for c in self._categories.keys()]
        extra = ""
        if is_cat:
            subs = sorted({e.get("subcategory") for e in matching if e.get("subcategory")})
            if subs:
                extra = f" (includes {', '.join(subs)})"

        reply = (
            f"You spent {self._fmt(total)} on {item_name} "
            f"in {tf_label}{suffix}{extra}."
        )
        return {"reply": reply, "total": round(total, 2),
                "item": item_name.lower(), "timeframe": tf_label}

    def _total_query(self, start: date, end: date,
                     tf_label: str, defaulted: bool) -> dict:
        """Return total of all expenses within the date range."""
        matching = [
            e for e in self.expenses
            if start <= self._parse_date(e.get("date", "")) <= end
        ]
        total = sum(e["amount"] for e in matching)

        suffix = " (defaulting to this month)" if defaulted else ""

        if total == 0:
            reply = f"No expenses recorded for {tf_label}{suffix}."
            return {"reply": reply, "total": 0.0,
                    "item": None, "timeframe": tf_label}

        reply = (
            f"Your total expenses for {tf_label} were "
            f"{self._fmt(total)}{suffix}."
        )
        return {"reply": reply, "total": round(total, 2),
                "item": None, "timeframe": tf_label}

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _filter_expenses(self, item_name: str, start: date, end: date) -> list:
        """
        Return expenses whose subcategory OR category contains item_name
        (case-insensitive, partial match for multi-word items).
        """
        item_lower = item_name.lower()
        results    = []
        for e in self.expenses:
            sub = (e.get("subcategory") or "").lower()
            cat = (e.get("category")    or "").lower()
            if item_lower in sub or item_lower in cat or sub in item_lower or cat in item_lower:
                try:
                    d = self._parse_date(e.get("date", ""))
                    if start <= d <= end:
                        results.append(e)
                except ValueError:
                    pass
        return results

    @staticmethod
    def _parse_date(date_str: str) -> date:
        """Parse ISO date string → date object. Raises ValueError on failure."""
        if not date_str:
            raise ValueError("empty date")
        return datetime.strptime(date_str, "%Y-%m-%d").date()

    def _fmt(self, amount: float) -> str:
        """Indian lakh-style currency formatting using the instance symbol."""
        symbol = self.currency_symbol
        v = int(round(amount))
        if v < 0:
            return f"-{self._fmt(-amount)}"
        s = str(v)
        if len(s) <= 3:
            return f"{symbol}{s}"
        last3 = s[-3:]
        rest  = s[:-3]
        groups: list[str] = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        return f"{symbol}{','.join(groups)},{last3}"

    @staticmethod
    def _clarify(msg: str) -> dict:
        return {"reply": msg, "total": None, "item": None, "timeframe": None}

    # ------------------------------------------------------------------
    # Enhancement Helpers
    # ------------------------------------------------------------------

    def _get_rank_info(self, msg: str, context: dict) -> dict | None:
        """Detect rank keywords and return (rank_num, direction)."""
        msg_l = msg.lower()
        rank_num = None
        for word, num in _RE_RANK_NUM.items():
            if word in msg_l:
                rank_num = num
                break
        
        direction = "desc" # default
        if "lowest" in msg_l or "smallest" in msg_l:
            direction = "asc"
            
        is_next = any(k in msg_l for k in ["next", "then", "after that"])
        if is_next and context.get("last_displayed_rank"):
            rank_num = context["last_displayed_rank"] + 1

        if rank_num:
            return {"rank": rank_num, "direction": direction}
        return None

    def _ranked_query(self, rank_info: dict, start: date, end: date, label: str, context: dict) -> dict | None:
        """Handle Nth highest/lowest query."""
        rank = rank_info["rank"]
        direction = rank_info["direction"]
        
        # Use context if timeframe matches, else recalculate
        if context.get("context_sorted_expenses") and context.get("context_timeframe") == label:
            sorted_list = context["context_sorted_expenses"]
        else:
            matching = [e for e in self.expenses if start <= self._parse_date(e.get("date", "")) <= end]
            if not matching: return None
            by_sub = {}
            for e in matching:
                sub = e.get("subcategory") or e.get("category")
                by_sub[sub] = by_sub.get(sub, 0.0) + e["amount"]
            sorted_list = sorted(by_sub.items(), key=lambda x: x[1], reverse=(direction == "desc"))

        if not sorted_list: return None
        
        if rank > len(sorted_list):
            last_item, last_amt = sorted_list[-1]
            return self._clarify(f"You only have {len(sorted_list)} expenses in {label}. Your {'lowest' if direction == 'desc' else 'highest'} was {last_item}: {self._fmt(last_amt)}")

        item, amt = sorted_list[rank-1]
        ordinals = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th", 5: "5th"}
        rank_label = ordinals.get(rank, f"{rank}th")
        dir_label = "highest" if direction == "desc" else "lowest"
        
        reply = f"{item}: {self._fmt(amt)} was your {rank_label} {dir_label} expense in {label}."
        
        # Also show neighbors if it's the first few
        if rank <= 3 and len(sorted_list) >= 3:
            lines = []
            icons = {1: "🥇", 2: "🥈", 3: "🥉"}
            for i in range(min(3, len(sorted_list))):
                it, a = sorted_list[i]
                lines.append(f"{icons.get(i+1, '')} {i+1}{'st' if i==0 else 'nd' if i==1 else 'rd'}: {it}: {self._fmt(a)}")
            reply = "\n".join(lines) + f"\n(Top 3 expenses in {label})"

        return {
            "reply": reply, "total": round(amt, 2), "item": item, "timeframe": label,
            "last_displayed_rank": rank, "context_sorted_expenses": sorted_list, "context_timeframe": label
        }

    def _get_smart_empty_response(self, item_name: str, start: date, end: date, label: str, defaulted: bool) -> dict:
        """Enhancement 1: Intelligent empty response suggestions."""
        # Step 1: Category check
        cat_name = None
        for cat, subs in self._categories.items():
            if item_name.lower() == cat.lower() or item_name.lower() in [s.lower() for s in subs]:
                cat_name = cat
                break
        
        if cat_name:
            matching_cat = self._filter_expenses(cat_name, start, end)
            if matching_cat:
                total_cat = sum(e["amount"] for e in matching_cat)
                subs = sorted({e.get("subcategory") for e in matching_cat if e.get("subcategory")})
                return self._clarify(f"You didn't spend anything on '{item_name}' in {label}. But you spent {self._fmt(total_cat)} on '{cat_name}' (includes {', '.join(subs[:3])}). Did you mean '{cat_name}' instead?")

        # Step 2: Historical check
        all_time_start = date(2000, 1, 1)
        history = self._filter_expenses(item_name, all_time_start, date.today())
        if history:
            history.sort(key=lambda x: x["date"], reverse=True)
            last_e = history[0]
            return self._clarify(f"No '{item_name}' expenses in {label}. Your last '{item_name}' expense was {self._fmt(last_e['amount'])} on {last_e['date']}. Want to see that month?")

        # Step 3: Similar subcategories
        similar = []
        for cat_subs in self._categories.values():
            for s in cat_subs:
                if SequenceMatcher(None, item_name.lower(), s.lower()).ratio() > 0.7:
                    similar.append(s)
        
        if similar:
            return self._clarify(f"No '{item_name}' found. Did you mean '{similar[0]}'? Or search all time?")

        # Step 4: Generic fallback
        return self._clarify(f"No expenses found for '{item_name}' in {label}. Want to add one? Try 'top categories this month' to see where you spent.")

    # ── EMI / Loan Handlers ───────────────────────────────────────────────

    def _handle_emi_total(self, msg: str) -> dict:
        if not self.recurring_payments:
            return self._clarify("No EMI/loan payments found. Add recurring payments to track them.")
        
        # Detect timeframe (default to current month)
        tf_res = self._detect_timeframe(msg)
        start, end, label, _ = tf_res if tf_res else (date.today().replace(day=1), date.today(), "this month", True)
        
        total = 0.0
        details = []
        for p in self.recurring_payments:
            amt = float(p.get("amount", 0))
            name = p.get("name", "EMI")
            
            # Simple boundary check
            p_start = p.get("start_date", "2000-01-01")
            p_end   = p.get("end_date", "2099-12-31")
            
            if p_start <= end.isoformat() and p_end >= start.isoformat():
                total += amt
                details.append(f"• {name}: {self._fmt(amt)}")
        
        if not details:
            return {"reply": f"No active EMIs found for {label}.", "total": 0, "item": "emi_total", "timeframe": label}
            
        reply = f"Your total EMI payments for {label}: {self._fmt(total)}\n" + "\n".join(details)
        return {"reply": reply, "total": round(total, 2), "item": "emi_total", "timeframe": label}

    def _handle_emi_end(self) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")
        
        near_end = []
        for p in self.recurring_payments:
            # Handle if missing
            rem = p.get("remaining_installments")
            if rem is None: continue
            try:
                rem = int(rem)
            except: continue

            if rem <= 3:
                name = p.get("name", "EMI")
                end  = p.get("end_date", "")
                try:
                    dt = datetime.fromisoformat(end) if "T" in end else datetime.strptime(end, "%Y-%m-%d")
                    end_label = dt.strftime("%B %Y")
                except:
                    end_label = end
                near_end.append(f"• {name} - Ends {end_label} ({rem} payments left)")
        
        if not near_end:
            return {"reply": "None of your active EMIs are ending in the next 3 months."}
            
        return {"reply": "EMIs ending in next 3 months:\n" + "\n".join(near_end)}

    def _handle_emi_when(self, item_name: str) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")
        
        search_name = item_name.strip().lower()
        best_p = None
        for p in self.recurring_payments:
            if search_name in p.get("name", "").lower():
                best_p = p
                break
        
        if not best_p:
            return {"reply": f"I couldn't find a recurring payment named '{item_name}'."}
            
        name = best_p.get("name")
        end  = best_p.get("end_date", "unknown")
        rem  = best_p.get("remaining_installments", "?")
        try:
            dt = datetime.fromisoformat(end) if "T" in end else datetime.strptime(end, "%Y-%m-%d")
            end_label = dt.strftime("%B %d, %Y")
        except:
            end_label = end
            
        return {"reply": f"{name} ends on {end_label} ({rem} payments remaining)."}

    def _handle_emi_next(self) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")
        
        today_val = date.today()
        upcoming = []
        for p in self.recurring_payments:
            day = int(p.get("day_of_month", 1))
            name = p.get("name", "EMI")
            amt = float(p.get("amount", 0))
            if today_val.day <= day:
                due = date(today_val.year, today_val.month, day)
            else:
                m = today_val.month + 1
                y = today_val.year
                if m > 12: m, y = 1, y + 1
                due = date(y, m, min(day, 28))
            upcoming.append((due, name, amt))
            
        upcoming.sort(key=lambda x: x[0])
        if not upcoming:
            return {"reply": "No upcoming EMIs found."}
            
        nxt = upcoming[0]
        days_left = (nxt[0] - today_val).days
        if days_left == 0:
            due_str = "today"
        elif days_left == 1:
            due_str = "tomorrow"
        else:
            due_str = f"in {days_left} days ({nxt[0].strftime('%b %d')})"
            
        return {"reply": f"Your next EMI is {nxt[1]} ({self._fmt(nxt[2])}) due {due_str}."}

    @staticmethod
    def _month_iter(start: date, end: date):
        cur = date(start.year, start.month, 1)
        last = date(end.year, end.month, 1)
        while cur <= last:
            yield cur.year, cur.month
            if cur.month == 12:
                cur = date(cur.year + 1, 1, 1)
            else:
                cur = date(cur.year, cur.month + 1, 1)

    @staticmethod
    def _is_payment_active_in_month(payment: dict, year: int, month: int) -> bool:
        sy = int(payment.get("start_year", 0) or 0)
        sm = int(payment.get("start_month", 0) or 0)
        if not sy or not sm:
            return False
        if payment.get("type") == "fixed":
            return (year > sy) or (year == sy and month >= sm)
        total_months = int(payment.get("total_months", 0) or 0)
        if total_months <= 0:
            return False
        start_val = sy * 12 + (sm - 1)
        end_val = start_val + total_months - 1
        cur_val = year * 12 + (month - 1)
        return start_val <= cur_val <= end_val

    @staticmethod
    def _payment_status(payment: dict, ref: date = None) -> dict:
        ref = ref or date.today()
        ptype = payment.get("type", "emi")
        sy = int(payment.get("start_year", 0) or 0)
        sm = int(payment.get("start_month", 0) or 0)

        if ptype == "emi":
            total = int(payment.get("total_months", 0) or 0)
            elapsed = (ref.year - sy) * 12 + (ref.month - sm) + 1
            paid = min(max(elapsed, 0), total)
            remaining = max(total - paid, 0)
            active = paid < total
            offset = max(total - 1, 0)
            ey = sy + (sm - 1 + offset) // 12 if sy and sm and total else None
            em = (sm - 1 + offset) % 12 + 1 if sy and sm and total else None
            return {
                "type": "emi",
                "is_active": active,
                "remaining": remaining,
                "paid": paid,
                "total": total,
                "end_year": ey,
                "end_month": em,
            }

        return {
            "type": "fixed",
            "is_active": (ref.year > sy) or (ref.year == sy and ref.month >= sm),
            "remaining": None,
            "paid": 0,
            "total": 0,
            "end_year": None,
            "end_month": None,
        }

    @staticmethod
    def _payment_scope(msg: str) -> str:
        lowered = msg.lower()
        if "recurring" in lowered or "fixed" in lowered:
            return "all"
        return "emi"

    def _scoped_payments(self, msg: str) -> list:
        scope = self._payment_scope(msg)
        if scope == "all":
            return list(self.recurring_payments)
        return [p for p in self.recurring_payments if p.get("type", "emi") == "emi"]

    def _handle_emi_total(self, msg: str) -> dict:
        payments = self._scoped_payments(msg)
        if not payments:
            return self._clarify("No EMI/loan payments found. Add recurring payments to track them.")

        tf_res = self._detect_timeframe(msg)
        start, end, label, _ = tf_res if tf_res else (date.today().replace(day=1), date.today(), "this month", True)

        total = 0.0
        details = []
        for p in payments:
            amt = float(p.get("amount", 0))
            name = p.get("name", "EMI")
            active_months = sum(
                1 for y, m in self._month_iter(start, end)
                if self._is_payment_active_in_month(p, y, m)
            )
            if active_months > 0:
                payment_total = amt * active_months
                total += payment_total
                month_note = f" x {active_months}" if active_months > 1 else ""
                details.append(f"* {name}: {self._fmt(payment_total)}{month_note}")

        if not details:
            return {"reply": f"No active EMIs found for {label}.", "total": 0, "item": "emi_total", "timeframe": label}

        reply = f"Your total EMI payments for {label}: {self._fmt(total)}\n" + "\n".join(details)
        return {"reply": reply, "total": round(total, 2), "item": "emi_total", "timeframe": label}

    def _handle_emi_end(self) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")

        near_end = []
        for p in self.recurring_payments:
            status = self._payment_status(p)
            rem = status.get("remaining")
            if status.get("type") != "emi" or rem is None or not status.get("is_active"):
                continue
            if rem <= 3:
                name = p.get("name", "EMI")
                ey = status.get("end_year")
                em = status.get("end_month")
                end_label = f"{MONTH_DISPLAY[em]} {ey}" if ey and em else "soon"
                near_end.append(f"* {name} - Ends {end_label} ({rem} payment{'s' if rem != 1 else ''} left)")

        if not near_end:
            return {"reply": "None of your active EMIs are ending in the next 3 months."}

        return {"reply": "EMIs ending in next 3 months:\n" + "\n".join(near_end)}

    def _handle_emi_when(self, item_name: str) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")

        search_name = item_name.strip().lower()
        best_p = None
        for p in self.recurring_payments:
            if search_name in p.get("name", "").lower():
                best_p = p
                break

        if not best_p:
            return {"reply": f"I couldn't find a recurring payment named '{item_name}'."}

        name = best_p.get("name")
        status = self._payment_status(best_p)
        if status.get("type") != "emi":
            return {"reply": f"{name} is a recurring fixed payment and does not have a fixed EMI end date."}

        ey = status.get("end_year")
        em = status.get("end_month")
        rem = status.get("remaining", "?")
        end_label = f"{MONTH_DISPLAY[em]} {ey}" if ey and em else "unknown"
        return {"reply": f"{name} ends in {end_label} ({rem} payment{'s' if rem != 1 else ''} remaining)."}

    def _handle_emi_next(self) -> dict:
        if not self.recurring_payments:
            return self._clarify("No recurring payments found.")

        today_val = date.today()
        upcoming = []
        for p in self.recurring_payments:
            if not self._payment_status(p, today_val).get("is_active"):
                continue
            day = int(p.get("day_of_month", 1) or 1)
            name = p.get("name", "EMI")
            amt = float(p.get("amount", 0))
            if today_val.day <= day:
                due = date(today_val.year, today_val.month, min(day, 28))
            else:
                m = today_val.month + 1
                y = today_val.year
                if m > 12:
                    m, y = 1, y + 1
                due = date(y, m, min(day, 28))
            upcoming.append((due, name, amt))

        upcoming.sort(key=lambda x: x[0])
        if not upcoming:
            return {"reply": "No upcoming EMIs found."}

        nxt = upcoming[0]
        days_left = (nxt[0] - today_val).days
        if days_left == 0:
            due_str = "today"
        elif days_left == 1:
            due_str = "tomorrow"
        else:
            due_str = f"in {days_left} days ({nxt[0].strftime('%b %d')})"

        return {"reply": f"Your next EMI is {nxt[1]} ({self._fmt(nxt[2])}) due {due_str}."}

    def _handle_emi_count(self, msg: str) -> dict:
        payments = self._scoped_payments(msg)
        scope = self._payment_scope(msg)
        if not payments:
            return self._clarify("No EMI/loan payments found.")

        tf_res = self._detect_timeframe(msg)
        start, end, label, _ = tf_res if tf_res else (date.today().replace(day=1), date.today(), "this month", True)

        active = []
        for p in payments:
            if any(self._is_payment_active_in_month(p, y, m) for y, m in self._month_iter(start, end)):
                active.append(p)

        if not active:
            noun = "EMIs or recurring payments" if scope == "all" else "EMIs"
            return {"reply": f"You don't have any active {noun} in {label}.", "total": 0, "item": "emi_count", "timeframe": label}

        names = ", ".join(p.get("name", "EMI") for p in active[:5])
        if len(active) > 5:
            names += ", ..."
        noun = "EMI/recurring payment" if scope == "all" else "EMI"
        return {
            "reply": f"You have {len(active)} active {noun}{'s' if len(active) != 1 else ''} in {label}: {names}.",
            "total": len(active),
            "item": "emi_count",
            "timeframe": label,
        }

    def _handle_emi_latest(self, msg: str) -> dict:
        payments = self._scoped_payments(msg)
        scope = self._payment_scope(msg)
        if not payments:
            return self._clarify("No recurring payments found.")

        payments = [
            p for p in payments
            if int(p.get("start_year", 0) or 0) and int(p.get("start_month", 0) or 0)
        ]
        if not payments:
            return {"reply": "I couldn't find any recurring payments with a valid start date."}

        latest = max(payments, key=lambda p: (int(p.get("start_year", 0)), int(p.get("start_month", 0)), p.get("name", "")))
        sy = int(latest.get("start_year", 0))
        sm = int(latest.get("start_month", 0))
        started = f"{MONTH_DISPLAY[sm]} {sy}"
        name = latest.get("name", "EMI")
        amount = self._fmt(float(latest.get("amount", 0)))
        kind = "recurring payment" if scope == "all" else "EMI"
        return {"reply": f"Your latest {kind} is {name} for {amount}, started in {started}."}
