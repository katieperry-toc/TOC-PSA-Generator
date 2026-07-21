"""Business intelligence: discovery-note parsing, normalization, and scope/
pricing rules.

This module knows nothing about Word documents — it accepts and returns
plain strings/dicts/lists only. python-docx must never be imported here.
Future scope-recommendation, hours-estimation, and pricing-recommendation
logic belongs in this module.
"""

import json
import os
import re
from datetime import date
from typing import Any, Dict, List, Sequence, Tuple

from dotenv import load_dotenv

from config import (
    ENGAGEMENT_BOTH,
    ENGAGEMENT_HR,
    ENGAGEMENT_TA,
    FIELDS,
    HR_SERVICES,
    TA_SERVICES,
)
from scope_library import get_scope_template, render_fallback_scope_text

load_dotenv()


def clean(value: Any) -> str:
    return "" if value is None else str(value).strip()

def first_name(value: Any) -> str:
    """Return a natural first-name salutation from a full contact name."""
    raw = clean(value)
    if not raw:
        return ""
    # Handles names such as "Jay Devine", "Jay A. Devine", and "Devine, Jay".
    if "," in raw:
        parts = [part.strip() for part in raw.split(",") if part.strip()]
        if len(parts) > 1:
            raw = parts[1]
    return raw.split()[0] if raw.split() else raw

def clean_multiline(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in text.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)

def tbc(value: Any) -> str:
    return clean(value) or "To Be Confirmed"

def normalize_hr_service(value: str) -> str:
    raw = clean(value)
    if raw in HR_SERVICES:
        return raw
    lowered = raw.lower()
    aliases = {
        "project": "HR Project Support",
        "fractional": "Fractional/Interim HR Support",
        "interim": "Fractional/Interim HR Support",
        "subscription": "HR Subscription",
    }
    for keyword, normalized in aliases.items():
        if keyword in lowered:
            return normalized
    return ""

def normalize_ta_service(value: str) -> str:
    raw = clean(value)
    if raw in TA_SERVICES:
        return raw
    lowered = raw.lower()
    aliases = {
        "full cycle": "Full Cycle Talent Acquisition Support",
        "full-cycle": "Full Cycle Talent Acquisition Support",
        "talent acquisition": "Full Cycle Talent Acquisition Support",
        "recruit": "Full Cycle Talent Acquisition Support",
        "sourcing": "Sourcing Support",
    }
    for keyword, normalized in aliases.items():
        if keyword in lowered:
            return normalized
    return ""

def engagement_from_services(hr_service: str, ta_service: str) -> str:
    has_hr = bool(normalize_hr_service(hr_service))
    has_ta = bool(normalize_ta_service(ta_service))
    if has_hr and has_ta:
        return ENGAGEMENT_BOTH
    if has_hr:
        return ENGAGEMENT_HR
    if has_ta:
        return ENGAGEMENT_TA
    return ""

def normalize_engagement(
    engagement_type: str,
    hr_service: str = "",
    ta_service: str = "",
) -> str:
    lowered = clean(engagement_type).lower()

    has_hr = any(
        phrase in lowered
        for phrase in [
            "human resources",
            "hr support",
            "hr consulting",
            "fractional hr",
            "interim hr",
        ]
    )
    has_ta = any(
        phrase in lowered
        for phrase in [
            "talent acquisition",
            "recruitment",
            "recruiting",
            "sourcing",
            "raas",
        ]
    )

    # An explicitly selected engagement type always wins. This prevents the
    # default hidden service dropdowns from incorrectly turning every PSA into
    # a combined HR and Talent Acquisition agreement.
    if has_hr and has_ta:
        return ENGAGEMENT_BOTH
    if has_hr:
        return ENGAGEMENT_HR
    if has_ta:
        return ENGAGEMENT_TA

    return engagement_from_services(hr_service, ta_service)

def infer_engagement_from_notes(notes: str, current: str = "") -> str:
    normalized_current = normalize_engagement(current)
    if normalized_current:
        return normalized_current
    lowered = clean(notes).lower()
    hr_signals = ["human resources", "hr project", "fractional hr", "interim hr", "hr subscription", "employee handbook", "leave policy", "compensation study", "benefits assessment", "hris", "employee relations", "compliance"]
    ta_signals = ["talent acquisition", "recruitment", "recruiting", "full cycle", "full-cycle", "sourcing", "candidate", "roles to fill", "per role"]
    has_hr = any(signal in lowered for signal in hr_signals)
    has_ta = any(signal in lowered for signal in ta_signals)
    if has_hr and has_ta:
        return ENGAGEMENT_BOTH
    if has_hr:
        return ENGAGEMENT_HR
    if has_ta:
        return ENGAGEMENT_TA
    return ""

def extract_timing_from_notes(notes: str) -> Dict[str, str]:
    """Deterministically recover hours fields when the AI leaves them blank."""
    text = clean_multiline(notes)
    extracted = {
        "estimated_hours": "",
        "weekly_commitment": "",
        "minimum_engagement": "",
    }

    def first_match(patterns: Sequence[str]) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
            if match:
                value = clean(match.group(1))
                value = re.sub(r"\s+", " ", value).strip(" .;,-")
                return value
        return ""

    extracted["estimated_hours"] = first_match(
        [
            r"^\s*Estimated(?:\s+Total)?\s*(?:Hours?)?\s*:?\s*(?:\n\s*)?(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?)\b",
            r"^\s*Total\s+Estimated\s+Hours?\s*:?\s*(?:\n\s*)?(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?)\b",
        ]
    )

    extracted["weekly_commitment"] = first_match(
        [
            r"^\s*Weekly\s+Commitment\s*:?\s*(?:\n\s*)?(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?(?:\s*(?:per|/)\s*week)?)\b",
            r"^\s*Projected\s*#?\s*of\s*Hours?\s*:?\s*(?:\n\s*)?(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?\s*(?:per|/)\s*week)\b",
            r"\b(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?\s*(?:per|/)\s*week)\b",
        ]
    )

    extracted["minimum_engagement"] = first_match(
        [
            r"^\s*(?:Project\s+)?Minimum(?:\s+Engagement)?\s*:?\s*(?:\n\s*)?(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?)?)\b",
            r"\bMinimum\s+(\d+(?:\s*[-–]\s*\d+)?\s*(?:hours?|hrs?))\b",
        ]
    )

    return extracted

def apply_raas_defaults(data: Dict[str, str]) -> Dict[str, str]:
    """Apply TOC's standard terms for full-cycle RaaS engagements."""
    updated = dict(data)
    engagement = normalize_engagement(
        updated.get("engagement_type", ""),
        updated.get("hr_service_type", ""),
        updated.get("ta_service_type", ""),
    )
    ta_service = normalize_ta_service(updated.get("ta_service_type", ""))

    if engagement in {ENGAGEMENT_TA, ENGAGEMENT_BOTH} and (
        ta_service == "Full Cycle Talent Acquisition Support"
        or not ta_service
    ):
        updated["ta_service_type"] = "Full Cycle Talent Acquisition Support"
        updated["weekly_commitment"] = "10 hours per week"
        updated["minimum_engagement"] = "Minimum 40 hours"

    # Estimated hours remain search-specific. Preserve an explicit value from
    # the notes, but do not create one when none was supplied.
    return updated

def parse_discovery_notes(notes: str) -> Dict[str, str]:
    notes = clean_multiline(notes)
    if not notes:
        raise ValueError("Paste discovery notes before selecting Read Notes.")
    api_key = clean(os.getenv("OPENAI_API_KEY"))
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY was not found in the .env file.")

    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    prompt = f"""
Extract information for a TOC Professional Services Agreement.
Return one valid JSON object with exactly these fields:
{", ".join(FIELDS)}

Rules:
- Use an empty string when information is not provided.
- Do not invent client, contact, pricing, address, timing, or scope details.
- engagement_type must be exactly one of: "{ENGAGEMENT_HR}", "{ENGAGEMENT_TA}", or "{ENGAGEMENT_BOTH}".
- Choose HR for HR consulting, projects, fractional/interim HR, subscription, policy, compliance, compensation, benefits, HRIS, handbook, and employee relations work.
- Choose Talent Acquisition for recruiting, full-cycle talent acquisition, candidate sourcing, or hiring support.
- Choose combined when both are included.
- hr_service_type must be one of: {", ".join(HR_SERVICES)}.
- ta_service_type must be one of: {", ".join(TA_SERVICES)}.
- Leave irrelevant service fields empty.
- For combined work, populate both service and both scope fields.
- Keep total hours, weekly commitment, and minimum engagement separate.
- Preserve the complete client mailing address. Put the street address in address_1 and put suite/unit plus city, state, and ZIP together in address_2. Do not drop the city, state, or ZIP when a suite is provided.
- Preserve ranges such as 10-15 or 60-75.
- Copy supplied scope language faithfully and do not invent missing scope.
- Write a concise client-facing scope summary using only facts expressly stated in the notes.
- Preserve the role title, compensation, deliverables, and search-specific estimated hours when supplied.
- Do not expand a short scope note into a generic list of recruiting or HR tasks.
- For Full Cycle Talent Acquisition Support, use the standard weekly commitment of 10 hours per week and the standard minimum engagement of Minimum 40 hours.
- Estimated total hours are search-specific: preserve them when explicitly provided, otherwise leave estimated_hours empty.
- Return JSON only.

Discovery notes:
{notes}
"""
    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "You accurately extract contract information from discovery notes, emails, and statements of work."},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0,
    )
    raw = response.choices[0].message.content
    if not raw:
        raise ValueError("OpenAI returned an empty response.")
    parsed = json.loads(raw)
    result = {field: clean(parsed.get(field, "")) for field in FIELDS}

    # The discovery-note layouts often place timing labels and values on
    # separate lines. Recover those values directly when the model omits them.
    timing_fallback = extract_timing_from_notes(notes)
    for timing_field in (
        "estimated_hours",
        "weekly_commitment",
        "minimum_engagement",
    ):
        if not clean(result.get(timing_field, "")):
            result[timing_field] = timing_fallback[timing_field]

    legacy_service = clean(result.get("service_type"))
    result["hr_service_type"] = normalize_hr_service(result["hr_service_type"] or legacy_service)
    result["ta_service_type"] = normalize_ta_service(result["ta_service_type"] or legacy_service)
    result["engagement_type"] = normalize_engagement(result["engagement_type"], result["hr_service_type"], result["ta_service_type"]) or infer_engagement_from_notes(notes, result["engagement_type"])
    legacy_scope = clean_multiline(result.get("scope_of_work"))
    result["hr_scope_of_work"] = clean_multiline(result["hr_scope_of_work"])
    result["ta_scope_of_work"] = clean_multiline(result["ta_scope_of_work"])
    if result["engagement_type"] == ENGAGEMENT_HR:
        result["hr_service_type"] = result["hr_service_type"] or "HR Project Support"
        result["hr_scope_of_work"] = result["hr_scope_of_work"] or legacy_scope
    elif result["engagement_type"] == ENGAGEMENT_TA:
        result["ta_service_type"] = result["ta_service_type"] or "Full Cycle Talent Acquisition Support"
        result["ta_scope_of_work"] = result["ta_scope_of_work"] or legacy_scope
    elif result["engagement_type"] == ENGAGEMENT_BOTH:
        result["hr_service_type"] = result["hr_service_type"] or "HR Project Support"
        result["ta_service_type"] = result["ta_service_type"] or "Full Cycle Talent Acquisition Support"
    return apply_raas_defaults(result)

def agreement_label(engagement_type: str) -> str:
    if engagement_type == ENGAGEMENT_HR:
        return "Human Resources Support"
    if engagement_type == ENGAGEMENT_TA:
        return "Talent Acquisition Support"
    if engagement_type == ENGAGEMENT_BOTH:
        return "Human Resources and Talent Acquisition Support"
    return "Professional Services"

def minimum_text(value: str) -> str:
    raw = clean(value)
    if not raw:
        return "To Be Confirmed"
    return raw

def normalize_commitment_text(value: str) -> str:
    raw = clean(value)
    if not raw:
        return ""
    if re.fullmatch(r"\d+(?:\.\d+)?(?:\s*-\s*\d+(?:\.\d+)?)?", raw):
        return f"{raw} hours per week"
    return raw

def normalize_minimum_sentence(value: str) -> str:
    raw = clean(value)
    if not raw:
        return ""

    cleaned = raw.rstrip(". ")
    number_match = re.fullmatch(
        r"(?:minimum\s*)?(\d+(?:\.\d+)?)\s*(?:hours?|hrs?)?",
        cleaned,
        flags=re.IGNORECASE,
    )
    if number_match:
        return f"Minimum {number_match.group(1)} hours"

    cleaned = re.sub(
        r"\bminimum\s+of\s+(\d+(?:\.\d+)?)\s*hours?\b",
        r"Minimum \1 hours",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(
        r"\bminimum\s+(\d+(?:\.\d+)?)\s*hours?\b",
        r"Minimum \1 hours",
        cleaned,
        flags=re.IGNORECASE,
    )
    return cleaned

def minimum_row_value(value: str) -> str:
    """Format Minimum Engagement for the Agreement Summary's "Project
    Minimum" row specifically.

    That row's own label already says "Minimum", so the value should read
    "40 hours" rather than "Minimum 40 hours" — saying it twice was the
    exact redundancy TOC asked to remove. Returns "" (not "To Be
    Confirmed") when there's nothing to format, so the caller can apply
    its own fallback the same way it does with normalize_minimum_sentence.

    Every other call site (e.g. the Standard Services pricing table's
    Notes column, which has no "Minimum" label of its own to lean on)
    should keep using normalize_minimum_sentence directly so the phrase
    stays self-contained there.
    """
    sentence = normalize_minimum_sentence(value)
    return re.sub(r"^\s*Minimum\s+", "", sentence, flags=re.IGNORECASE)

def format_rate(value: str) -> str:
    raw = clean(value)
    if not raw:
        return "To Be Confirmed"
    compact = raw.lower().replace("per hour", "").replace("/hour", "").replace("$", "").replace(",", "").strip()
    try:
        return f"${float(compact):,.0f}/hour"
    except ValueError:
        return raw if "$" in raw else f"${raw}/hour"

def hours_value(value: str) -> str:
    raw = clean(value)
    if not raw:
        return "To Be Confirmed"
    return re.sub(r"\s*(hours?|hrs?)\s*$", "", raw, flags=re.IGNORECASE).strip()

def normalize_hours_display(value: str) -> str:
    raw = clean(value)
    if not raw:
        return ""
    if re.search(r"\b(hours?|hrs?)\b", raw, flags=re.IGNORECASE):
        return raw
    return f"{raw} hours"

def service_scope(
    service: str,
    supplied_scope: str,
    estimated_hours: str = "",
) -> str:
    """Return the final, client-facing scope text for one service.

    supplied_scope is whatever is currently in the service's Scope of Work
    field (personalized by generate_scope_for_service() when notes were
    read, or manually edited by the user) — it always wins when present.
    When it's empty, fall back to the approved TOC Scope Library's evergreen
    baseline text for the service (scope_library.render_fallback_scope_text)
    instead of a generic one-line placeholder, so the Scope of Services
    section is never blank even before any notes are read.
    """
    supplied = clean_multiline(supplied_scope)
    if supplied:
        return supplied

    fallback = render_fallback_scope_text(service)
    return fallback or "To Be Confirmed"


# ---------------------------------------------------------------------------
# AI scope personalization + validation
# ---------------------------------------------------------------------------
# Turns an approved scope_library.py template into a final, contract-ready
# Scope of Services for one service, using only facts explicitly present in
# the client's discovery notes. Always falls back to the approved evergreen
# baseline (never blank, never raises) if personalization isn't possible or
# doesn't pass validation — see generate_scope_for_service().

_PLACEHOLDER_PATTERN = re.compile(r"\[[^\]\n]{1,80}\]|\{\{[^}\n]{1,80}\}\}")
_EXEC_SEARCH_PATTERN = re.compile(r"executive\s+search", re.IGNORECASE)
_PRICE_PATTERN = re.compile(r"\$\s?\d")


def build_scope_context(form_data: Dict[str, str], discovery_notes: str) -> Dict[str, str]:
    """Assemble the facts available to personalize an approved scope template.

    Pulls only values already present in form_data/discovery_notes — this
    function never invents or infers a fact that isn't already there.
    """
    return {
        "client_name": clean(form_data.get("client_name")),
        "engagement_type": clean(form_data.get("engagement_type")),
        "hr_service_type": clean(form_data.get("hr_service_type")),
        "ta_service_type": clean(form_data.get("ta_service_type")),
        "hourly_rate": clean(form_data.get("hourly_rate")),
        "weekly_commitment": clean(form_data.get("weekly_commitment")),
        "minimum_engagement": clean(form_data.get("minimum_engagement")),
        "estimated_hours": clean(form_data.get("estimated_hours")),
        "discovery_notes": clean_multiline(discovery_notes),
    }


def validate_generated_scope(
    scope: str, service_key: str, context: Dict[str, str]
) -> List[str]:
    """Return a list of problems found in an AI-personalized scope.

    An empty list means the scope is safe to use. Never raises. Callers
    should fall back to the approved baseline scope whenever this returns
    any problems — see generate_scope_for_service().
    """
    problems: List[str] = []
    text = clean_multiline(scope)

    if not text:
        problems.append("the generated scope was blank")
        return problems
    if _EXEC_SEARCH_PATTERN.search(text):
        problems.append("the generated scope referenced Executive Search")
    if _PRICE_PATTERN.search(text):
        problems.append("the generated scope appears to include pricing")
    rate = context.get("hourly_rate", "")
    if rate and rate in text:
        problems.append("the generated scope repeated the hourly rate")
    if re.search(r"\bTBD\b", text):
        problems.append("the generated scope left a visible TBD placeholder")
    if _PLACEHOLDER_PATTERN.search(text):
        problems.append("the generated scope left an unresolved placeholder")
    if _FORBIDDEN_HEADING_PATTERN.search(text):
        problems.append(
            "the generated scope used a sales-style or duplicate heading "
            "instead of a plain title followed by bullets"
        )
    # The approved format is always Title / short paragraph / bullets — a
    # model that collapses everything into one narrative paragraph (no
    # bullets at all) has abandoned the approved Scope Library structure
    # even if nothing else looks wrong, so this is checked explicitly
    # rather than relying on the other checks to catch it indirectly.
    bullet_lines = [
        line for line in text.split("\n") if line.strip().startswith("•")
    ]
    if len(bullet_lines) < 2:
        problems.append(
            "the generated scope did not keep the required Primary "
            "Services bullet list (it collapsed into a narrative paragraph)"
        )
    word_count = len(text.split())
    if word_count > 200:
        problems.append(
            "the generated scope was unexpectedly long (TOC's target is "
            "roughly 90-120 words)"
        )
    return problems


_FORBIDDEN_HEADING_PATTERN = re.compile(
    r"^\s*(what.?s included|our solution|why choose us|benefits|primary services)\s*:?\s*$",
    re.IGNORECASE | re.MULTILINE,
)


def personalize_scope(template: Dict[str, Any], context: Dict[str, str]) -> str:
    """Ask the model to personalize one approved scope template with context.

    Raises on any failure (missing API key, no notes to work from, a
    network/model error, or an empty response) so the caller
    (generate_scope_for_service) can fall back to the approved baseline
    scope rather than ever returning unvalidated text.
    """
    notes = context.get("discovery_notes", "")
    if not notes:
        raise ValueError("no discovery notes available to personalize with")

    api_key = clean(os.getenv("OPENAI_API_KEY"))
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY was not found in the .env file.")

    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    exclusions_line = (
        "Excluded from this service — never include: "
        + "; ".join(template["exclusions"])
        if template["exclusions"]
        else "No standing exclusions for this service."
    )

    prompt = f"""
You are drafting a short, standardized Scope of Services summary for a TOC Professional Services Agreement with {context.get('client_name') or 'the Client'}.

Approved service: {template['display_title']}

Approved baseline overview (do not contradict this):
{template['overview']}

Approved Primary Services menu for this service — choose the 3 (or, only when genuinely necessary, 4) that best fit the discovery notes. Do not invent items outside this list and the optional items below unless the notes clearly describe something in it:
{chr(10).join('- ' + item for item in template['primary_services'])}

Optional additional items (use at most one, as a 4th bullet, and only if the discovery notes clearly support it):
{chr(10).join('- ' + item for item in template['optional_modules'])}

{exclusions_line}

Guardrails:
{chr(10).join('- ' + g for g in template['guardrails'])}

Client discovery notes (use only to tailor wording and choose the most relevant Primary Services bullets — never invent a fact that is not stated here; if the notes are long or detailed, summarize them rather than reproducing every item):
{notes}

Worked example of the required format (a different service, shown only so the shape is unmistakable — do not copy its wording):

HR Subscription
Provide ongoing HR guidance for a 60-person distribution company navigating rapid headcount growth.

• Strategic HR Guidance
• Employee Relations Support
• Policy and Compliance Assistance

Output format — follow exactly, nothing more:
Line 1: a short engagement title. Use "{template['display_title']}" unless the discovery notes clearly describe a distinct, named project, in which case a more specific title is fine — never invent an unrelated service.
Line 2: ONE short paragraph — one or two sentences, never more — describing the purpose and value of the engagement. This is the only place a client-specific detail (role title, industry, a notable challenge, a headcount figure, etc.) gets woven in; keep it brief even when the notes contain a lot of detail.
Line 3: a blank line.
Then: 3 bullets (4 only when genuinely necessary), one per line, each a short 2-6 word phrase starting with "• ", drawn from the approved Primary Services menu above using that menu's own wording — do not rewrite, expand, or merge the bullets into sentences.

Never collapse this into a single narrative paragraph, and never drop the bullets — even when the discovery notes are long or contain many facts, choose the 1-2 most important ones for the overview sentence and let the rest go; do not try to fit every detail in. The bullets are a required, separate list every time, never optional.

Do not include a "Primary Services" heading, or any other heading, anywhere in the output. The Word template this text is inserted into already renders a bold "Primary Services" heading immediately above it, so repeating that heading (or using a different sales-style heading such as "What's Included," "Our Solution," "Why Choose Us," or "Benefits") would duplicate or clash with it.

Strict rules:
- Roughly 90-120 words total across the whole output (a little more is fine if the client-specific detail genuinely needs it, but this is a short executive summary, not a proposal).
- No timeline unless the discovery notes specifically call for one.
- No pricing, rate, or investment language of any kind — this includes the client's own budget or a candidate's target salary; that belongs elsewhere, not in this scope text.
- No long introduction — start directly with the title.
- No detailed methodology or step-by-step process narration.
- No repetitive language between the paragraph and the bullets.
- No legal or contractual language (that lives elsewhere in the agreement).
- No exaggerated marketing or sales language (no "best-in-class," "world-class," "cutting-edge," etc.).
- No promises or guarantees of outcome.
- Never include Executive Search.
- Never include pricing, hourly rate, weekly commitment, or minimum engagement — those already appear elsewhere in the agreement.
- Never use placeholder text such as TBD, [brackets], or {{{{double braces}}}}.
- Output only the final scope text in the exact format above — no commentary, no markdown formatting, no extra headings.
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role": "system",
                "content": (
                    "You write short, polished, executive-friendly Scope of "
                    "Services summaries for The O'Connor Group: a title, one "
                    "short overview paragraph (one or two sentences), and 3 "
                    "(occasionally 4) concise bullet points under an implied "
                    "\"Primary Services\" heading that you never write out "
                    "yourself. Roughly 90-120 words total. You never invent "
                    "facts, never include Executive Search, never include "
                    "pricing or commercial terms, and never use sales or "
                    "marketing language."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )
    text = clean_multiline(response.choices[0].message.content or "")
    if not text:
        raise ValueError("OpenAI returned an empty scope")
    return text


def generate_scope_for_service(
    service_key: str,
    form_data: Dict[str, str],
    discovery_notes: str,
) -> Tuple[str, List[str]]:
    """Produce the final Scope of Services text for one service.

    Always returns usable text and never raises. Falls back to the approved
    evergreen baseline (scope_library.render_fallback_scope_text) whenever
    the service has no approved template, there are no discovery notes to
    personalize with, the AI call fails, or the AI output fails
    validate_generated_scope. The returned warnings list explains any such
    fallback but never blocks the PSA from generating.
    """
    warnings: List[str] = []
    template = get_scope_template(service_key)
    if template is None:
        warnings.append(
            f"'{service_key}' has no approved TOC scope template; the "
            "Scope of Services section may need to be completed manually."
        )
        return "To Be Confirmed", warnings

    context = build_scope_context(form_data, discovery_notes)

    if not context["discovery_notes"]:
        return render_fallback_scope_text(service_key), warnings

    try:
        personalized = personalize_scope(template, context)
    except Exception:
        warnings.append(
            f"AI personalization was not applied for {service_key}; the "
            "approved baseline scope was used instead."
        )
        return render_fallback_scope_text(service_key), warnings

    problems = validate_generated_scope(personalized, service_key, context)
    if problems:
        warnings.append(
            f"The AI-personalized scope for {service_key} did not pass "
            "validation (" + "; ".join(problems) + "); the approved "
            "baseline scope was used instead."
        )
        return render_fallback_scope_text(service_key), warnings

    return personalized, warnings


# ---------------------------------------------------------------------------
# Discovery-notes vs. structured-selection conflict detection
# ---------------------------------------------------------------------------
# Structured selections (the HR/TA Service Type dropdowns) and commercial
# terms typed into the app always win over discovery notes — this never
# changes which service_key is used for scope generation. It only produces
# a visible internal warning when the notes seem to describe a different
# service than the one currently selected, so a mismatch is surfaced rather
# than silently guessed at.

_HR_NOTES_HINTS = {
    "HR Project Support": ["hr project", "project-based hr", "hr initiative"],
    "Fractional/Interim HR Support": ["fractional hr", "interim hr", "fractional/interim"],
    "HR Subscription": ["hr subscription", "subscription hours", "monthly hr retainer"],
}
_TA_NOTES_HINTS = {
    "Full Cycle Talent Acquisition Support": ["full cycle", "full-cycle", "raas", "recruitment-as-a-service"],
    "Sourcing Support": ["sourcing support", "sourcing only", "candidate sourcing"],
}


def _hint_from_notes(notes: str, hints: Dict[str, List[str]]) -> str:
    lowered = clean(notes).lower()
    if not lowered:
        return ""
    for service, keywords in hints.items():
        if any(keyword in lowered for keyword in keywords):
            return service
    return ""


def detect_scope_conflicts(
    discovery_notes: str, hr_service_type: str, ta_service_type: str
) -> List[str]:
    """Return warnings when discovery notes seem to describe a different
    service than the one currently selected in the structured form.

    The selected service (hr_service_type/ta_service_type) always remains
    the one used for scope generation — this function only surfaces a
    warning, it never changes the selection.
    """
    warnings: List[str] = []
    hr_hint = _hint_from_notes(discovery_notes, _HR_NOTES_HINTS)
    if hr_hint and clean(hr_service_type) and hr_hint != clean(hr_service_type):
        warnings.append(
            f"Discovery notes appear to describe '{hr_hint}', but "
            f"'{hr_service_type}' is the selected HR Service Type; the "
            "selected service was used for the Scope of Services."
        )
    ta_hint = _hint_from_notes(discovery_notes, _TA_NOTES_HINTS)
    if ta_hint and clean(ta_service_type) and ta_hint != clean(ta_service_type):
        warnings.append(
            f"Discovery notes appear to describe '{ta_hint}', but "
            f"'{ta_service_type}' is the selected Talent Acquisition "
            "Service Type; the selected service was used for the Scope of "
            "Services."
        )
    return warnings


def engagement_includes_hr(engagement_type: str) -> bool:
    return engagement_type in {ENGAGEMENT_HR, ENGAGEMENT_BOTH}

def engagement_includes_ta(engagement_type: str) -> bool:
    return engagement_type in {ENGAGEMENT_TA, ENGAGEMENT_BOTH}

def selected_services_from_data(data: Dict[str, str]) -> List[str]:
    engagement_type = normalize_engagement(
        data.get("engagement_type", ""),
        data.get("hr_service_type", ""),
        data.get("ta_service_type", ""),
    )
    selected: List[str] = []
    if engagement_includes_hr(engagement_type):
        selected.append(
            normalize_hr_service(data.get("hr_service_type", ""))
            or "HR Project Support"
        )
    if engagement_includes_ta(engagement_type):
        selected.append(
            normalize_ta_service(data.get("ta_service_type", ""))
            or "Full Cycle Talent Acquisition Support"
        )
    return selected

def combined_scope_from_data(data: Dict[str, str]) -> str:
    engagement_type = normalize_engagement(
        data.get("engagement_type", ""),
        data.get("hr_service_type", ""),
        data.get("ta_service_type", ""),
    )
    sections: List[str] = []
    if engagement_includes_hr(engagement_type):
        hr_scope = service_scope(
            data.get("hr_service_type", ""),
            data.get("hr_scope_of_work", ""),
            data.get("estimated_hours", ""),
        )
        if engagement_type == ENGAGEMENT_BOTH:
            sections.append(f"Human Resources Support\n{hr_scope}")
        else:
            sections.append(hr_scope)
    if engagement_includes_ta(engagement_type):
        ta_scope = service_scope(
            data.get("ta_service_type", ""),
            data.get("ta_scope_of_work", ""),
            data.get("estimated_hours", ""),
        )
        if engagement_type == ENGAGEMENT_BOTH:
            sections.append(f"Talent Acquisition Support\n{ta_scope}")
        else:
            sections.append(ta_scope)
    return "\n\n".join(sections) or "To Be Confirmed"

def current_agreement_date() -> str:
    """Return today's date without a leading zero on Windows or Unix."""
    try:
        return date.today().strftime("%B %#d, %Y")
    except ValueError:
        return date.today().strftime("%B %d, %Y").replace(" 0", " ")

def build_replacements(data: Dict[str, str]) -> Dict[str, str]:
    engagement_type = normalize_engagement(data.get("engagement_type", ""), data.get("hr_service_type", ""), data.get("ta_service_type", "")) or "To Be Confirmed"
    hr_service = normalize_hr_service(data.get("hr_service_type", ""))
    ta_service = normalize_ta_service(data.get("ta_service_type", ""))
    current_date = current_agreement_date()
    return {
        "{{CLIENT_NAME}}": tbc(data.get("client_name")), "{{CLIENT}}": tbc(data.get("client_name")), "[CLIENT NAME]": tbc(data.get("client_name")), "<<CLIENT NAME>>": tbc(data.get("client_name")),
        "{{CONTACT_NAME}}": tbc(data.get("contact_name")), "[CONTACT NAME]": tbc(data.get("contact_name")),
        "{{CONTACT_TITLE}}": tbc(data.get("contact_title")), "[CONTACT TITLE]": tbc(data.get("contact_title")),
        "{{CONTACT_EMAIL}}": tbc(data.get("contact_email")), "[CONTACT EMAIL]": tbc(data.get("contact_email")),
        "{{ADDRESS_1}}": tbc(data.get("address_1")), "[ADDRESS 1]": tbc(data.get("address_1")),
        "{{ADDRESS_2}}": clean(data.get("address_2")), "[ADDRESS 2]": clean(data.get("address_2")),
        "{{PHONE}}": tbc(data.get("phone")), "[PHONE]": tbc(data.get("phone")),
        "{{WEBSITE}}": tbc(data.get("website")), "[WEBSITE]": tbc(data.get("website")),
        "{{DATE}}": current_date,
        "{{ENGAGEMENT_TYPE}}": engagement_type, "[ENGAGEMENT TYPE]": engagement_type,
        "{{HR_SERVICE_TYPE}}": hr_service or "To Be Confirmed", "{{TA_SERVICE_TYPE}}": ta_service or "To Be Confirmed",
        "{{HOURLY_RATE}}": format_rate(data.get("hourly_rate", "")), "[HOURLY RATE]": format_rate(data.get("hourly_rate", "")),
        "{{ESTIMATED_HOURS}}": hours_value(data.get("estimated_hours", "")), "[ESTIMATED HOURS]": hours_value(data.get("estimated_hours", "")),
        "{{WEEKLY_COMMITMENT}}": hours_value(data.get("weekly_commitment", "")), "[WEEKLY COMMITMENT]": hours_value(data.get("weekly_commitment", "")),
        "{{MINIMUM_ENGAGEMENT}}": tbc(data.get("minimum_engagement", "")), "[MINIMUM ENGAGEMENT]": tbc(data.get("minimum_engagement", "")),
        "{{SCOPE_OF_WORK}}": combined_scope_from_data(data), "[SCOPE OF WORK]": combined_scope_from_data(data),
        "{{HR_SCOPE_OF_WORK}}": service_scope(
            hr_service,
            data.get("hr_scope_of_work", ""),
            data.get("estimated_hours", ""),
        ),
        "{{TA_SCOPE_OF_WORK}}": service_scope(
            ta_service,
            data.get("ta_scope_of_work", ""),
            data.get("estimated_hours", ""),
        ),
    }
