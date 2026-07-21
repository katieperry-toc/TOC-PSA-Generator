"""Approved TOC Scope Library — reusable baseline content for PSA (and future
Addendum/SOW/Proposal) scope generation.

This module knows nothing about Word documents, Streamlit, or OpenAI —
scoping.py is responsible for turning this content plus discovery notes
into a final personalized scope, and psa_builder.py is responsible for
inserting that final text into the document.

Storage: all scope content lives in scope_library_data.json (next to this
file), not hardcoded in Python — so an administrator can add, edit,
archive, clone, categorize, and version scopes from the Scope Library Admin
page (pages/1_Scope_Library_Admin.py) without ever touching application
code. This module is the only thing that reads/writes that JSON file, and
every function below keeps working exactly as it did when the content was
a hardcoded dict, so psa_builder.py and scoping.py need no changes.

Two kinds of entries live in the same file:

  - The 5 "locked" services (locked: true) are the ones wired into the
    live generator: their service_key must exactly match
    config.HR_SERVICES / config.TA_SERVICES, because psa_builder.py maps
    each one to a specific, fixed checkbox row already printed on the
    approved Word template's Standard Services table. Their service_key,
    category, and locked flag cannot be changed from the admin page, and
    they cannot be archived or deleted — doing so would leave a template
    checkbox with no matching scope, or empty a service-type dropdown.
    Everything else about them (wording, bullets, exclusions, guardrails)
    is fully editable, and every edit is versioned.

  - "Draft" entries (locked: false) are created via Clone or Add New. They
    are stored and fully editable/versioned/archivable in the library, but
    are never surfaced in the live PSA generator's dropdowns, because that
    would require a matching new checkbox row on the approved Word
    template — a template change, which this system deliberately never
    does on its own. Promoting a draft to a live service is a deliberate,
    separate step (adding it to config.HR_SERVICES/TA_SERVICES and adding
    the matching row to the template) that only a developer should do.

Executive Search is intentionally not represented anywhere in this module.
It uses a separate agreement and must never be added here.

Content style (standardized-scope pass): every generated scope follows one
short, consistent, executive-friendly structure:

    [Service or Engagement Title]
    [One short paragraph — one or two sentences — describing the purpose
    and value of the engagement.]

    • [Primary service]
    • [Primary service]
    • [Primary service]

Roughly 90-120 words total. The Word template already renders a bold
"Primary Services" section heading directly above where this text is
inserted, so the generated text itself must never repeat that heading.
"""

import json
from copy import deepcopy
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, TypedDict

DATA_PATH = Path(__file__).resolve().parent / "scope_library_data.json"

# Fields that make up one version's editable content — used both to build
# a history snapshot and to know what counts as "content" vs. bookkeeping
# (status/version/history/last_updated) when comparing versions.
_CONTENT_FIELDS = [
    "display_title",
    "category",
    "overview",
    "primary_services",
    "optional_modules",
    "exclusions",
    "guardrails",
]


class ScopeTemplate(TypedDict):
    service_key: str
    display_title: str
    services_summary_category: str
    overview: str
    primary_services: List[str]
    optional_modules: List[str]
    exclusions: List[str]
    guardrails: List[str]


def _load_raw() -> Dict:
    if not DATA_PATH.exists():
        return {"schema_version": 1, "services": {}}
    with open(DATA_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save_raw(data: Dict) -> None:
    tmp_path = DATA_PATH.with_suffix(".json.tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    tmp_path.replace(DATA_PATH)


def _all_entries() -> Dict[str, Dict]:
    return _load_raw().get("services", {})


# ---------------------------------------------------------------------------
# Public API used by scoping.py / psa_builder.py — unchanged signatures from
# the previous hardcoded-dict version, so nothing downstream needs to change.
# ---------------------------------------------------------------------------

SCOPE_LIBRARY: Dict[str, ScopeTemplate] = {}  # populated below, kept for
# backward compatibility with any code that reads scope_library.SCOPE_LIBRARY
# directly. Always reflects the active, locked (live) services only.


def _refresh_scope_library_cache() -> None:
    global SCOPE_LIBRARY
    SCOPE_LIBRARY = {
        key: {
            "service_key": entry["service_key"],
            "display_title": entry["display_title"],
            "services_summary_category": entry["services_summary_category"],
            "overview": entry["overview"],
            "primary_services": entry["primary_services"],
            "optional_modules": entry["optional_modules"],
            "exclusions": entry["exclusions"],
            "guardrails": entry["guardrails"],
        }
        for key, entry in _all_entries().items()
        if entry.get("locked") and entry.get("status", "active") == "active"
    }


_refresh_scope_library_cache()


def get_scope_template(service_key: str) -> Optional[ScopeTemplate]:
    """Return the approved scope template for service_key, or None if unknown.

    Never raises — an unrecognized key (including Executive Search, which
    is never in this library) simply has no template. Always reads the
    live JSON file, so an edit saved from the admin page takes effect on
    the very next PSA generated, with no restart required.
    """
    _refresh_scope_library_cache()
    return SCOPE_LIBRARY.get(service_key)


def get_services_summary(service_key: str) -> str:
    """Return the high-level Services Summary category for service_key.

    Returns "" for an unrecognized key rather than raising, so a missing
    mapping shows as blank/To Be Confirmed upstream instead of crashing
    document generation.
    """
    template = get_scope_template(service_key)
    return template["services_summary_category"] if template else ""


def render_fallback_scope_text(service_key: str) -> str:
    """Render the evergreen baseline scope as plain text.

    Used when AI personalization is unavailable, fails, or a service has no
    discovery-notes context to personalize with. Produces the standardized
    structure: title, one short paragraph, a blank line, then 3 concise
    bullets. Deliberately excludes optional_modules — those are only
    appropriate when discovery notes support them, which a fallback-by-
    definition does not have.

    Does NOT include a "Primary Services" heading line: the Word template
    already renders that heading immediately above where this text is
    inserted, so repeating it here would show it twice on the page.
    """
    template = get_scope_template(service_key)
    if template is None:
        return ""

    lines = [template["display_title"], template["overview"], ""]
    lines.extend(f"• {item}" for item in template["primary_services"])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Admin API — used by pages/1_Scope_Library_Admin.py. Every write goes
# through save_scope(), which always versions (bumps version, snapshots the
# previous content into history, stamps last_updated) so nothing is ever
# silently overwritten without a recoverable trail.
# ---------------------------------------------------------------------------


class ScopeLockedError(Exception):
    """Raised when an admin action would affect a locked (live) service in
    a way that could break the generator or the approved Word template."""


def list_all_scopes() -> List[Dict]:
    """Return every scope entry (locked and draft, active and archived),
    sorted by category then display title, for the admin page's listing."""
    entries = _all_entries()
    return sorted(
        entries.values(),
        key=lambda e: (e.get("category", ""), e.get("display_title", "")),
    )


def get_scope_entry(service_key: str) -> Optional[Dict]:
    """Return the full admin-facing record (including status/version/
    history/locked) for one scope, or None if it doesn't exist."""
    return _all_entries().get(service_key)


def save_scope(service_key: str, content: Dict) -> Dict:
    """Save edited content for an existing scope, versioning the change.

    content should include the _CONTENT_FIELDS the admin page lets someone
    edit (display_title, category, overview, primary_services,
    optional_modules, exclusions, guardrails). service_key, locked, and
    status are never changed by this function — use rename-safe helpers
    below for those. Raises ScopeLockedError if service_key doesn't exist.
    """
    data = _load_raw()
    entries = data.setdefault("services", {})
    entry = entries.get(service_key)
    if entry is None:
        raise ScopeLockedError(f"'{service_key}' does not exist in the Scope Library.")

    previous_snapshot = {field: entry.get(field) for field in _CONTENT_FIELDS}
    previous_snapshot["version"] = entry.get("version", 1)
    previous_snapshot["last_updated"] = entry.get("last_updated", "")
    entry.setdefault("history", []).append(previous_snapshot)

    for field in _CONTENT_FIELDS:
        if field in content:
            entry[field] = content[field]

    # Locked entries must keep the Services Summary category text that the
    # rest of the document depends on in sync with the editable category.
    if entry.get("locked") and "category" in content:
        entry["services_summary_category"] = content["category"]

    entry["version"] = entry.get("version", 1) + 1
    entry["last_updated"] = date.today().isoformat()

    _save_raw(data)
    _refresh_scope_library_cache()
    return entry


def archive_scope(service_key: str) -> None:
    """Mark a scope archived (hidden from Clone-source pickers going
    forward, kept fully in the file for history). Refuses to archive a
    locked (live) service — that would silently break its dropdown/
    checkbox mapping in the live generator with no template change to
    match, so this is a hard stop rather than a soft warning.
    """
    data = _load_raw()
    entry = data.get("services", {}).get(service_key)
    if entry is None:
        raise ScopeLockedError(f"'{service_key}' does not exist in the Scope Library.")
    if entry.get("locked"):
        raise ScopeLockedError(
            f"'{service_key}' is a live service tied to the approved Word "
            "template's Standard Services table and cannot be archived."
        )
    entry["status"] = "archived"
    _save_raw(data)
    _refresh_scope_library_cache()


def restore_scope(service_key: str) -> None:
    """Restore a previously archived draft scope to active status."""
    data = _load_raw()
    entry = data.get("services", {}).get(service_key)
    if entry is None:
        raise ScopeLockedError(f"'{service_key}' does not exist in the Scope Library.")
    entry["status"] = "active"
    _save_raw(data)
    _refresh_scope_library_cache()


def clone_scope(source_service_key: str, new_display_title: str) -> str:
    """Duplicate an existing scope's content under a brand-new draft entry.

    The clone is always created as a draft (locked: false) — even when
    cloning one of the 5 live services — because giving it a new
    service_key means it has no matching checkbox row on the approved Word
    template yet. It's fully editable/versioned/archivable in the library,
    and can be promoted to live only via a deliberate code + template
    change, never automatically. Returns the new entry's service_key.
    """
    data = _load_raw()
    entries = data.setdefault("services", {})
    source = entries.get(source_service_key)
    if source is None:
        raise ScopeLockedError(f"'{source_service_key}' does not exist in the Scope Library.")

    new_key = new_display_title.strip()
    if not new_key:
        raise ValueError("Provide a name for the cloned scope.")
    base_key = new_key
    suffix = 2
    while new_key in entries:
        new_key = f"{base_key} ({suffix})"
        suffix += 1

    clone = deepcopy(source)
    clone["service_key"] = new_key
    clone["display_title"] = new_key
    clone["locked"] = False
    clone["status"] = "active"
    clone["version"] = 1
    clone["last_updated"] = date.today().isoformat()
    clone["history"] = []

    entries[new_key] = clone
    _save_raw(data)
    _refresh_scope_library_cache()
    return new_key


def add_new_scope(display_title: str, category: str) -> str:
    """Create a brand-new draft scope from a blank template. Always a
    draft (locked: false) — see clone_scope for why. Returns the new
    entry's service_key.
    """
    data = _load_raw()
    entries = data.setdefault("services", {})

    new_key = display_title.strip()
    if not new_key:
        raise ValueError("Provide a name for the new scope.")
    base_key = new_key
    suffix = 2
    while new_key in entries:
        new_key = f"{base_key} ({suffix})"
        suffix += 1

    entries[new_key] = {
        "service_key": new_key,
        "display_title": new_key,
        "services_summary_category": category,
        "category": category,
        "status": "active",
        "locked": False,
        "version": 1,
        "last_updated": date.today().isoformat(),
        "overview": "",
        "primary_services": [],
        "optional_modules": [],
        "exclusions": [],
        "guardrails": [],
        "history": [],
    }
    _save_raw(data)
    _refresh_scope_library_cache()
    return new_key


def delete_draft_scope(service_key: str) -> None:
    """Permanently remove a draft scope. Refuses to delete a locked
    (live) service under any circumstance."""
    data = _load_raw()
    entry = data.get("services", {}).get(service_key)
    if entry is None:
        return
    if entry.get("locked"):
        raise ScopeLockedError(
            f"'{service_key}' is a live service and cannot be deleted."
        )
    del data["services"][service_key]
    _save_raw(data)
    _refresh_scope_library_cache()
