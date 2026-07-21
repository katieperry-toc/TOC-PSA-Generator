"""Shared configuration for the TOC document-generation platform.

This module owns everything that is pure data: file paths, template
locations, signature/consultant profiles, and shared constants. It must
never import from scoping.py, word_helpers.py, psa_builder.py, or
addendum_builder.py, so that every other module can safely import config
without risking a circular import.
"""

from pathlib import Path
from typing import Iterable

from docx.shared import RGBColor

APP_DIR = Path(__file__).resolve().parent


def find_existing_file(candidates: Iterable[Path]) -> Path | None:
    return next((path for path in candidates if path.exists()), None)


# ---------------------------------------------------------------------------
# Template registry
# ---------------------------------------------------------------------------
# Builders should look documents up by key (see TEMPLATES below) instead of
# hardcoding filenames. Each entry lists every filename this template has
# been saved/uploaded under, in preference order, so a rename on disk (as
# has already happened once with "Meghan PSA Template.docx" vs
# "Meghan PSA.docx") doesn't break anything.
#
# "Meghan PSA Template (v1.0).docx" was a branding/tone wording pass over
# the cover letter and intro page. It's been pulled out of the preference
# order (real Word reported font issues and page 2 merging into page 1
# that did not show up in this sandbox's LibreOffice-based rendering
# checks) until that's root-caused. TOC's approved, never-edited-by-code
# template is the source of truth again — nothing about its formatting
# should be touched.

TEMPLATE_CANDIDATES = [
    APP_DIR / "Meghan PSA Template.docx",
    APP_DIR / "Meghan PSA.docx",
]

ADDENDUM_TEMPLATE_CANDIDATES = [
    APP_DIR / "addendum_template.docx",
    APP_DIR / "Professional Services Agreement - 6.16.26.docx",
]

TEMPLATES = {
    "psa_default": TEMPLATE_CANDIDATES,
    "addendum": ADDENDUM_TEMPLATE_CANDIDATES,
}

LOGO_CANDIDATES = [
    # toc_logo_official.png was extracted directly from the approved PSA
    # Word template's own cover-page artwork (read-only; the template
    # itself was never modified) — it's the cleanest, highest-quality,
    # properly-padded copy of TOC's logo available, so it's preferred
    # over the other candidates below.
    APP_DIR / "toc_logo_official.png",
    APP_DIR / "logo-2024.png",
    APP_DIR / "toc_logo.png",
    APP_DIR / "TOC Logo.png",
    APP_DIR / "logo.png",
]

# ---------------------------------------------------------------------------
# TOC signer configuration — SINGLE SOURCE OF TRUTH
# ---------------------------------------------------------------------------
# Centralized so PSA and Addendum generation, and the Streamlit signer
# selector, all read from one place instead of hardcoding a person's
# name/title/image anywhere else. Confirmed byte-identical to the signature
# already embedded in the PSA template: meghan_popoleo_addendum.png is the
# same graphic psa_builder swaps in for every signer (Meghan included), so
# selecting Meg is a true no-op against today's template.

SIGNATURES_DIR = APP_DIR / "signatures"

SIGNERS = {
    "Meg": {
        "name": "Meghan Popoleo",
        "title": "President",
        "signature_path": SIGNATURES_DIR / "meghan_popoleo_addendum.png",
    },
    "Katie": {
        "name": "Katie Perry",
        "title": "Vice President of Operations",
        "signature_path": SIGNATURES_DIR / "katie_perry.png",
    },
    "Barb": {
        "name": "Barbara Rader",
        "title": "Director of Client Engagement and Growth",
        "signature_path": SIGNATURES_DIR / "barbara_rader.png",
    },
    "Marcia": {
        "name": "Marcia Zaruba O’Connor",
        "title": "CEO",
        "signature_path": SIGNATURES_DIR / "marcia_zaruba_oconnor.jpeg",
    },
}

# Meghan is the application's existing, already-established default signer
# (every generated PSA has always shown her name/title/signature) — preserved
# here rather than switched to Marcia, per "preserve the existing default if
# one has already been established."
DEFAULT_SIGNER_KEY = "Meg"

# Addendum builder still looks profiles up by full display name; derive that
# view from SIGNERS instead of maintaining a second copy of the same data.
SIGNATURE_PROFILES = {
    profile["name"]: {
        "image_path": profile["signature_path"],
        "name": profile["name"],
        "title": profile["title"],
    }
    for profile in SIGNERS.values()
}

DEFAULT_SIGNATURE_PROFILE = SIGNERS[DEFAULT_SIGNER_KEY]["name"]

# ---------------------------------------------------------------------------
# Shared constants
# ---------------------------------------------------------------------------

BLACK = RGBColor(0, 0, 0)

HR_SERVICES = [
    "HR Project Support",
    "Fractional/Interim HR Support",
    "HR Subscription",
]

TA_SERVICES = [
    "Full Cycle Talent Acquisition Support",
    "Sourcing Support",
]

ALL_SERVICES = HR_SERVICES + TA_SERVICES

# Scope of Services content lives in scope_library.py, keyed by these same
# HR_SERVICES/TA_SERVICES values — one taxonomy drives the Standard Services
# checkbox table, the Services Summary, and the Scope of Services alike.
# (The earlier separate 10-service HR_SCOPE_LIBRARY_SERVICES /
# TA_SCOPE_LIBRARY_SERVICES taxonomy has been retired in favor of this.)

ENGAGEMENT_HR = "Human Resources Support"
ENGAGEMENT_TA = "Talent Acquisition Support"
ENGAGEMENT_BOTH = "Human Resources & Talent Acquisition Support"

ENGAGEMENT_OPTIONS = ["", ENGAGEMENT_HR, ENGAGEMENT_TA, ENGAGEMENT_BOTH]

GENERIC_NOTES_PLACEHOLDER = """Paste discovery notes, an email, an SOW, or meeting notes here.

Helpful details may include:
• Client name and primary contact
• Pricing or hourly rate
• Human Resources, Talent Acquisition, or combined support
• Specific service requested
• Estimated total hours
• Weekly commitment and minimum engagement
• Scope of work
• Address, email, phone, and website

Only Client Name and Pricing are required."""

FIELDS = [
    "client_name", "contact_name", "contact_title", "contact_email",
    "address_1", "address_2", "phone", "website", "hourly_rate",
    "weekly_commitment", "minimum_engagement", "estimated_hours",
    "engagement_type", "hr_service_type", "ta_service_type",
    "hr_scope_of_work", "ta_scope_of_work", "service_type", "scope_of_work",
]
