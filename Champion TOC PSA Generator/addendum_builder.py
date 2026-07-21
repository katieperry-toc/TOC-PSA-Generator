"""Professional Services Agreement Addendum generation.

This module owns Addendum generation the same way psa_builder.py owns PSA
generation. It is a deliberate scaffold, not a full implementation: the
Addendum's field population, table updates, and signature placement logic
still need to be written. What IS real and working here:

* Template loading via the config.TEMPLATES registry (no hardcoded
  filenames).
* A structured-data entry point (collect_addendum_data) so app.py has a
  single, stable place to hand off form input, mirroring
  scoping.build_replacements for the PSA.
* Signer lookup via config.SIGNATURE_PROFILES, so Barbara Rader / Katie
  Perry / Marcia Zaruba O'Connor / Meghan Popoleo can each sign an
  Addendum without hardcoding a single person's name/title/image.
* build_addendum() end-to-end: it loads the real template and returns real
  document bytes today, so the Streamlit app can wire up an "Addendum"
  tab/button now. It does not yet populate any fields — it returns the
  blank template with an explicit warning saying so, rather than silently
  pretending to be finished.

Do not mix PSA-specific logic into this module, and do not mix Addendum
logic into psa_builder.py — future document types (Proposal, SOW, Change
Order) should each get their own builder module following this same shape.
"""

import io
from typing import Dict, List, Tuple

from docx import Document

from config import (
    ADDENDUM_TEMPLATE_CANDIDATES,
    DEFAULT_SIGNATURE_PROFILE,
    SIGNATURE_PROFILES,
    find_existing_file,
)
from scoping import clean

# Fields the Addendum form will eventually collect. Extend this as the
# Streamlit Addendum form is built out — see collect_addendum_data below.
ADDENDUM_FIELDS = [
    "client_name",
    "contact_name",
    "contact_title",
    "address_1",
    "address_2",
    "original_psa_date",
    "scope_changes",
    "signer_profile",  # key into config.SIGNATURE_PROFILES
]


def load_addendum_template() -> Document:
    """Load the blank Addendum template.

    Raises FileNotFoundError if none of config.ADDENDUM_TEMPLATE_CANDIDATES
    exist on disk, the same pattern psa_builder.build_psa uses for the PSA
    template via config.TEMPLATE_CANDIDATES.
    """
    template_path = find_existing_file(ADDENDUM_TEMPLATE_CANDIDATES)
    if template_path is None:
        raise FileNotFoundError(
            "No Addendum template found. Expected one of: "
            + ", ".join(str(path) for path in ADDENDUM_TEMPLATE_CANDIDATES)
        )
    return Document(template_path)


def collect_addendum_data(raw: Dict[str, str]) -> Dict[str, str]:
    """Normalize raw form input into the structured shape build_addendum expects.

    Placeholder: today this just cleans whitespace on every known field.
    As Addendum-specific business rules are added (e.g. validating
    original_psa_date, generating scope-change language), add them here —
    mirroring scoping.build_replacements for the PSA — rather than in
    app.py or build_addendum itself, so this module stays the single
    source of truth for Addendum business logic.
    """
    return {field: clean(raw.get(field, "")) for field in ADDENDUM_FIELDS}


def resolve_signer(profile_name: str) -> Dict[str, str]:
    """Look up a signer's name/title/signature image path.

    Falls back to config.DEFAULT_SIGNATURE_PROFILE if profile_name is
    missing or unrecognized, rather than raising, so a not-yet-configured
    signer never blocks document generation.
    """
    return SIGNATURE_PROFILES.get(profile_name, SIGNATURE_PROFILES[DEFAULT_SIGNATURE_PROFILE])


def build_addendum(data: Dict[str, str]) -> Tuple[bytes, List[str]]:
    """Generate an Addendum document from structured data.

    NOT YET IMPLEMENTED beyond template loading: this currently returns
    the blank template's own bytes unmodified, with a warning saying so.
    Field population, table updates, and signature placement (following
    the same patterns as psa_builder.populate_template_specific_fields /
    psa_builder.move_cover_letter_signature) belong here once the
    Addendum's field set and signer-selection UI are finalized.
    """
    warnings: List[str] = [
        "Addendum generation is not yet implemented — returning the blank template unmodified.",
    ]

    doc = load_addendum_template()

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), warnings
