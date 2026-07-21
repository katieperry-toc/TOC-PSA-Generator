"""TOC PSA Generator — Streamlit application.

This is the UI layer only: page config, CSS, layout, form rendering,
session state, buttons, previews, and download buttons. It calls into
scoping.py (business rules / discovery-note parsing) and psa_builder.py
(document generation) rather than containing that logic itself.

Launch with:  python -m streamlit run app.py
"""

import html
import re
from typing import Dict

import streamlit as st

from config import (
    DEFAULT_SIGNER_KEY,
    FIELDS,
    GENERIC_NOTES_PLACEHOLDER,
    HR_SERVICES,
    LOGO_CANDIDATES,
    SIGNERS,
    TA_SERVICES,
    find_existing_file,
)
from psa_builder import build_psa
from scope_library import render_fallback_scope_text
from scoping import (
    agreement_label,
    clean,
    clean_multiline,
    combined_scope_from_data,
    detect_scope_conflicts,
    engagement_includes_hr,
    engagement_includes_ta,
    format_rate,
    generate_scope_for_service,
    normalize_commitment_text,
    normalize_engagement,
    normalize_hours_display,
    normalize_hr_service,
    minimum_row_value,
    normalize_ta_service,
    parse_discovery_notes,
)

SIGNER_OPTIONS = list(SIGNERS.keys())


def initialize_state() -> None:
    defaults = {
        "notes": "", "client_name": "", "contact_name": "", "contact_title": "", "contact_email": "",
        "address_1": "", "address_2": "", "phone": "", "website": "", "hourly_rate": "",
        "weekly_commitment": "", "minimum_engagement": "", "estimated_hours": "", "engagement_type": "",
        "hr_service_type": "HR Project Support", "ta_service_type": "Full Cycle Talent Acquisition Support",
        "hr_scope_of_work": "", "ta_scope_of_work": "", "generated_psa": None, "generated_filename": "",
        "generation_warnings": [], "engagement_type_selector": "",
        "signer": DEFAULT_SIGNER_KEY,
        # Tracks the last hr_service_type/ta_service_type the Scope of Work
        # field was auto-populated for, so the baseline scope only refreshes
        # when the selected service changes (or the field is empty) — never
        # overwriting a manual edit or an AI-personalized scope on rerun.
        "last_hr_service_for_scope": "", "last_ta_service_for_scope": "",
        "scope_warnings": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def apply_parsed_to_state(parsed: Dict[str, str], notes: str = "") -> None:
    engagement = normalize_engagement(
        parsed.get("engagement_type", ""),
        parsed.get("hr_service_type", ""),
        parsed.get("ta_service_type", ""),
    )

    for field in FIELDS:
        if field in {"service_type", "scope_of_work"}:
            continue
        st.session_state[field] = parsed.get(field, "")

    st.session_state.engagement_type = engagement

    st.session_state.hr_service_type = (
        normalize_hr_service(parsed.get("hr_service_type", ""))
        or "HR Project Support"
    )

    st.session_state.ta_service_type = (
        normalize_ta_service(parsed.get("ta_service_type", ""))
        or "Full Cycle Talent Acquisition Support"
    )

    # Scope of Work: the selected hr_service_type/ta_service_type (a
    # structured selection, set above) always determines which approved
    # TOC Scope Library template is used — discovery notes only personalize
    # the wording of that template, they never pick a different service.
    # generate_scope_for_service() always returns usable text (falling back
    # to the approved evergreen baseline when personalization isn't
    # possible or doesn't validate), and any internal warnings it raises
    # are surfaced to the user rather than silently swallowed.
    scope_warnings = detect_scope_conflicts(
        notes, st.session_state.hr_service_type, st.session_state.ta_service_type
    )
    scope_form_data = {
        "client_name": st.session_state.client_name,
        "engagement_type": st.session_state.engagement_type,
        "hr_service_type": st.session_state.hr_service_type,
        "ta_service_type": st.session_state.ta_service_type,
        "hourly_rate": st.session_state.hourly_rate,
        "weekly_commitment": st.session_state.weekly_commitment,
        "minimum_engagement": st.session_state.minimum_engagement,
        "estimated_hours": st.session_state.estimated_hours,
    }

    if engagement_includes_hr(engagement):
        hr_scope, hr_warnings = generate_scope_for_service(
            st.session_state.hr_service_type, scope_form_data, notes
        )
        st.session_state.hr_scope_of_work = hr_scope
        st.session_state.last_hr_service_for_scope = st.session_state.hr_service_type
        scope_warnings.extend(hr_warnings)

    if engagement_includes_ta(engagement):
        ta_scope, ta_warnings = generate_scope_for_service(
            st.session_state.ta_service_type, scope_form_data, notes
        )
        st.session_state.ta_scope_of_work = ta_scope
        st.session_state.last_ta_service_for_scope = st.session_state.ta_service_type
        scope_warnings.extend(ta_warnings)

    st.session_state.scope_warnings = scope_warnings

    st.session_state.generated_psa = None
    st.session_state.generated_filename = ""
    st.session_state.generation_warnings = []
    st.session_state.pop("preview_text", None)

def clear_form() -> None:
    for key in ["notes", "client_name", "contact_name", "contact_title", "contact_email", "address_1", "address_2", "phone", "website", "hourly_rate", "weekly_commitment", "minimum_engagement", "estimated_hours", "engagement_type", "hr_scope_of_work", "ta_scope_of_work"]:
        st.session_state[key] = ""
    st.session_state.engagement_type_selector = ""
    st.session_state.hr_service_type = "HR Project Support"
    st.session_state.ta_service_type = "Full Cycle Talent Acquisition Support"
    st.session_state.last_hr_service_for_scope = ""
    st.session_state.last_ta_service_for_scope = ""
    st.session_state.scope_warnings = []
    st.session_state.signer = DEFAULT_SIGNER_KEY
    st.session_state.generated_psa = None
    st.session_state.generated_filename = ""
    st.session_state.generation_warnings = []

def _maybe_autopopulate_baseline_scope(prefix: str) -> None:
    """Fill {prefix}_scope_of_work with the approved TOC Scope Library
    baseline for the currently selected service.

    Only acts when the selected hr_service_type/ta_service_type changed
    since the field was last populated, or the field is currently empty —
    never overwrites a manually edited scope or an AI-personalized scope
    produced by Read Notes on an unrelated rerun. Must be called before the
    corresponding st.text_area(key=f"{prefix}_scope_of_work") is
    rendered — Streamlit forbids writing to a widget's session_state key
    after that widget has already been instantiated in the same script run.
    """
    service_key = f"{prefix}_service_type"
    scope_key = f"{prefix}_scope_of_work"
    last_key = f"last_{prefix}_service_for_scope"

    selected = st.session_state.get(service_key, "")
    if not selected:
        return

    last_used = st.session_state.get(last_key, "")
    current_scope = clean_multiline(st.session_state.get(scope_key, ""))

    if selected != last_used or not current_scope:
        st.session_state[scope_key] = render_fallback_scope_text(selected)
        st.session_state[last_key] = selected

def safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]', "", clean(value) or "Client")
    return f"{re.sub(r'\s+', ' ', value).strip()} Professional Services Agreement.docx"

def collect_form_data() -> Dict[str, str]:
    engagement = clean(st.session_state.engagement_type)
    includes_hr = engagement_includes_hr(engagement)
    includes_ta = engagement_includes_ta(engagement)
    return {
        "client_name": clean(st.session_state.client_name),
        "contact_name": clean(st.session_state.contact_name),
        "contact_title": clean(st.session_state.contact_title),
        "contact_email": clean(st.session_state.contact_email),
        "address_1": clean(st.session_state.address_1),
        "address_2": clean(st.session_state.address_2),
        "phone": clean(st.session_state.phone),
        "website": clean(st.session_state.website),
        "hourly_rate": clean(st.session_state.hourly_rate),
        "weekly_commitment": clean(st.session_state.weekly_commitment),
        "minimum_engagement": clean(st.session_state.minimum_engagement),
        "estimated_hours": clean(st.session_state.estimated_hours),
        "engagement_type": engagement,
        "hr_service_type": clean(st.session_state.hr_service_type) if includes_hr else "",
        "ta_service_type": clean(st.session_state.ta_service_type) if includes_ta else "",
        "signer": clean(st.session_state.get("signer", "")),
       "hr_scope_of_work": (
    clean_multiline(st.session_state.hr_scope_of_work)
    or clean_multiline(st.session_state.get("hr_scope_from_notes", ""))
) if includes_hr else "",

"ta_scope_of_work": (
    clean_multiline(st.session_state.ta_scope_of_work)
    or clean_multiline(st.session_state.get("ta_scope_from_notes", ""))
) if includes_ta else "",
    }

def render_header() -> None:
    left, right = st.columns([4, 1])
    with left:
        st.markdown('<div class="toc-title">Professional Services Agreement Generator</div><div class="toc-subtitle">Turn discovery notes, an email, or an SOW into a client-ready TOC Professional Services Agreement.</div>', unsafe_allow_html=True)
    with right:
        logo_path = find_existing_file(LOGO_CANDIDATES)
        if logo_path:
            st.image(str(logo_path), use_container_width=True)

def render_notes_section() -> None:
    st.markdown('<div class="section-heading">1. Paste Your Notes</div>', unsafe_allow_html=True)
    st.text_area("Discovery notes, email, or SOW", key="notes", height=260, placeholder=GENERIC_NOTES_PLACEHOLDER, label_visibility="collapsed")

    st.radio(
        "TOC Signer",
        SIGNER_OPTIONS,
        key="signer",
        horizontal=True,
        help="Selects the TOC signer's name, title, and signature for this agreement.",
    )

    left, right, _ = st.columns([1.4, 1, 4])
    with left:
        read_notes = st.button("Read Notes", type="primary", use_container_width=True)
    with right:
        st.button("Clear", on_click=clear_form, use_container_width=True)
    if read_notes:
        try:
            with st.spinner("Reading notes and organizing the PSA details..."):
                notes_text = st.session_state.notes
                apply_parsed_to_state(parse_discovery_notes(notes_text), notes_text)
            st.success("Notes read successfully. Review the details below.")
            for warning in st.session_state.get("scope_warnings", []):
                st.warning(warning)
            st.rerun()
        except Exception as exc:
            st.error(f"Could not read the notes: {exc}")

def render_client_details() -> None:
    st.markdown('<div class="section-heading">2. Review Client Details</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: st.text_input("Client Name *", key="client_name", placeholder="Organization name")
    with c2: st.text_input("Primary Contact", key="contact_name", placeholder="Contact name")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Contact Title", key="contact_title", placeholder="Title")
    with c2: st.text_input("Contact Email", key="contact_email", placeholder="name@company.com")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Address Line 1", key="address_1", placeholder="Street address")
    with c2: st.text_input("Address Line 2", key="address_2", placeholder="City, State ZIP")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Phone", key="phone", placeholder="Phone number")
    with c2: st.text_input("Website", key="website", placeholder="www.company.com")

def render_pricing_details() -> None:
    st.markdown('<div class="section-heading">3. Confirm Pricing and Timing</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: st.text_input("Pricing or Hourly Rate *", key="hourly_rate", placeholder="Example: 175")
    with c2: st.text_input("Estimated Total Hours", key="estimated_hours", placeholder="Example: 60-75")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Weekly Commitment", key="weekly_commitment", placeholder="Example: 10-15 hours per week")
    with c2: st.text_input("Minimum Engagement", key="minimum_engagement", placeholder="Example: Minimum 40 hours")

def render_hr_fields() -> None:
    st.selectbox("HR Service Type", HR_SERVICES, key="hr_service_type")
    _maybe_autopopulate_baseline_scope("hr")
    st.text_area(
        "Human Resources Scope of Work",
        key="hr_scope_of_work",
        height=240,
        placeholder="Enter the HR services, responsibilities, deliverables, and timing.",
        help="Pre-filled from the approved TOC Scope Library for the selected service, and personalized from discovery notes after Read Notes. Edit freely.",
    )

def render_ta_fields() -> None:
    st.selectbox("Talent Acquisition Service Type", TA_SERVICES, key="ta_service_type")
    _maybe_autopopulate_baseline_scope("ta")
    st.text_area(
        "Talent Acquisition Scope of Work",
        key="ta_scope_of_work",
        height=240,
        placeholder="Enter the recruiting, sourcing, hiring, candidate management, or related support.",
        help="Pre-filled from the approved TOC Scope Library for the selected service, and personalized from discovery notes after Read Notes. Edit freely.",
    )

def render_service_details() -> None:
    st.markdown(
        '<div class="section-heading">4. Confirm Services and Scope</div>',
        unsafe_allow_html=True,
    )

    selected = normalize_engagement(
        st.session_state.get("engagement_type", ""),
        st.session_state.get("hr_service_type", ""),
        st.session_state.get("ta_service_type", ""),
    )

    # Do not ask the user to choose the engagement type. It is determined
    # from the notes when Read Notes is selected.
    st.session_state.engagement_type = selected

    if engagement_includes_hr(selected) and engagement_includes_ta(selected):
        c1, c2 = st.columns(2)
        with c1:
            render_hr_fields()
        with c2:
            render_ta_fields()
    elif engagement_includes_hr(selected):
        render_hr_fields()
    elif engagement_includes_ta(selected):
        render_ta_fields()
    else:
        st.info(
            "Paste the client notes above and select Read Notes. "
            "The app will identify the service type and display the correct scope fields."
        )

def render_generation_section() -> None:
    st.markdown(
        '<div class="section-heading">5. Review and Generate the PSA</div>',
        unsafe_allow_html=True,
    )

    # Keep this review experience stable: it is the final user checkpoint
    # before generation and should remain visible and editable upstream.
    review = collect_form_data()
    with st.expander("Final Review", expanded=True):
        left, right = st.columns(2)
        with left:
            st.write(f"**Client:** {review.get('client_name') or 'Not provided'}")
            st.write(f"**Primary Contact:** {review.get('contact_name') or 'Not provided'}")
            st.write(f"**Service:** {agreement_label(normalize_engagement(review.get('engagement_type', ''), review.get('hr_service_type', ''), review.get('ta_service_type', '')))}")
            st.write(f"**Rate:** {format_rate(review.get('hourly_rate', '')) if review.get('hourly_rate') else 'Not provided'}")
        with right:
            st.write(f"**Estimated Hours:** {normalize_hours_display(review.get('estimated_hours', '')) or 'Search-specific / not provided'}")
            st.write(f"**Weekly Commitment:** {normalize_commitment_text(review.get('weekly_commitment', '')) or 'To Be Confirmed'}")
            st.write(f"**Minimum Engagement:** {minimum_row_value(review.get('minimum_engagement', '')) or 'To Be Confirmed'}")

        scope_preview = combined_scope_from_data(review)

        st.markdown(
            """
            <div style="
                color: #17233c;
                font-family: Arial, Helvetica, sans-serif;
                font-size: 16px;
                font-weight: 700;
                margin-top: 12px;
                margin-bottom: 6px;
            ">
                Scope Preview:
            </div>
            """,
            unsafe_allow_html=True,
        )

        safe_scope_preview = html.escape(
            scope_preview or "No scope provided"
        ).replace("\n", "<br>")

        st.markdown(
            f"""
            <div style="
                color: #111827;
                background-color: #ffffff;
                font-family: Arial, Helvetica, sans-serif;
                font-size: 15px;
                font-weight: 400;
                line-height: 1.55;
                border: 1px solid #d9deea;
                border-radius: 8px;
                padding: 14px 16px;
            ">
                {safe_scope_preview}
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.caption(
            "Client Name and Pricing are required. Missing optional details will appear as To Be Confirmed."
        )

    for warning in st.session_state.get("scope_warnings", []):
        st.warning(warning)

    if st.button(
        "Generate Professional Services Agreement",
        type="primary",
        use_container_width=True,
    ):
        try:
            form_data = collect_form_data()
            with st.spinner("Creating the Professional Services Agreement..."):
                document_bytes, warnings = build_psa(form_data)
            st.session_state.generated_psa = document_bytes
            st.session_state.generated_filename = safe_filename(
                form_data["client_name"]
            )
            st.session_state.generation_warnings = warnings + st.session_state.get(
                "scope_warnings", []
            )
            st.success("The Professional Services Agreement is ready.")
        except Exception as exc:
            st.error(f"Could not generate the PSA: {exc}")

    if st.session_state.get("generated_psa"):
        st.download_button(
            "Download Professional Services Agreement",
            data=st.session_state.generated_psa,
            file_name=st.session_state.generated_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

        for warning in st.session_state.get("generation_warnings", []):
            st.warning(warning)

def apply_app_style() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: #f4f6fa;
        }
        .block-container {
            max-width: 1180px;
            padding-top: 2.5rem;
            padding-bottom: 4rem;
        }
        .toc-title {
            color: #17233c;
            font-size: 2.15rem;
            font-weight: 750;
            line-height: 1.15;
            margin-bottom: .45rem;
        }
        .toc-subtitle {
            color: #5e687a;
            font-size: 1.02rem;
            margin-bottom: 1.4rem;
        }
        .section-heading {
            color: #17233c;
            font-size: 1.55rem;
            font-weight: 750;
            border-bottom: 1px solid #d9deea;
            padding: 1.4rem 0 .75rem 0;
            margin-bottom: 1rem;
        }
        div[data-testid="stForm"],
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 14px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def main() -> None:
    st.set_page_config(
        page_title="TOC PSA Generator",
        page_icon="📄",
        layout="wide",
    )
    apply_app_style()
    initialize_state()
    render_header()
    render_notes_section()
    render_client_details()
    render_pricing_details()
    render_service_details()
    render_generation_section()


if __name__ == "__main__":
    main()
