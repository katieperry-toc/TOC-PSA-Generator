"""Scope Library Admin — manage TOC's approved scope content without code.

This is a separate Streamlit page (Streamlit auto-discovers anything in a
pages/ folder and adds it to the app's sidebar navigation) so day-to-day
PSA generation in app.py stays completely separate from library
management. Everything here reads and writes scope_library_data.json
through scope_library.py's admin functions — it never touches the Word
template, psa_builder.py, or any document-generation code.
"""

import streamlit as st

import scope_library as sl

st.set_page_config(page_title="Scope Library Admin", page_icon="🗂️", layout="wide")

st.markdown(
    """
    <style>
    .stApp { background: #eaf1fb; }
    .block-container { max-width: 1100px; padding-top: 2rem; }
    div[data-testid="stExpander"] { background: #ffffff; border: 1px solid #d9deea; border-radius: 12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Scope Library Admin")
st.caption(
    "Add, edit, archive, clone, categorize, and version TOC's approved "
    "Scope of Services content — no code changes required. This never "
    "touches the Word template or any document-generation logic."
)

st.info(
    "The 5 services marked 🔒 Live are tied to a fixed checkbox row already "
    "printed on the approved Word template, so their name and category are "
    "locked and they can't be archived or deleted — but every word of "
    "their content is fully editable, and every save is versioned. Clone "
    "or Add New to create draft variants, which can be freely edited, "
    "archived, restored, or deleted; promoting a draft to a live service "
    "requires a separate, deliberate code + template change."
)

def _lines_to_list(text: str) -> list:
    return [line.strip() for line in text.splitlines() if line.strip()]

def _list_to_lines(items: list) -> str:
    return "\n".join(items)

scopes = sl.list_all_scopes()
categories = sorted({s.get("category", "") for s in scopes if s.get("category")})

tab_manage, tab_new = st.tabs(["Manage Scopes", "Add New Scope"])

with tab_manage:
    for category in categories:
        st.subheader(category)
        category_scopes = [s for s in scopes if s.get("category") == category]
        for entry in category_scopes:
            key = entry["service_key"]
            locked = bool(entry.get("locked"))
            status = entry.get("status", "active")
            badge = "🔒 Live" if locked else ("📦 Archived" if status == "archived" else "📝 Draft")
            with st.expander(f"{entry['display_title']}  —  {badge}  (v{entry.get('version', 1)})"):
                st.caption(f"Service key: `{key}` · Last updated: {entry.get('last_updated', '—')}")

                with st.form(key=f"form_{key}"):
                    display_title = st.text_input(
                        "Display title", value=entry.get("display_title", ""),
                        disabled=locked,
                        help="Locked for live services — this name is tied to the template's checkbox table." if locked else None,
                    )
                    category_value = st.text_input(
                        "Category", value=entry.get("category", ""), disabled=locked,
                    )
                    overview = st.text_area(
                        "Overview (one short paragraph)", value=entry.get("overview", ""), height=90,
                    )
                    primary_services = st.text_area(
                        "Primary Services menu (one per line — the AI chooses 3-4 of these per PSA)",
                        value=_list_to_lines(entry.get("primary_services", [])), height=110,
                    )
                    optional_modules = st.text_area(
                        "Optional modules (one per line — used only as an occasional 4th bullet)",
                        value=_list_to_lines(entry.get("optional_modules", [])), height=90,
                    )
                    exclusions = st.text_area(
                        "Exclusions (one per line)",
                        value=_list_to_lines(entry.get("exclusions", [])), height=70,
                    )
                    guardrails = st.text_area(
                        "Guardrails / AI instructions (one per line)",
                        value=_list_to_lines(entry.get("guardrails", [])), height=90,
                    )

                    saved = st.form_submit_button("Save New Version", type="primary")
                    if saved:
                        content = {
                            "overview": overview.strip(),
                            "primary_services": _lines_to_list(primary_services),
                            "optional_modules": _lines_to_list(optional_modules),
                            "exclusions": _lines_to_list(exclusions),
                            "guardrails": _lines_to_list(guardrails),
                        }
                        if not locked:
                            content["display_title"] = display_title.strip() or entry["display_title"]
                            content["category"] = category_value.strip() or entry["category"]
                        updated = sl.save_scope(key, content)
                        st.success(f"Saved as version {updated['version']}.")
                        st.rerun()

                history = entry.get("history", [])
                if history:
                    with st.popover(f"Version history ({len(history)} prior version{'s' if len(history) != 1 else ''})"):
                        for past in reversed(history):
                            st.markdown(f"**v{past.get('version')}** — {past.get('last_updated', '—')}")
                            st.text(past.get("overview", ""))
                            for item in past.get("primary_services", []):
                                st.text(f"  • {item}")
                            st.divider()

                action_cols = st.columns(4)
                with action_cols[0]:
                    clone_name = st.text_input("New name for clone", key=f"clone_name_{key}", placeholder="e.g. HR Project Support (EU)")
                with action_cols[1]:
                    st.write("")
                    st.write("")
                    if st.button("Clone", key=f"clone_btn_{key}", use_container_width=True):
                        if clone_name.strip():
                            new_key = sl.clone_scope(key, clone_name.strip())
                            st.success(f"Cloned as draft: {new_key}")
                            st.rerun()
                        else:
                            st.warning("Enter a name for the clone first.")
                with action_cols[2]:
                    st.write("")
                    st.write("")
                    if not locked and status == "active":
                        if st.button("Archive", key=f"archive_btn_{key}", use_container_width=True):
                            sl.archive_scope(key)
                            st.rerun()
                    elif not locked and status == "archived":
                        if st.button("Restore", key=f"restore_btn_{key}", use_container_width=True):
                            sl.restore_scope(key)
                            st.rerun()
                with action_cols[3]:
                    st.write("")
                    st.write("")
                    if not locked:
                        if st.button("Delete Draft", key=f"delete_btn_{key}", use_container_width=True):
                            sl.delete_draft_scope(key)
                            st.rerun()

with tab_new:
    st.write("Create a brand-new draft scope from scratch.")
    with st.form("add_new_scope_form"):
        new_title = st.text_input("Display title", placeholder="e.g. Payroll Advisory Support")
        new_category = st.selectbox(
            "Category",
            options=categories + ["Other"],
        )
        if new_category == "Other":
            new_category = st.text_input("Custom category name")
        submitted = st.form_submit_button("Create Draft", type="primary")
        if submitted:
            if not new_title.strip():
                st.error("Enter a display title.")
            else:
                new_key = sl.add_new_scope(new_title.strip(), new_category.strip())
                st.success(f"Created draft: {new_key}. Edit it under Manage Scopes.")
                st.rerun()

    st.caption(
        "New scopes are always created as drafts. To make one available in "
        "the live PSA generator, a developer needs to add its service_key "
        "to config.HR_SERVICES or config.TA_SERVICES and add a matching "
        "checkbox row to the approved Word template — this page never does "
        "that automatically, to keep the template protected."
    )
