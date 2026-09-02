# kanban_ui.py
"""
Drag-and-drop Kanban UI for Manual Group Alignment, plus per-student
info popovers.

DESIGN NOTE:
streamlit-sortables renders draggable tiles as plain text labels -- it
wraps SortableJS, which does not support interactive child widgets (like
a button) inside a draggable element, since click/drag events collide.
So this is split into two panels:

  1. DRAG-AND-DROP BOARD (top): compact text tiles, dragged between group
     columns. A medical-flag icon appears directly on the tile so
     something safety-relevant is visible at a glance, without needing
     a click.
  2. STUDENT DETAILS PANEL (below): one row per student, grouped the
     same way as the board, each with an (i) popover showing friend
     requests, form data, and medical flags in full.

If inline (i) buttons ON the draggable tiles themselves are a hard
requirement, that needs streamlit-elements + react-beautiful-dnd instead
of streamlit-sortables -- happy to swap it in, but it's a heavier and
less actively maintained dependency. Flagging now before this goes
further into the build.

Requires: pip install streamlit-sortables
"""

import streamlit as st
from streamlit_sortables import sort_items

MEDICAL_FLAG_ICON = "⚕️"
FRIEND_REQUEST_ICON = "🤝"


def _tile_label(student_id: str, students: dict) -> str:
    """Build a draggable tile's display text: name + medical-flag icon if relevant."""
    info = students.get(student_id, {})
    name = info.get("name", student_id)
    flags = info.get("medical_flags", [])
    icon = f"{MEDICAL_FLAG_ICON} " if flags else ""
    return f"{icon}{name}"


def _build_label_lookup(students: dict) -> dict:
    """Map every possible tile label -> student_id, built fresh each render (cheap, avoids stale state)."""
    lookup = {}
    for sid, info in students.items():
        name = info.get("name", sid)
        lookup[name] = sid                              # unflagged label
        lookup[f"{MEDICAL_FLAG_ICON} {name}"] = sid      # flagged label
    return lookup


def render_kanban_board(groups: dict, students: dict) -> dict:
    """
    Render the drag-and-drop board.

    groups:   { "Group Name": ["S001", "S002", ...], ... }
    students: { "S001": {"name": ..., "friend_requests": [...], ...}, ... }

    Returns an UPDATED groups dict reflecting the user's drag-and-drop
    moves in this render. The caller is responsible for diffing this
    against session state and deciding when to persist it via camp_sync.
    """
    if not groups:
        st.info("No groups yet. Add groups below before assigning students.")
        return groups

    label_lookup = _build_label_lookup(students)

    # streamlit-sortables (multi_containers=True) expects:
    #   [{"header": str, "items": [str, ...]}, ...]
    # and returns the same shape back, reordered/moved per the user's drag.
    containers = [
        {
            "header": f"{group_name}  ({len(student_ids)})",
            "items": [_tile_label(sid, students) for sid in student_ids],
        }
        for group_name, student_ids in groups.items()
    ]

    sorted_containers = sort_items(
        containers,
        multi_containers=True,
        direction="vertical",
        key="camp_kanban_board",
    )

    # Translate labels back into student IDs. Containers come back in the
    # same order they were sent, so zip against the original group names.
    updated_groups = {}
    for original_name, container in zip(groups.keys(), sorted_containers):
        ids = [label_lookup[label] for label in container["items"] if label in label_lookup]
        updated_groups[original_name] = ids

    return updated_groups


def render_student_details_panel(students: dict, groups: dict):
    """
    Companion panel: one row per student with an (i) popover showing
    friend requests, raw form data, and medical flags.
    Organized by current group so staff can scan roster + details together.
    """
    st.markdown("#### Student Details")
    st.caption("Click the ⓘ next to a student to see friend requests, form answers, and medical flags.")

    for group_name, student_ids in groups.items():
        if not student_ids:
            continue
        with st.expander(f"{group_name} — {len(student_ids)} students", expanded=False):
            for sid in student_ids:
                info = students.get(sid, {})
                name = info.get("name", sid)
                flags = info.get("medical_flags", [])

                col_name, col_info = st.columns([6, 1])
                with col_name:
                    label = f"{MEDICAL_FLAG_ICON} **{name}**" if flags else f"**{name}**"
                    st.markdown(label)
                with col_info:
                    with st.popover("ⓘ", use_container_width=True):
                        _render_student_popover_content(sid, info)


def _render_student_popover_content(student_id: str, info: dict):
    """Content shown inside a student's (i) popover."""
    st.markdown(f"**{info.get('name', student_id)}**  \n`ID: {student_id}`")
    st.markdown("---")

    st.markdown(f"{FRIEND_REQUEST_ICON} **Friend Requests**")
    friend_requests = info.get("friend_requests", [])
    if friend_requests:
        for f in friend_requests:
            st.markdown(f"- {f}")
    else:
        st.caption("No friend requests submitted.")

    st.markdown("---")

    flags = info.get("medical_flags", [])
    if flags:
        st.markdown(f"{MEDICAL_FLAG_ICON} **Medical Flags**")
        for flag in flags:
            st.error(flag, icon="⚕️")
    else:
        st.caption("No medical flags on record.")

    form_data = info.get("form_data", {})
    if form_data:
        st.markdown("---")
        st.markdown("**Form Data**")
        for key, value in form_data.items():
            if value not in (None, "", "nan"):
                st.markdown(f"- **{key}:** {value}")


def render_group_management(groups: dict) -> dict:
    """
    Small UI for adding/removing groups (columns on the board).
    Returns the updated groups dict. Removing a non-empty group is blocked
    until it's emptied, to avoid silently losing student assignments.
    """
    with st.expander("⚙️ Manage Groups", expanded=False):
        col1, col2 = st.columns([3, 1])
        with col1:
            new_group_name = st.text_input(
                "New group name",
                key="new_group_name_input",
                label_visibility="collapsed",
                placeholder="e.g. Bay of Fires",
            )
        with col2:
            if st.button("➕ Add Group", use_container_width=True):
                name = new_group_name.strip()
                if name and name not in groups:
                    groups[name] = []
                elif name in groups:
                    st.warning(f"'{name}' already exists.")

        if groups:
            st.markdown("**Remove a group:**")
            for group_name in list(groups.keys()):
                gcol1, gcol2 = st.columns([4, 1])
                gcol1.markdown(f"- {group_name} ({len(groups[group_name])} students)")
                if gcol2.button("🗑️", key=f"remove_{group_name}"):
                    if groups[group_name]:
                        st.warning(
                            f"'{group_name}' still has {len(groups[group_name])} students. "
                            f"Move them out before deleting."
                        )
                    else:
                        del groups[group_name]
                        st.rerun()

    return groups