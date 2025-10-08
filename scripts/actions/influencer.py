"""Utilities for building speaker/chair information from influencer data."""
from __future__ import annotations
from typing import Iterable, Iterator, List, Dict, Tuple


def iter_dicts(items: Iterable) -> Iterator[dict]:
    """Recursively yield dict objects from a nested list structure."""
    for item in items or []:
        if isinstance(item, dict):
            yield item
        elif isinstance(item, list):
            for sub in iter_dicts(item):
                yield sub


def build_profile(info: dict) -> str:
    """Compose bio text from influencer information.

    It stitches together organization, education and experience into a
    multi-line string suitable for template rendering.
    """
    parts: List[str] = []

    # Current position
    current = info.get("current_position")
    if isinstance(current, dict):
        org = current.get("organization")
        if org:
            parts.append(str(org))

    # Education
    edu = info.get("highest_education")
    if isinstance(edu, dict):
        school = edu.get("school")
        dept = edu.get("department")
        edu_line = " ".join(filter(None, [school, dept]))
        if edu_line:
            parts.append(edu_line)

    # Experience
    exp = info.get("experience")
    if isinstance(exp, list):
        for e in exp:
            if isinstance(e, dict):
                org = e.get("organization")
                title = e.get("title")
                line = " ".join(filter(None, [org, title]))
                if line:
                    parts.append(line)
            else:
                parts.append(str(e))
    elif isinstance(exp, str):
        parts.append(exp)

    # Achievements
    ach = info.get("achievements")
    if isinstance(ach, list):
        for a in ach:
            parts.append(str(a))
    elif isinstance(ach, str):
        parts.append(ach)

    return "\n".join(parts)


def build_profile_sections(info: dict) -> Dict[str, List[str]]:
    """Return structured profile sections with headings.

    The sections are returned in an ordered dictionary mapping a Chinese
    heading (e.g. ``"現職"``) to a list of lines. Only non-empty sections are
    included.
    """
    sections: Dict[str, List[str]] = {}

    # Current position
    current = info.get("current_position")
    if isinstance(current, dict):
        org = current.get("organization")
        title = current.get("title")
        line = " ".join(filter(None, [org, title]))
        if line:
            sections["現職"] = [line]

    # Education
    edu = info.get("highest_education")
    if isinstance(edu, dict):
        school = edu.get("school")
        dept = edu.get("department")
        line = " ".join(filter(None, [school, dept]))
        if line:
            sections["學歷"] = [line]

    # Experience
    exp = info.get("experience")
    if isinstance(exp, list) and exp:
        lines: List[str] = []
        for e in exp:
            if isinstance(e, dict):
                org = e.get("organization")
                title = e.get("title")
                line = " ".join(filter(None, [org, title]))
                if line:
                    lines.append(line)
            else:
                lines.append(str(e))
        if lines:
            sections["經歷"] = lines
    elif isinstance(exp, str) and exp:
        sections["經歷"] = [exp]

    # Achievements
    ach = info.get("achievements")
    if isinstance(ach, list) and ach:
        sections["成就"] = [str(a) for a in ach]
    elif isinstance(ach, str) and ach:
        sections["成就"] = [ach]

    # Specialties
    spec = info.get("specialties")
    if isinstance(spec, list) and spec:
        sections["專長"] = [str(s) for s in spec]
    elif isinstance(spec, str) and spec:
        sections["專長"] = [spec]

    return sections
def EN_build_profile_sections(info: dict) -> Dict[str, List[str]]:
    """Return structured profile sections with headings.

    The sections are returned in an ordered dictionary mapping a Chinese
    heading (e.g. ``"現職"``) to a list of lines. Only non-empty sections are
    included.
    """
    sections: Dict[str, List[str]] = {}

    # Current position
    current = info.get("current_position")
    if isinstance(current, dict):
        org = current.get("organization")
        title = current.get("title")
        line = " ".join(filter(None, [org, title]))
        if line:
            sections[""] = [line]

    # Education
    edu = info.get("highest_education")
    if isinstance(edu, dict):
        school = edu.get("school")
        dept = edu.get("department")
        line = " ".join(filter(None, [school, dept]))
        if line:
            sections["EDUCATION"] = [line]

    # Experience
    exp = info.get("experience")
    if isinstance(exp, list) and exp:
        lines: List[str] = []
        for e in exp:
            if isinstance(e, dict):
                org = e.get("organization")
                title = e.get("title")
                line = " ".join(filter(None, [org, title]))
                if line:
                    lines.append(line)
            else:
                lines.append(str(e))
        if lines:
            sections["PROFESSIONAL EXPERIENCE"] = lines
    elif isinstance(exp, str) and exp:
        sections["PROFESSIONAL EXPERIENCE"] = [exp]

    # Achievements
    ach = info.get("achievements")
    if isinstance(ach, list) and ach:
        sections["ACHIEVEMENTS"] = [str(a) for a in ach]
    elif isinstance(ach, str) and ach:
        sections["ACHIEVEMENTS"] = [ach]

    # Specialties
    spec = info.get("specialties")
    if isinstance(spec, list) and spec:
        sections["SPECIALTIES"] = [str(s) for s in spec]
    elif isinstance(spec, str) and spec:
        sections["SPECIALTIES"] = [spec]

    return sections


def build_people(program: dict, influencers: Iterable) -> Tuple[List[dict], List[dict]]:
    """Return (chairs, speakers) lists enriched with bio information."""
    infl_by_name: Dict[str, dict] = {p.get("name"): p for p in iter_dicts(influencers)}

    chairs: List[dict] = []
    speakers: List[dict] = []
    for entry in program.get("speakers", []) or []:
        name = entry.get("name")
        info = infl_by_name.get(name, {}) or {}
        enriched = {

            "name": name,


            "title": info.get("current_position", {}).get("title", "")
            if isinstance(info.get("current_position"), dict)
            else "",
            "short_title": info.get("short_title", {}).get("title", "")
            if isinstance(info.get("short_title"), dict)
            else "",
            "organization": info.get("current_position", {}).get("organization", "")

            if isinstance(info.get("current_position"), dict)
            else "",
            "profile": build_profile(info),
            "photo_url": info.get("photo_url", ""),
            # "profile_sections": build_profile_sections(info),
            "profile_sections": EN_build_profile_sections(info),
        }
        if entry.get("type") == "主持人":
            chairs.append(enriched)
        elif entry.get("type") == "講者":
            speakers.append(enriched)
    return chairs, speakers
