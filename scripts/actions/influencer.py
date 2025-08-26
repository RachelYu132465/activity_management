"""Utilities for building speaker/chair information from influencer data."""
from __future__ import annotations
from typing import Iterable, Iterator, List, Dict, Tuple


def iter_dicts(items: Iterable) -> Iterator[dict]:
    """Recursively yield dict objects from a nested list structure.

    The influencer JSON mixes lists and objects; this helper flattens the
    structure so each influencer record can be processed uniformly.
    """
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
    current = info.get("current_position")
    if isinstance(current, dict):
        org = current.get("organization")
        if org:
            parts.append(org)
    edu = info.get("highest_education")
    if isinstance(edu, dict):
        school = edu.get("school")
        dept = edu.get("department")
        edu_line = " ".join(filter(None, [school, dept]))
        if edu_line:
            parts.append(edu_line)
    exp = info.get("experience")
    if isinstance(exp, list):
        parts.extend(exp)
    elif isinstance(exp, str):
        parts.append(exp)
    return "\n".join(parts)


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
            "profile": build_profile(info),
            "photo_url": info.get("photo_url", ""),
        }
        if entry.get("type") == "主持人":
            chairs.append(enriched)
        else:
            speakers.append(enriched)
    return chairs, speakers
