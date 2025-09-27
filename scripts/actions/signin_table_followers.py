"""Helpers to transform follower records into sign-in table rows."""

from __future__ import annotations

from typing import Iterable, List

from scripts.actions.follower import FollowerRecord
from scripts.actions.signin_table_render import SignInRow

__all__ = [
    "follower_to_signin_row",
    "followers_to_signin_rows",
]


def follower_to_signin_row(follower: FollowerRecord) -> SignInRow:
    """Convert a :class:`FollowerRecord` into a :class:`SignInRow`."""

    topic = (follower.company_department or "").strip()
    if not topic:
        topic = (follower.registration_no or "").strip()
    return SignInRow(
        topic=topic,
        name=(follower.chinese_name or "").strip(),
        title=(follower.title or "").strip(),
        organization=(follower.service_unit_category or "").strip(),
    )


def followers_to_signin_rows(followers: Iterable[FollowerRecord]) -> List[SignInRow]:
    """Convert an iterable of followers into sign-in table rows."""

    return [follower_to_signin_row(follower) for follower in followers]
