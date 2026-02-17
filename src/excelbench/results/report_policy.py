"""Shared reporting policy for public-facing outputs."""

from __future__ import annotations

from typing import Any

_HIDDEN_LIBRARIES = frozenset({"pyumya"})

# Public-facing modify semantics.
_MODIFY_BY_LIBRARY: dict[str, str] = {
    "wolfxl": "Patch",
    "openpyxl": "Rewrite",
    "pandas": "Rebuild",
    "pyexcel": "Rebuild",
    "tablib": "Rebuild",
    "pylightxl": "Rebuild",
}


def is_visible_library(name: str) -> bool:
    """Return True when a library should appear in public report outputs."""
    return name not in _HIDDEN_LIBRARIES


def modify_mode_label(library: str, capabilities: set[str] | list[str]) -> str:
    """Return a user-facing modify capability label for a library."""
    caps = set(capabilities)
    if "read" not in caps or "write" not in caps:
        return "No"
    return _MODIFY_BY_LIBRARY.get(library, "Rewrite")


def filter_report_data(data: dict[str, Any]) -> dict[str, Any]:
    """Return a copy of report JSON data with hidden libraries removed."""
    out: dict[str, Any] = dict(data)

    libraries = data.get("libraries")
    if isinstance(libraries, dict):
        out["libraries"] = {
            name: info for name, info in libraries.items() if is_visible_library(name)
        }

    results = data.get("results")
    if isinstance(results, list):
        out["results"] = [
            entry
            for entry in results
            if isinstance(entry, dict) and is_visible_library(str(entry.get("library", "")))
        ]

    return out


def filter_memory_rows(memory: list[dict[str, Any]] | None) -> list[dict[str, Any]] | None:
    """Filter memory profile rows using adapter/library visibility policy."""
    if memory is None:
        return None
    return [
        row
        for row in memory
        if is_visible_library(str(row.get("adapter", "")))
    ]
