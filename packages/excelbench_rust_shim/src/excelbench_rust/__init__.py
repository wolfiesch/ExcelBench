"""Compatibility shim for legacy `excelbench_rust` imports.

This package re-exports the WolfXL native extension module (`wolfxl._rust`).

Rationale:
- WolfXL's native module is branded as `wolfxl._rust`.
- Existing code (including older ExcelBench integrations) may still import `excelbench_rust`.

This shim lets both coexist without requiring a big-bang rename for downstream users.
"""

from __future__ import annotations

from types import ModuleType


def _load_impl() -> ModuleType:
    try:
        from wolfxl import _rust as impl  # type: ignore[attr-defined]
    except Exception as e:  # pragma: no cover
        raise ImportError(
            "wolfxl._rust is not available. Install wolfxl-rust (native wheels) or build from source."
        ) from e
    return impl


_impl = _load_impl()

# Re-export all public attributes.
__all__ = []
for _name in dir(_impl):
    if _name.startswith("_"):
        continue
    globals()[_name] = getattr(_impl, _name)
    __all__.append(_name)
