import json
from datetime import UTC, datetime

from excelbench.generator.generate import load_manifest
from excelbench.models import Importance


def test_manifest_load_defaults(tmp_path):
    data = {
        "generated_at": datetime.now(UTC).isoformat(),
        "excel_version": "test",
        "generator_version": "1.0.0",
        "files": [
            {
                "path": "tier1/01_cell_values.xlsx",
                "feature": "cell_values",
                "tier": 1,
                "test_cases": [
                    {"id": "case1", "label": "Case 1", "row": 2, "expected": {"type": "string"}}
                ],
            }
        ],
    }

    manifest_path = tmp_path / "manifest.json"
    manifest_path.write_text(json.dumps(data))

    manifest = load_manifest(manifest_path)
    tc = manifest.files[0].test_cases[0]
    assert tc.sheet is None
    assert tc.cell is None
    assert tc.importance == Importance.BASIC


def test_manifest_load_explicit_fields(tmp_path):
    data = {
        "generated_at": datetime.now(UTC).isoformat(),
        "excel_version": "test",
        "generator_version": "1.0.0",
        "files": [
            {
                "path": "tier1/09_multiple_sheets.xlsx",
                "feature": "multiple_sheets",
                "tier": 1,
                "test_cases": [
                    {
                        "id": "case1",
                        "label": "Case 1",
                        "row": 2,
                        "expected": {"type": "string"},
                        "sheet": "Alpha",
                        "cell": "B2",
                        "importance": "edge",
                    }
                ],
            }
        ],
    }

    manifest_path = tmp_path / "manifest.json"
    manifest_path.write_text(json.dumps(data))

    manifest = load_manifest(manifest_path)
    tc = manifest.files[0].test_cases[0]
    assert tc.sheet == "Alpha"
    assert tc.cell == "B2"
    assert tc.importance == Importance.EDGE
