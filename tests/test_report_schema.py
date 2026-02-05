import json
from datetime import UTC, datetime

from excelbench.cli import report


def _base_results():
    return {
        "metadata": {
            "benchmark_version": "0.1.0",
            "run_date": datetime.now(UTC).isoformat(),
            "excel_version": "test",
            "platform": "test",
        },
        "libraries": {
            "openpyxl": {
                "name": "openpyxl",
                "version": "3.1.0",
                "language": "python",
                "capabilities": ["read", "write"],
            }
        },
    }


def test_report_new_schema(tmp_path):
    data = _base_results()
    data["results"] = [
        {
            "feature": "cell_values",
            "library": "openpyxl",
            "scores": {"read": 3, "write": 2},
            "test_cases": {
                "case1": {
                    "read": {
                        "passed": True,
                        "expected": {"type": "string"},
                        "actual": {"type": "string"},
                    },
                    "write": {
                        "passed": True,
                        "expected": {"type": "string"},
                        "actual": {"type": "string"},
                    },
                }
            },
            "notes": None,
        }
    ]

    results_path = tmp_path / "results.json"
    results_path.write_text(json.dumps(data))

    output_dir = tmp_path / "out"
    report(results_path=results_path, output_dir=output_dir)

    assert (output_dir / "README.md").exists()
    assert (output_dir / "matrix.csv").exists()


def test_report_legacy_schema(tmp_path):
    data = _base_results()
    data["results"] = [
        {
            "feature": "cell_values",
            "library": "openpyxl",
            "scores": {"read": 3, "write": None},
            "test_cases": {
                "case1": {
                    "passed": True,
                    "expected": {"type": "string"},
                    "actual": {"type": "string"},
                }
            },
            "notes": None,
        }
    ]

    results_path = tmp_path / "results.json"
    results_path.write_text(json.dumps(data))

    output_dir = tmp_path / "out"
    report(results_path=results_path, output_dir=output_dir)

    assert (output_dir / "README.md").exists()
    assert (output_dir / "matrix.csv").exists()
