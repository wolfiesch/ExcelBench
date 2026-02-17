from excelbench.results.dashboard import _build_dashboard


def test_dashboard_includes_best_adapter_by_workload_profile() -> None:
    fidelity = {
        "metadata": {"profile": "xlsx", "run_date": "2026-01-01T00:00:00Z"},
        "libraries": {
            "openpyxl": {"capabilities": ["read", "write"]},
            "xlsxwriter": {"capabilities": ["write"]},
        },
        "results": [],
    }
    perf = {
        "results": [
            {
                "feature": "cell_values_1k",
                "library": "openpyxl",
                "workload_size": "small",
                "perf": {
                    "read": {"op_count": 1000, "wall_ms": {"p50": 10.0}},
                    "write": {"op_count": 1000, "wall_ms": {"p50": 25.0}},
                },
            },
            {
                "feature": "cell_values_10k_bulk_write",
                "library": "xlsxwriter",
                "workload_size": "medium",
                "perf": {
                    "read": None,
                    "write": {"op_count": 10000, "wall_ms": {"p50": 20.0}},
                },
            },
            {
                "feature": "cell_values_100k_bulk_read",
                "library": "openpyxl",
                "workload_size": "large",
                "perf": {
                    "read": {"op_count": 100000, "wall_ms": {"p50": 250.0}},
                    "write": None,
                },
            },
        ]
    }

    lines = _build_dashboard(fidelity, perf)
    doc = "\n".join(lines)

    assert "## Best Adapter by Workload Profile" in doc
    assert "| small |" in doc
    assert "| medium |" in doc
    assert "| large |" in doc


def test_dashboard_filters_pyumya_and_shows_modify_column() -> None:
    fidelity = {
        "metadata": {"profile": "xlsx", "run_date": "2026-01-01T00:00:00Z"},
        "libraries": {
            "wolfxl": {"capabilities": ["read", "write"]},
            "openpyxl": {"capabilities": ["read", "write"]},
            "pyumya": {"capabilities": ["read", "write"]},
        },
        "results": [
            {
                "feature": "cell_values",
                "library": "wolfxl",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
            {
                "feature": "cell_values",
                "library": "openpyxl",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
            {
                "feature": "cell_values",
                "library": "pyumya",
                "scores": {"read": 3, "write": 3},
                "test_cases": {},
            },
        ],
    }

    lines = _build_dashboard(fidelity, perf=None)
    doc = "\n".join(lines)

    assert "| Library | Caps | Modify | Green Features | Pass Rate | Best For |" in doc
    assert "| wolfxl | R+W | Patch |" in doc
    assert "| openpyxl | R+W | Rewrite |" in doc
    assert "pyumya" not in doc
