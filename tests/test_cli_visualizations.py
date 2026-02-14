from __future__ import annotations

from pathlib import Path

import pytest
import typer

from excelbench import cli


def test_cli_heatmap_writes_outputs(tmp_path: Path) -> None:
    results_json = Path("results/xlsx/results.json")
    assert results_json.exists(), "expected repo fixture results/xlsx/results.json to exist"

    cli.heatmap(results_path=results_json, output_dir=tmp_path)

    assert (tmp_path / "heatmap.png").exists()
    assert (tmp_path / "heatmap.svg").exists()


def test_cli_scatter_writes_outputs(tmp_path: Path) -> None:
    fidelity = Path("results/xlsx/results.json")
    perf = Path("results/perf/results.json")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json to exist"
    assert perf.exists(), "expected repo fixture results/perf/results.json to exist"

    cli.scatter(fidelity_path=fidelity, perf_path=perf, output_dir=tmp_path)

    assert (tmp_path / "scatter_tiers.png").exists()
    assert (tmp_path / "scatter_tiers.svg").exists()
    assert (tmp_path / "scatter_features.png").exists()
    assert (tmp_path / "scatter_features.svg").exists()


def test_cli_html_dashboard_writes_output(tmp_path: Path) -> None:
    fidelity = Path("results/xlsx/results.json")
    perf = Path("results/perf/results.json")
    scatter_dir = Path("results/xlsx")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json to exist"
    assert perf.exists(), "expected repo fixture results/perf/results.json to exist"
    assert scatter_dir.exists(), "expected repo fixture results/xlsx to exist"

    out = tmp_path / "dashboard.html"
    cli.html_dashboard(
        fidelity_path=fidelity,
        perf_path=perf,
        output_path=out,
        scatter_dir=scatter_dir,
    )

    assert out.exists() and out.stat().st_size > 0
    assert "<title>ExcelBench Dashboard</title>" in out.read_text()


def test_cli_scatter_errors_when_perf_missing(tmp_path: Path) -> None:
    fidelity = Path("results/xlsx/results.json")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json to exist"

    with pytest.raises(typer.Exit):
        cli.scatter(
            fidelity_path=fidelity,
            perf_path=tmp_path / "missing.json",
            output_dir=tmp_path,
        )

