from __future__ import annotations

from pathlib import Path

from excelbench.results.scatter import render_scatter_features, render_scatter_tiers


def test_render_scatter_tiers_creates_png_and_svg(tmp_path: Path) -> None:
    fidelity = Path("results/xlsx/results.json")
    perf = Path("results/perf/results.json")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json to exist"
    assert perf.exists(), "expected repo fixture results/perf/results.json to exist"

    paths = render_scatter_tiers(fidelity, perf, tmp_path)

    png = tmp_path / "scatter_tiers.png"
    svg = tmp_path / "scatter_tiers.svg"
    assert png in paths
    assert svg in paths
    assert png.exists() and png.stat().st_size > 0
    assert svg.exists() and svg.stat().st_size > 0
    assert "<svg" in svg.read_text()


def test_render_scatter_features_creates_png_and_svg(tmp_path: Path) -> None:
    fidelity = Path("results/xlsx/results.json")
    perf = Path("results/perf/results.json")
    assert fidelity.exists(), "expected repo fixture results/xlsx/results.json to exist"
    assert perf.exists(), "expected repo fixture results/perf/results.json to exist"

    paths = render_scatter_features(fidelity, perf, tmp_path)

    png = tmp_path / "scatter_features.png"
    svg = tmp_path / "scatter_features.svg"
    assert png in paths
    assert svg in paths
    assert png.exists() and png.stat().st_size > 0
    assert svg.exists() and svg.stat().st_size > 0
    assert "<svg" in svg.read_text()

