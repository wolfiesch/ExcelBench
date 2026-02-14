from __future__ import annotations

from pathlib import Path

from excelbench.results.heatmap import render_heatmap


def test_render_heatmap_creates_png_and_svg(tmp_path: Path) -> None:
    results_json = Path("results/xlsx/results.json")
    assert results_json.exists(), "expected repo fixture results/xlsx/results.json to exist"

    paths = render_heatmap(results_json, tmp_path)

    png = tmp_path / "heatmap.png"
    svg = tmp_path / "heatmap.svg"

    assert png in paths
    assert svg in paths
    assert png.exists() and png.stat().st_size > 0
    assert svg.exists() and svg.stat().st_size > 0
    assert "<svg" in svg.read_text()


def test_render_heatmap_empty_matrix_returns_empty(tmp_path: Path) -> None:
    empty = tmp_path / "empty.json"
    empty.write_text('{"metadata": {}, "libraries": {}, "results": []}')

    paths = render_heatmap(empty, tmp_path)
    assert paths == []

