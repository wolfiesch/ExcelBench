import json

from excelbench.generator.generate import load_manifest
from excelbench.generator.generate_xls import generate_xls


def test_generate_xls_creates_expected_manifest_and_files(tmp_path):
    output_dir = tmp_path / "excel_xls"
    manifest = generate_xls(output_dir)

    assert manifest.file_format == "xls"
    assert len(manifest.files) == 4
    assert (output_dir / "manifest.json").exists()

    parsed = load_manifest(output_dir / "manifest.json")
    assert parsed.file_format == "xls"
    assert {f.feature for f in parsed.files} == {
        "cell_values",
        "alignment",
        "dimensions",
        "multiple_sheets",
    }
    assert all(f.file_format == "xls" for f in parsed.files)
    assert all((output_dir / f.path).exists() for f in parsed.files)


def test_generate_xls_cell_values_uses_literal_error_tokens(tmp_path):
    output_dir = tmp_path / "excel_xls"
    generate_xls(output_dir, features=["cell_values"])

    with open(output_dir / "manifest.json") as f:
        manifest_data = json.load(f)

    file_data = manifest_data["files"][0]
    cases = {tc["id"]: tc["expected"] for tc in file_data["test_cases"]}
    assert cases["error_div0"] == {"type": "error", "value": "#DIV/0!"}
    assert cases["error_na"] == {"type": "error", "value": "#N/A"}
    assert cases["error_value"] == {"type": "error", "value": "#VALUE!"}
