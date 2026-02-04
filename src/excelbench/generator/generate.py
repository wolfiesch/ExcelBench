"""Main entry point for test file generation."""

import json
from datetime import datetime, timezone
from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.generator.features import (
    CellValuesGenerator,
    TextFormattingGenerator,
    BordersGenerator,
)
from excelbench.models import Manifest, TestFile

# Version of the generator
GENERATOR_VERSION = "0.1.0"


def get_all_generators() -> list[FeatureGenerator]:
    """Get all available feature generators."""
    return [
        CellValuesGenerator(),
        TextFormattingGenerator(),
        BordersGenerator(),
    ]


def get_excel_version() -> str:
    """Get the version of Excel being used."""
    try:
        app = xw.apps.active
        if app is None:
            # Start Excel temporarily to get version
            app = xw.App(visible=False)
            version = app.version
            app.quit()
            return version
        return app.version
    except Exception:
        return "unknown"


def generate_test_file(
    generator: FeatureGenerator,
    output_dir: Path,
) -> TestFile:
    """Generate a single test file using the given generator.

    Args:
        generator: The feature generator to use.
        output_dir: Directory to save test files.

    Returns:
        TestFile metadata describing what was generated.
    """
    print(f"Generating {generator.feature_name}...")

    # Create workbook
    wb, output_path = generator.create_workbook(output_dir)

    try:
        # Get the sheet and generate test cases
        sheet = wb.sheets[0]
        test_cases = generator.generate(sheet)

        # Save and close
        generator.save_and_close(wb, output_path)

        print(f"  Created {output_path} with {len(test_cases)} test cases")

        return TestFile(
            path=str(output_path.relative_to(output_dir)),
            feature=generator.feature_name,
            tier=generator.tier,
            test_cases=test_cases,
        )

    except Exception as e:
        # Make sure to close workbook on error
        try:
            wb.close()
        except Exception:
            pass
        raise RuntimeError(f"Failed to generate {generator.feature_name}: {e}") from e


def generate_all(
    output_dir: Path,
    generators: list[FeatureGenerator] | None = None,
) -> Manifest:
    """Generate all test files.

    Args:
        output_dir: Directory to save test files.
        generators: Optional list of generators. If None, uses all available.

    Returns:
        Manifest describing all generated files.
    """
    if generators is None:
        generators = get_all_generators()

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Get Excel version before generating
    excel_version = get_excel_version()
    print(f"Using Excel version: {excel_version}")

    # Generate each test file
    test_files: list[TestFile] = []
    for generator in generators:
        test_file = generate_test_file(generator, output_dir)
        test_files.append(test_file)

    # Create manifest
    manifest = Manifest(
        generated_at=datetime.now(timezone.utc),
        excel_version=excel_version,
        generator_version=GENERATOR_VERSION,
        files=test_files,
    )

    # Write manifest to JSON
    manifest_path = output_dir / "manifest.json"
    write_manifest(manifest, manifest_path)
    print(f"Wrote manifest to {manifest_path}")

    return manifest


def write_manifest(manifest: Manifest, path: Path) -> None:
    """Write manifest to JSON file."""
    data = {
        "generated_at": manifest.generated_at.isoformat(),
        "excel_version": manifest.excel_version,
        "generator_version": manifest.generator_version,
        "files": [
            {
                "path": f.path,
                "feature": f.feature,
                "tier": f.tier,
                "test_cases": [
                    {
                        "id": tc.id,
                        "label": tc.label,
                        "row": tc.row,
                        "expected": tc.expected,
                    }
                    for tc in f.test_cases
                ],
            }
            for f in manifest.files
        ],
    }

    with open(path, "w") as f:
        json.dump(data, f, indent=2)


def load_manifest(path: Path) -> Manifest:
    """Load manifest from JSON file."""
    from excelbench.models import TestCase

    with open(path) as f:
        data = json.load(f)

    return Manifest(
        generated_at=datetime.fromisoformat(data["generated_at"]),
        excel_version=data["excel_version"],
        generator_version=data["generator_version"],
        files=[
            TestFile(
                path=f["path"],
                feature=f["feature"],
                tier=f["tier"],
                test_cases=[
                    TestCase(
                        id=tc["id"],
                        label=tc["label"],
                        row=tc["row"],
                        expected=tc["expected"],
                    )
                    for tc in f["test_cases"]
                ],
            )
            for f in data["files"]
        ],
    )


if __name__ == "__main__":
    # Simple CLI for testing
    import sys

    output = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("test_files")
    generate_all(output)
