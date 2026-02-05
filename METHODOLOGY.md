# ExcelBench Methodology

## Purpose
ExcelBench measures **feature fidelity** for Excel libraries—how accurately they can read and write Excel features compared to native Excel. It is not a performance benchmark.

## Fidelity vs Performance
- **Fidelity** evaluates correctness and completeness of features (formats, formulas, borders, etc.).
- **Performance** (speed, memory) is out of scope for now and may be added later.

## Oracle Strategy
ExcelBench uses a **hybrid write verification** model:
- **Primary oracle:** Excel itself via `xlwings` (highest fidelity).
- **Fallback oracle:** `openpyxl` when Excel is unavailable (CI / headless).

This allows reliable local verification while keeping CI runnable.

## Fixtures Policy
Two fixture paths exist:
- `fixtures/excel/` (tracked): **canonical Excel-generated files** used in CI.
- `test_files/` (gitignored): local scratch output for development.

Generate canonical fixtures with:
```bash
uv run excelbench generate --output fixtures/excel
```

## Scoring
Current scoring (Phase 1/Tier 1) uses a simplified pass-rate model:
- 3: all tests pass
- 2: ≥80% pass
- 1: ≥50% pass
- 0: <50% pass

Rubric-aligned scoring (see `rubrics/fidelity-rubrics.md`) will replace this once test case importance tiers are wired in.

## Reproducibility
All results are written to `results/results.json` as the source of truth, then rendered into `results/README.md` and `results/matrix.csv` for human and CSV consumption.
