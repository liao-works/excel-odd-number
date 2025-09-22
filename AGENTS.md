# Repository Guidelines

## Project Structure & Module Organization
`main.py` boots the ttkbootstrap UI, wires logging, and delegates feature work into `src`. Business logic lives in `src/core` (`dpd/`, `ups/`, and shared processors), UI widgets in `src/ui`, and file helpers in `src/utils`. Global flags and paths belong in `config.py`, while distributable templates should stay under `templates/` and any icons or static files under `assets/`. Keep documentation updates in `docs/`, and reserve top-level `build*.py` scripts for packaging tweaks only.

## Build, Test, and Development Commands
Create or activate a virtual environment, then install runtime deps with `pip install -r requirements.txt`. Use `python main.py` for local runs; it writes runtime diagnostics to `app.log`. Package a Windows executable through `python build.py` (full environment checks) or `python build_simple.py` for a quicker single-file build. `python build_en.py` and `python build_github.py` are locale- or CI-focused variations—adjust them instead of duplicating new scripts.

## Coding Style & Naming Conventions
Follow PEP 8 with four-space indentation and `snake_case` for modules, functions, and filenames (`main_window.py`, `file_handler.py`). Classes stay in `PascalCase`, constants in `UPPER_CASE`. Prefer explicit imports from package modules over relative wildcards, and keep GUI wiring separate from data transformers (UI -> `core` -> `utils`). Extend `config.py` rather than scattering literals; update docstrings when adding public functions.

## Testing Guidelines
There is no automated suite yet—new work should introduce `pytest` tests under a top-level `tests/` directory mirroring the package layout (e.g., `tests/core/test_dpd_processor.py`). Seed fixtures with small `.xlsx` samples in `tests/fixtures/` to avoid bloating the repo. Run `pytest -q` locally and document expected coverage in your PR description until formal thresholds exist.

## Commit & Pull Request Guidelines
Commits follow Conventional Commits (`feat:`, `fix:`, `refactor:`) with short, present-tense summaries; include a scope (`dpd`/`ui`) when it clarifies the surface area. For PRs, provide a concise changelog, link relevant GitHub issues, and attach screenshots or sample exports when UI or template behavior changes. Note any build or packaging impacts and flag required follow-up tasks so release scripts can be updated promptly.

## Template & Asset Management
Treat `templates/` as shipping artifacts: version Excel files alongside code changes and document schema expectations in commit notes. When altering assets, check that packaging scripts copy them into `dist/` and update `config.py` paths if directory names change. Clean temporary workbooks before committing to keep the repository lightweight.
