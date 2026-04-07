# Project Guidelines

## Code Style
- Keep UI and plotting logic separated: UI code belongs in `ui/main_window.py`; data parsing/plotting logic belongs in `core/plotter.py`.
- Follow existing Python naming style: snake_case for functions/variables and descriptive method names.
- Preserve current user-facing language style (English UI text with Chinese comments/docs where already present).
- Keep error handling user-visible for GUI workflows (use message dialogs, return `None` on recoverable parse/plot failures).

## Architecture
- Entry point: `main.py` creates and starts `DataVisualizerApp`.
- UI layer: `ui/main_window.py` owns tab setup, file dialogs, preview/save actions, and help display.
- Core layer: `core/plotter.py` owns data parsing, calculations, and Matplotlib figure generation for SMPS/PTR/FTIR.
- Preferred boundary: UI gathers inputs and calls core functions; core should not depend on UI widgets.

## Build And Run
- Create and activate a virtual environment, then install deps:
  - `python -m venv .venv`
  - `source .venv/bin/activate` (macOS/Linux)
  - `pip install -r requirements.txt`
- Run the app with `python main.py` from repository root.
- No automated test or lint command is currently defined in this workspace.

## Conventions
- Input files are expected to be `.xlsx`; keep this behavior unless explicitly changing file format support.
- Required columns are strict and domain-specific:
  - SMPS: `Start Time` and `Total Conc.`
  - PTR: `AbsTime` and species columns beginning with `m`
  - FTIR: structured header rows parsed by `parse_ftir_file`
- Time parsing is strict (for example `%H:%M:%S` and `%H:%M` in SMPS paths). Keep validation explicit and fail with clear messages.
- Keep plotting defaults aligned with current scientific usage (density defaults, labels, high-resolution export).

## Documentation
- Link to `README.md` for user workflow, data templates, and troubleshooting instead of duplicating long instructions.
