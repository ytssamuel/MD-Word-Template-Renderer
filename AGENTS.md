# AGENTS.md — SpeedBOT_2 (MD-Word Template Renderer)

## What this is
Python 3.10+ project that parses a structured Markdown format (`編號. 欄位名稱 | 值`, hierarchy via indent) and renders it into `.docx` templates using Jinja2/docxtpl. Ships a CLI, a CustomTkinter GUI, and PyInstaller builds. Source-of-truth docs: `README.md` and `doc/user_guide.md`.

## Layout (non-obvious)
- Real package lives at `src/md_word_renderer/`. **There is no `setup.py` / `pyproject.toml` / editable install** — the package is found by path manipulation, not packaging.
- Root-level entrypoint scripts do `sys.path.insert(0, "src")` themselves, then import `md_word_renderer`:
  - `md2word.py` → CLI (`src/md_word_renderer/cli/main.py`, exposes subcommands: `render`, `batch`, `batch-templates`, `validate`, `info`)
  - `run_gui.py` → GUI (`src/md_word_renderer/gui/`)
- `src/main.py` provides a top-level helper `render_markdown_to_word(...)` for programmatic use.
- Top-level `config.yaml` is an example; the actually-loaded default is `src/md_word_renderer/config/default_config.yaml` (via `md_word_renderer.config.ConfigLoader`).

## Common commands (Windows / PowerShell; venv at `venv\Scripts\python.exe`)
```bash
# Install
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
pip install -r requirements-gui.txt   # only for GUI / build

# Run
python md2word.py --help
python md2word.py render referance/sample_data.md templates/simple_template.docx output/out.docx
python md2word.py batch  referance templates/example_template.docx output/
python md2word.py batch-templates referance/sample_data.md templates/ output/ --prefix pre_ --suffix _v2
python md2word.py validate referance/sample_data.md
python run_gui.py

# Tests (100 collected; 47 Word 既有 + 17 v2.2.0 Excel + 36 v2.2.1 Excel template engine)
python -m pytest test/ -v
python -m pytest test/test_parser.py -v          # single file
python -m pytest test/test_excel_renderer.py -v  # Excel 元件單元測試（含 Phase 1/2/3）
python -m pytest test/test_cli_excel.py -v       # CLI 在 Excel 樣板下的行為
python -m pytest test/ --cov=src/md_word_renderer

# End-to-end smoke (writes to output/). Requires referance/sample_data.md + templates/*.docx + templates/excel/*.xlsx present.
python test_e2e.py       # 7 個測試：5 Word + 2 Excel (v2.2.1)
python test_parse.py       # quick parse smoke against referance/sample_data.md

# Build standalone exe (auto-cleans build/ and dist/, installs PyInstaller if missing)
python build_exe.py --version gui      # → dist/MD-Word-Renderer.exe
python build_exe.py --version cli      # → dist/md2word.exe
python build_exe.py --version both
```

There is no Makefile, no pre-commit, no CI workflow (no `.github/`). Lint/typecheck: only `flake8` is listed in `requirements.txt`; no formatter config checked in, so don't assume Black's defaults are enforced.

## Markdown input format (subtleties)
- Line shape: `N. 欄位名稱 | 值` — the `|` separator is required. Empty value: `欄位名稱 |`.
- Hierarchy is **indentation-based** (default 4 spaces, configurable via `indent_size` in parser and `parsing.indent_size` in config). Mixed Tab/space can be disallowed by `allow_mixed_indent`.
- Escapes inside values: `\|` for literal pipe, `\n` for newline, `\\` for backslash — handled by `parser/escape_handler.py`.
- Images: `![alt](path)` — relative paths are resolved against the source `.md` dir (`parser/markdown_parser.py:_source_dir`).
- Parsed output has **dual indexing**: every field is keyed by `欄位名稱` AND by positional `#N` (e.g. `data["#1"]["key"] == "系統名稱"`). List items get `number`, `key`, `value`, and nested `children`.

## Excel renderer (v2.2.1 — template-driven, Jinja2)

**重大語意翻轉（v2.2.0 → v2.2.1）**：
- v2.2.0：auto-append 純量欄位，{{ }} 不處理
- v2.2.1：樣板為主；只有含 `{{var}}` 的 cell 才會被替換，不自動 append

行為模式：
- cell 含 `{{var}}` → 替換為頂層純量值
- cell 含 `{{data["key with (paren)"]}}` → 特殊字元鍵（透過 `data` 字典）
- cell 含 `{{data["#1"].key}}` → 編號索引
- cell 含 `{% if %}` → 條件渲染（Jinja2 原生）
- 整張 sheet 含 `{% for var in items %}` + `{% endfor %}` → 標記必須在「自己的 row」，body 中間；支援 `loop.index` 等
- 缺變數 → 空字串（silent；可改 `keep` / `error_format`）
- 設定 `excel.template_engine.auto_flatten_lists: true`（預設）→ 對沒有 `{% for %}` 的 list sheet 維持 v2.2.0 自動攤平行為
- 範例：`templates/excel/sample_template.xlsx`（基本替換）、`with_lists_template.xlsx`（for 展開示範）

詳細語意見 `doc/template_syntax_reference.md` 第 10 章。

## Template syntax (Jinja2 + docxtpl)
- Simple vars: `{{系統名稱}}`.
- Chinese-keyed vars with brackets / slashes need bracket access: `{{data["需求依據(INC/PBI)"]}}`, `{{data["通知/公告方式"]}}`.
- Image items in lists have `type == "image"` and render via `{{child.image}}` (handled by `renderer/image_handler.py`).
- Missing-variable formatting is configurable via `rendering.error_format` in config (default `"[ERROR: 變數 '{var}' 不存在]"`).

## Fixtures and what's gitignored
- The only Markdown kept in the repo is `referance/sample_data.md`. Everything else under `referance/` (including `OH_*.md`, `*_real.md`, generated folders like `OH_20251214/`) is gitignored — those are real business data, not fixtures.
- Templates that ship with the repo (in `templates/`): `simple_template.docx`, `example_template.docx`, `full_template.docx`. There is also `scripts/create_templates.py` that regenerates them.
- Generated artifacts gitignored: `output/`, `dist/`, `build/`, `__pycache__/`, `.pytest_cache/`, `venv/`, `*.docx` (with whitelist for `templates/` and `examples/`).
- `templates/` and `examples/` `.docx` files are whitelisted in `.gitignore` to be tracked despite `*.docx`.

## CLI quirks worth knowing
- `render`, `batch`, and `batch-templates` all accept `--format {auto,docx,xlsx}` (default `auto`); both `batch` and `batch-templates` also have `--no-validate`.
- `batch` and `batch-templates` need explicit `--continue-on-error` to keep going past a failure (default halts on first error).
- `batch-templates` has `--prefix` and `--suffix` flags that aren't documented in README's CLI table but are real (see `src/md_word_renderer/cli/main.py:174-182`).
- Default batch glob is `*.md` for inputs; default templates pattern depends on command (`*.docx`/`*.xlsx`).
- `batch-templates` accepts both `.docx` and `.xlsx` templates in the same folder; output extension follows each template's extension.
- Exit codes: `0`=success, `1`=any failure (non-zero on partial batch failure too).

## When you change things
- After touching code under `src/`, run `python -m pytest test/ -v` before pushing. 47 tests should all pass.
- If you change parser grammar or escape rules, also re-run `python test_parse.py` and `python test_e2e.py` — the e2e covers the round-trip into `templates/*.docx`.
- If you touch GUI windows, run `python run_gui.py` interactively; there are no GUI tests.
- Bumping Python parser defaults? `config.yaml` and `src/md_word_renderer/config/default_config.yaml` both document `parsing.indent_size` / `allow_mixed_indent` / `encoding`. Keep them in sync if you change one.

## Build gotchas
- `build_exe.py` runs `shutil.rmtree('build')` and `shutil.rmtree('dist')` at the start — don't keep uncommitted exe artifacts there.
- `md_word_renderer.spec` (GUI) bundles `doc/`, `README.md`, `assets/`, plus the `.ico` icon. `hiddenimports` lists every `md_word_renderer.gui.*` / `.parser.*` / `.renderer.*` / `.validator.*` / `renderer.excel_*` / `renderer.factory` submodule — add new submodules here or PyInstaller will silently strip them. Excel adds `openpyxl*` and `et_xmlfile`.
- `md_word_renderer_cli.spec` is the CLI counterpart (no GUI window, no `customtkinter` / `tkinter`).
- Both specs set `excludes` to drop large unused libs (`matplotlib`, `numpy`, `pandas`, `pytest`, …). Don't add deps without testing that they're not in this exclude list. Note: `pillow` is kept (openpyxl needs it for image embedding).
