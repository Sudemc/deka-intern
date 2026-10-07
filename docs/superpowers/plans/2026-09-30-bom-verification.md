# BOM Verification and ERP Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Python workflow that validates customer BOM fields using synthetic samples and optional Nexar/Octopart data, then writes a reviewable ERP-shaped workbook.

**Architecture:** Use pandas for workbook input, deterministic validation, a batched component-provider interface with an offline fictional catalog and a Nexar GraphQL adapter, and openpyxl for a separate ERP workbook with changed cells highlighted and review details on a second sheet. Keep credentials in environment variables and keep sample mode usable offline.

**Tech Stack:** Python, pandas, openpyxl, requests, pytest.

**Spec:** `docs/superpowers/specs/2026-09-30-bom-verification-design.md`

## Global Constraints

- Preserve the customer source workbook and write output to a separate file.
- Clearly label interim workbooks and catalog data as synthetic and unverified.
- Never invent a replacement description, quantity, or mounting classification without evidence.
- Distinguish board references in `Part` from manufacturer's code in `Part Number`.
- Mark both added and corrected descriptions; report unresolved checks for manual review.
- Use current Nexar/Octopart GraphQL API for online lookups and batch exact manufacturer-part-number matching.
- Read `NEXAR_CLIENT_ID` and `NEXAR_CLIENT_SECRET` from the environment; never log credentials or tokens.
- Handle API errors, rate limits, missing data, partial responses, and ambiguous matches as unverified results; expose unresolved code candidates separately.
- Use configurable workbook paths and column mappings so real files can replace samples later.

## Review Focus

- Ambiguous duplicate non-code fields must fail with workbook/sheet context; missing fields remain blank for review, while code-like candidate columns are checked through exact provider results.
- Comma-separated references and `Qty` mismatch must be detected without treating one multi-reference row as duplicate records (`test_quantity_matches_reference_count`, Task 2).
- Unknown or duplicate catalog/API matches must not be reported as verified (`test_unknown_and_ambiguous_part_number`, Task 2; `test_nexar_rejects_non_exact_mpn`, Task 6).
- Missing and incorrect `Value` descriptions must both be visibly marked only when exact part identity and sufficient supported attributes exist (`test_description_add_and_correction_require_match`, Task 2 and `test_changed_cells_are_highlighted`, Task 4).
- Unknown/conflicting mounting clues and API failures must be surfaced without guessing (`test_mounting_conflict_and_unknown`, Task 2; `test_nexar_graphql_and_http_errors_are_unverified`, Task 6).

---

### Task 1: Project scaffold and synthetic workbook examples

**Files:**
- Create: `pyproject.toml`
- Create: `README.md`
- Create: `.gitignore`
- Create: `examples/customer_bom_sample.xlsx`
- Create: `examples/erp_bom_sample.xlsx`
- Create: `examples/component_catalog_sample.json`

**Interfaces:**
- Produces the two workbook layouts named in the spec and a clearly fictional lookup catalog consumed by Task 2.
- Customer workbook columns: `Part`, `Qty`, `Value`, `Part Number`, `Manufacture`, `Satın Alma`, `Genel Stok`.
- ERP workbook columns: `No`, `Part`, `Qty`, `Description`, `TH`, `Part Number`, `Alternative`.

- [ ] Add setuptools build metadata, runtime dependencies (`pandas`, `openpyxl`, `requests`), and the pytest development dependency in `pyproject.toml`; configure pytest to discover `tests` and add `src` to its import path.
- [ ] Create the customer workbook with correct, missing, and intentionally incorrect `Value` examples; include multi-reference rows and one quantity/reference-count mismatch.
- [ ] Create the ERP-shaped workbook with exact target headers and representative `TH`/`SMD` values; label it as a format example, not as a source of product truth.
- [ ] Create a fictional component JSON catalog keyed by exact manufacturer part number; include expected description, mounting, and structured attributes for sample parts, with a synthetic-data disclaimer.
- [ ] Document Windows PowerShell setup, offline sample usage, and that all sample data is fictional; Task 6 documents live Nexar setup.
- [ ] Confirm sample workbooks open and contain the required headers; do not make external API calls.

### Task 2: Domain models, batch provider interface, offline catalog, and validation

**Files:**
- Create: `src/bom_enricher/__init__.py`
- Create: `src/bom_enricher/models.py`
- Create: `src/bom_enricher/component_data.py`
- Create: `src/bom_enricher/validation.py`
- Create: `tests/test_validation.py`

**Interfaces:**
- `CustomerBomRow(part: str, qty: object, value: str, part_number: str, manufacturer: str = "")`.
- `ComponentData(part_number: str, manufacturer: str, short_description: str, attributes: dict[str, str], datasheet_url: str = "")`.
- `ComponentLookupResult(status: str, component: ComponentData | None, message: str = "")`, with status `matched`, `not_found`, `ambiguous`, or `unverified`.
- `CheckIssue(field: str, status: str, original: str, proposed: str, message: str)` where status is `added`, `corrected`, `conflict`, `unverified`, or `invalid`.
- `RowResult(source: CustomerBomRow, description: str, mounting: str, issues: list[CheckIssue])`.
- `ComponentDataProvider.lookup_many(part_numbers: list[str]) -> dict[str, ComponentLookupResult]`.
- `JsonCatalogProvider(catalog_path).lookup_many(part_numbers)`: exact normalized part-number lookup; duplicate normalized keys are a load error.
- `validate_row(row: CustomerBomRow, lookup: ComponentLookupResult) -> RowResult`.

- [ ] Write `test_unknown_and_ambiguous_part_number`, `test_description_add_and_correction_require_match`, `test_quantity_matches_reference_count`, and `test_mounting_conflict_and_unknown` before implementation.
- [ ] Run targeted tests and confirm they fail because the interfaces/behavior do not exist yet.
- [ ] Implement typed domain models and the batch provider interface; make the JSON provider return an explicit status for every requested part number.
- [ ] Parse `Part` as comma-separated, trimmed references; require positive numeric `Qty` and report mismatch when it differs from listed reference count.
- [ ] Preserve a description when supported attributes agree; add or correct it only when exact component identity and sufficient structured evidence support deterministic ERP formatting. Otherwise keep the source text and mark it unverified.
- [ ] Derive `TH`/`SMD` only from a clear supported mounting attribute; report conflicts with explicit `Value` clues and leave missing/unclear cases unresolved.
- [ ] Run `pytest tests/test_validation.py -q` and confirm all validation tests pass.

### Task 3: Customer workbook loading and column validation

**Files:**
- Create: `src/bom_enricher/excel_io.py`
- Create: `tests/test_excel_io.py`

**Interfaces:**
- Explicit `ColumnMapping` contains source names for `part`, `qty`, `value`, `part_number`, and optional `manufacturer`; default reading recognizes common Turkish/English header variants and candidate product-code columns.
- `read_customer_bom(path: Path, sheet_name: str | int = 0, mapping: ColumnMapping | None = None) -> list[CustomerBomRow]`.
- Inspect raw rows before pandas parsing to locate a likely header row. Ambiguous duplicate semantic columns other than product-code columns raise with workbook/sheet context. Missing fields remain blank so remaining BOM rows can still be processed.

- [ ] Write `test_read_customer_sample`, `test_missing_and_duplicate_headers`, `test_custom_column_mapping`, and tests for title rows, header synonyms, and unlabeled product-code columns.
- [ ] Run targeted tests and confirm missing behavior fails first.
- [ ] Implement input reading with pandas `read_excel`, first sheet and observed header aliases by default; infer unlabeled product-code values only as candidates.
- [ ] Normalize blank cells to empty strings while retaining quantity values; preserve row order.
- [ ] Run `pytest tests/test_excel_io.py -q` and confirm all loader tests pass.

### Task 4: ERP workbook output and review markings

**Files:**
- Create: `src/bom_enricher/output.py`
- Create: `tests/test_output.py`

**Interfaces:**
- `write_erp_workbook(results: list[RowResult], output_path: Path) -> None`.
- `BOM` columns: `No`, `Part`, `Qty`, `Description`, `TH`, `Part Number`, `Alternative`.
- `Review` columns: `No`, `Part`, `Part Number`, `Field`, `Original`, `Proposed`, `Status`, `Reason`.
- `Unresolved Parts` lists unique unverified product-code candidates with source columns and row numbers.

- [ ] Write `test_erp_columns_and_separate_output`, `test_added_and_corrected_cells_are_highlighted`, and `test_review_sheet_records_original_and_proposed_values` before implementation.
- [ ] Run targeted tests and confirm they fail before the writer exists.
- [ ] Use openpyxl to write a separate workbook with ERP columns on `BOM`, field issues on `Review`, and unverified product-code candidates on `Unresolved Parts`; leave the source unchanged.
- [ ] Highlight added/corrected cells and list every issue, including unresolved checks, with original/proposed values and reason on `Review`.
- [ ] Run `pytest tests/test_output.py -q` and confirm all writer tests pass.

### Task 5: CLI, offline end-to-end run, and user documentation

**Files:**
- Create: `src/bom_enricher/main.py`
- Create: `tests/test_cli.py`
- Modify: `README.md`

**Interfaces:**
- Command: `python -m bom_enricher.main --input PATH --provider sample|nexar --output PATH [--catalog PATH] [--sheet SHEET]`.
- CLI loads rows, batches unique manufacturer part numbers through the chosen provider, validates rows, writes a separate ERP workbook, and prints counts for unchanged, added, corrected, invalid, and unresolved checks.
- `sample` is the offline default and uses the synthetic catalog; no credentials/network required.

- [ ] Write `test_cli_generates_erp_workbook_from_samples` and assert ERP columns, review entries, unresolved candidates, and unchanged source.
- [ ] Run targeted test to confirm it fails before the CLI exists.
- [ ] Implement argument parsing and orchestration without embedding provider, validation, or workbook logic in the CLI module.
- [ ] Document PowerShell editable install (`python -m pip install -e ".[dev]"`), offline command, and limitations of synthetic catalog validation.
- [ ] Run `pytest -q` and the documented offline CLI command; inspect output workbook and confirm source sample is unchanged.

### Task 6: Nexar OAuth and Octopart GraphQL provider

**Files:**
- Modify: `src/bom_enricher/component_data.py`
- Create: `tests/test_nexar_provider.py`
- Modify: `README.md`

**Interfaces:**
- `NexarOctopartProvider(client_id: str | None = None, client_secret: str | None = None, session: requests.Session | None = None)`; absent explicit values, read `NEXAR_CLIENT_ID` and `NEXAR_CLIENT_SECRET` from the environment.
- Implements `lookup_many(part_numbers: list[str]) -> dict[str, ComponentLookupResult]` using Nexar GraphQL `supMultiMatch`.
- Obtains OAuth 2.0 client-credentials token from `https://identity.nexar.com/connect/token` with `scope=supply.domain`; sends GraphQL to `https://api.nexar.com/graphql`.

- [ ] Write mocked HTTP tests `test_nexar_auth_and_batch_query`, `test_nexar_rejects_non_exact_mpn`, `test_nexar_graphql_and_http_errors_are_unverified`, `test_nexar_timeout_is_unverified`, and `test_missing_credentials_fails_clearly`.
- [ ] Run targeted tests and confirm they fail before provider implementation.
- [ ] Implement token acquisition and in-memory token reuse until expiry; never log tokens or secrets.
- [ ] Implement one `supMultiMatch` batch query requesting MPN, manufacturer, short description, technical specs, and datasheet link using field names verified against current Nexar GraphQL schema.
- [ ] Accept only exact MPN matches; when the customer manufacturer is supplied, reject/flag manufacturer disagreement. Convert missing, ambiguous, GraphQL, HTTP, timeout, and rate-limit results to `unverified` rather than corrections.
- [ ] Parse structured specifications into provider-neutral attributes without fabricating absent values. Do not automatically download datasheet PDFs in this phase.
- [ ] Document account/application setup, environment-variable names, plan/usage caveat, and Nexar CLI command; never request the user to paste secrets into chat or commit them.
- [ ] Run `pytest tests/test_nexar_provider.py -q` with mocked responses; no live API call unless credentials and access have been provisioned.

## Deferred until real workbook and account access are available

- Confirm actual worksheet names, merged cells, formulas, exact header variants, formatting, and ERP import constraints.
- Verify Nexar access plan, usage limits, and expected BOM volume before live or production use.
- Run a live API smoke check and refine actual attribute mapping using returned specifications.
- Verify factual quantity against an approved reference BOM; initial checks remain structural.
- Finalize ERP-specific description vocabulary, ordering, abbreviations, and length limits.
