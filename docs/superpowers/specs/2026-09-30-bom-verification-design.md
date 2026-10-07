# BOM Verification and ERP Export Design

**Date:** 2026-09-30  
**Status:** Approved design, updated to use Nexar/Octopart

## Purpose

Build a local Python tool that reads a customer BOM workbook, checks component descriptions and selected BOM fields against component data, and writes a separate workbook shaped for the Deka ERP workflow. The tool must make corrections and unresolved issues easy to review.

## Observed workbook structures

The supplied images show two report formats:

- Customer BOM: `Part`, `Qty`, `Value`, `Part Number`, `Manufacture`, `Satın Alma`, and `Genel Stok`.
- Deka ERP BOM: `No`, `Part`, `Qty`, `Description`, `TH`, `Part Number`, and `Alternative`.

`Part` identifies board references (for example `C1` or `R2`), while `Part Number` identifies the manufacturer's component. The implementation must keep these meanings separate. These names and layouts are based on images; actual workbook files are still needed to confirm sheet names, exact headers, merged cells, formulas, and data types.

## Proposed workflow

1. Inspect the selected workbook sheet for a likely header row and map common Turkish/English header variants. When a product-code heading is missing or unfamiliar, identify code-like column candidates and use exact provider matches to resolve a unique candidate. Preserve unresolved candidates for manual review rather than guessing.
2. Preserve the input file and create a separate output workbook using the ERP-oriented column layout.
3. Treat customer `Value` as the initial description candidate. Batch-look up manufacturer part numbers through Nexar/Octopart and check the returned exact MPN, manufacturer, and available technical specifications.
4. Keep descriptions that agree with available authoritative attributes. Fill missing descriptions or correct descriptions only when the matched part and required attributes provide enough evidence for deterministic formatting. Mark both added and corrected descriptions in the output.
5. Validate `Qty` for presence, numeric format, positivity, and internal consistency (for example, compare the quantity with comma-separated board references). Without an approved reference BOM, do not claim to verify the real-world quantity.
6. Validate `TH/SMD` against explicit customer description clues and reliable component/package information. Mark conflicts and uncertain classifications for review.
7. Preserve original values for fields that are changed or flagged so the reviewer can compare before and after. Record a clear per-row/per-field result such as unchanged, added, corrected, conflict, invalid, or unverified. Final placement will use ERP columns on the `BOM` sheet and review metadata on a separate `Review` sheet.
8. Save a new ERP-oriented Excel file; never overwrite the customer source. Make added/corrected cells visibly distinguishable, list field issues on `Review`, and list unique unverified product-code candidates plus their source locations on `Unresolved Parts`.

## Component data provider: Nexar/Octopart

Use the current Octopart supply-data API through Nexar's GraphQL endpoint (`https://api.nexar.com/graphql`). Submit multiple manufacturer part numbers in a batch with `supMultiMatch`; request only the fields needed for part identity, manufacturer, short description, technical specifications, and datasheet link. Use documented current fields and confirm their exact GraphQL schema before implementing the query.

Treat returned MPNs and manufacturer names as match evidence, not as blindly trusted lookup keys. Compare returned MPNs with the input exactly after conservative whitespace normalization; if results do not uniquely identify the requested manufacturer's part, report the lookup as ambiguous or unverified. Use technical specifications as structured evidence; missing or conflicting attributes must not produce invented values.

Authenticate using Nexar OAuth 2.0 client credentials. Read `NEXAR_CLIENT_ID` and `NEXAR_CLIENT_SECRET` from environment variables; never write credentials into source, example workbooks, logs, or committed configuration. Reuse the access token in memory until it expires. Handle timeouts, HTTP failures, GraphQL errors, and rate limits as lookup failures that leave data unverified. Account access and subscription/usage limits must be confirmed before production-scale use; do not assume free access or a specific quota.

The provider interface must support a synthetic local catalog so sample development and tests work offline. Selecting the Nexar provider without credentials must produce a clear actionable error. Mocked HTTP tests must cover API behavior without requiring a live account. Real API verification can run only after the user provisions credentials and access.

## Description and field normalization

Description comparison and output must use deterministic rules based on the returned structured attributes. Preserve a `Value` that contains all relevant available attributes and has no contradictions. Add or correct a description only when an exact component match and sufficient relevant attributes support the proposed ERP text. If a relevant attribute is absent, conflicting, or not supported by a deterministic rule, retain the source value and mark the field for review. The exact ERP vocabulary, ordering, abbreviations, and length limits must be confirmed against a real workbook before relying on descriptions for import.

Derive `TH` or `SMD` only when the available component/package information clearly supports the mounting type. Do not equate a package shape with a mounting type without a reliable attribute mapping. If source clues disagree with the matched component data, mark a conflict and show the evidence in the review sheet.

Description normalization should initially use deterministic rules. An AI service is not part of the first implementation because it could produce plausible but unsupported technical claims and make validation less repeatable.

## Error and uncertainty handling

- Missing or ambiguous workbook fields must be reported with actionable row/column context. Header synonyms and code-pattern analysis are heuristics; when more than one candidate remains, the row must not be silently assigned to one code.
- API/network errors, rate limits, absent matches, ambiguous matches, GraphQL errors, and missing attributes must not be treated as verified results.
- Do not invent a replacement description, quantity, or mounting classification when evidence is insufficient.
- Preserve the source workbook and save results to a different output path.
- Mark changes and unresolved checks so a reviewer can decide what enters ERP.
- Keep API credentials out of files and logs.

## Sample workbooks before real files arrive

Create two clearly labeled synthetic Excel workbooks based on the supplied report images:

- a customer-input sample with the observed customer columns (`Part`, `Qty`, `Value`, `Part Number`, `Manufacture`, `Satın Alma`, `Genel Stok`);
- an ERP-target sample with the observed ERP columns (`No`, `Part`, `Qty`, `Description`, `TH`, `Part Number`, `Alternative`).

Use representative fictional rows and cases needed to exercise the workflow, including descriptions that are correct, missing, or intentionally incorrect; valid and review-worthy quantity cases; and TH/SMD matches or conflicts. Clearly mark the files and data as synthetic examples, not real customer or verified manufacturer data. Do not transcribe uncertain text from the photographs as authoritative data. Develop and demonstrate the initial workflow against these examples, keeping file paths and column mappings configurable so actual workbooks can replace the examples later.

## Initial project scope

The first implementation should support a local Windows command-line workflow with two modes: an offline sample-catalog mode and an optional Nexar/Octopart mode. It should separate workbook I/O, validation/normalization, and provider responsibilities. It should not include a GUI, automatic ERP integration, credentials in source code, or a mandatory AI dependency.

The actual workbook samples should drive final decisions about column mapping, worksheet selection, formatting preservation, status reporting, and which fields can be updated without disrupting the ERP import process.

## Verification approach

Use the synthetic workbooks and a fictional catalog as end-to-end fixtures to verify:

- header variants, header-row detection, and missing/ambiguous code columns;
- extraction and mapping of `Part`, `Qty`, `Value`, `Part Number`, and ERP output fields;
- missing and corrected descriptions are both marked, and no correction is made when evidence is insufficient;
- quantity format and internal consistency checks;
- `TH/SMD` agreement, conflict, and unknown cases;
- API result exact-MPN/manufacturer matching, missing and ambiguous matches, GraphQL errors, HTTP errors, timeouts, and rate limits using mocked responses;
- credentials are read from the environment and never printed;
- source workbook remains intact and a separate output workbook is written.
- unknown product-code candidates appear on a separate `Unresolved Parts` sheet with source location.

A live API smoke check and compatibility check against real Excel files require the user to provide/provision the relevant access and files.

## Remaining implementation decisions

- Exact input/output workbook layouts and worksheet names after actual Excel files arrive.
- Candidate-code heuristics and header aliases may need adjustment from real customer workbook examples; exact online matching remains the final disambiguation evidence.
- Attribute mapping from Nexar specs to supported ERP description attributes and to mounting type; unsupported or insufficient data stays unverified.
- Whether a trusted reference BOM is available for factual quantity verification.
- Exact ERP description vocabulary, abbreviations, ordering, and length limits.
- Whether the user's Nexar account plan and usage limits fit expected BOM volume.
