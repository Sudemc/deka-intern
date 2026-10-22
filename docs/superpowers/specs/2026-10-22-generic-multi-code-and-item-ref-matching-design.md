# Generic Multi-Code and Item/Reference Disambiguation Design

**Date:** 2026-10-22  
**Status:** Approved  
**Author:** Pair Programming Session with DEKA

---

## 1. Problem Statement & Motivation

In PCB manufacturing and BOM comparison workflows, customer BOMs vary widely across different electronics companies. Two critical issues have emerged:

1. **Ambiguous Column Name Collision (`Parça` vs `REFERANS`):**  
   In customer BOMs, a column titled `Parça` frequently contains row/item sequence numbers (`37, 44, 127...`), while a separate column titled `REFERANS` contains actual schematic circuit designators (`D13, D15`, `C1, C2`, `R1, R2`).  
   Because the system's keyword dictionary mapped both `parça` and `referans` to the internal field `part` (Reference Designator), `_auto_layout` detected duplicate columns for `part` and raised a fatal `ValueError: ... has ambiguous column(s): Parça`.

2. **Turkish Unicode Casefolding Discrepancy:**  
   In Python standard library, `"ÜRETİCİ KODU".casefold()` produces `'üretici\u0307 kodu'` (with a combining dot `\u0307` for `İ`), causing equality checks against `'üretici kodu'` or `'uretici kodu'` to fail silently.

3. **Customer Internal Part Number vs. Global Manufacturer Part Number (MPN):**  
   Customer BOMs often contain **multiple code columns** representing the same component:
   - A customer-specific internal stock code (e.g., `DTSIS KODU: H010101005006`, or in other companies `FIRMA KODU`, `STOK KODU`, `MUSTERI KODU`).
   - A universal manufacturer part number (e.g., `ÜRETİCİ KODU: MBR0540T1G`).
   
   In DEKA's ERP system, the component may be registered either under the customer's internal code (`H010101005006`) or under the manufacturer code (`MBR0540T1G`).  
   Currently, the system either chooses only one column as `part_number` or crashes with `"birden fazla ürün kodu adayı bulundu"`. As a result, valid components fail to match directly against ERP.

4. **Strict Constraint:**  
   **No company-specific names (such as `DTSIS`) must ever be hardcoded.** The solution must be completely generic, detecting code columns by pattern (e.g., ending with `kodu`, `kod`, `code`, `number`, `no`, `pn`) and content analysis.

---

## 2. Architecture & Design Decisions

### 2.1 Robust Turkish Header Normalization
Replace ad-hoc lowercasing with deterministic diacritic normalization:
```python
def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("İ", "i").replace("I", "ı").lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.split())
```
This guarantees `'ÜRETİCİ KODU'` normalizes to `'uretici kodu'`, `'DEĞER'` to `'deger'`, `'REFERANS'` to `'referans'`.

### 2.2 Disambiguating Item / Line Numbers vs. Circuit Reference Designators
1. In `_HEADER_ALIASES`:
   - `part` represents Circuit Reference Designators (`referans`, `ref`, `refdes`, `designator`, `component reference`, `reference designator`).
   - Create an explicit `item` alias category (`item`, `kalem no`, `sira no`, `line no`, `parca no`, `parca`).
2. If a sheet contains both a candidate matching `item`/`parça` and another candidate matching `referans`/`ref`, `referans` is prioritized as `part` (Reference Designator).
3. If inspection of sample rows shows integer row numbers (e.g., `37, 44, 127`), the column is confirmed as `item` number, preventing it from colliding with `part`.

### 2.3 Generic Multi-Code Column Detection
1. Generic pattern matching for headers:
   A column is recognized as a product code column if its normalized header:
   - Ends with or contains `kodu`, `kod`, `code`, `number`, `no`, `pn`, `p n`, `mpn`, `part_number`
   - AND is not a row index/item number (e.g. not `kalem no`, `sira no`, `line no`).
2. Both internal company codes (e.g., `... KODU`, `... NO`) and manufacturer codes (`ÜRETİCİ KODU`, `MPN`) are detected as code columns.

### 2.4 Multi-Code Data Modeling in `PartRecord`
Enhance `PartRecord` in `comparison.py`:
```python
@dataclass(frozen=True)
class PartRecord:
    part_number: str
    description: str
    mounting: str = ""
    reference: str = ""
    material_type: str = ""
    qty: str = ""
    manufacturer: str = ""
    package: str = ""
    duplicate: bool = False
    source_file: str = ""
    source_sheet: str = ""
    source_breakdown: str = ""
    unmapped_prefix: bool = False
    all_codes: tuple[str, ...] = ()
    code_labels: dict[str, str] = field(default_factory=dict)
```
- `part_number`: Primary code (e.g. MPN if available, or first non-empty code).
- `all_codes`: Tuple of all distinct non-empty candidate codes for that row (e.g., `("H010101005006", "MBR0540T1G")`).
- `code_labels`: Mapping of code to source column name (e.g. `{"H010101005006": "DTSIS KODU", "MBR0540T1G": "ÜRETİCİ KODU"}`).

### 2.5 Cross-Matching in `compare_workbooks` (Phase 1)
When comparing a customer record `c_rec` with an ERP record `e_rec`:
1. Build code lookup dictionaries for customer records:
   - Map **each code in `c_rec.all_codes`** to `c_rec`.
   - Also allow normalized versions of all codes.
2. In Phase 1 Part Number Matching:
   - If `e_rec.part_number` (or any of its normalized variants) matches **ANY code** in `c_rec.all_codes`:
     $\rightarrow$ Direct Phase 1 Match!
3. Multiple customer rows mapping to the same code are handled cleanly without collision.

### 2.6 Report Output Formatting
In the output report:
- `ERP Part Number`: The exact code used in ERP (e.g., `H010101005006`).
- `Müşteri Part Number`:
  - If the customer row has multiple codes and ERP matched one of them:
    Show the matched code alongside the alternate code for clarity:
    e.g. `H010101005006 [Üretici: MBR0540T1G]` or `MBR0540T1G [Müşteri: H010101005006]`.
  - If single code, show that code directly.
- Status is classified as `UYGUN` (or normal match) without false "Part Number Uyuşmazlığı" warnings.

---

## 3. Verification & Testing Strategy
1. **Unit Test for Turkish Header Normalization:**
   Verify `ÜRETİCİ KODU`, `DEĞER`, `REFERANS`, `MontajTipi` normalize without combining diacritics.
2. **Unit Test for Parça vs. REFERANS Disambiguation:**
   Feed a DataFrame / workbook with `Parça` (integers 37, 44) and `REFERANS` (D13, D15), ensuring no `ValueError: ambiguous column` is raised.
3. **Unit Test for Generic Multi-Code Matching:**
   Feed customer BOM with `DTSIS KODU: H010101005006` and `ÜRETİCİ KODU: MBR0540T1G`.
   - Case A: ERP has `H010101005006` $\rightarrow$ must match with status UYGUN.
   - Case B: ERP has `MBR0540T1G` $\rightarrow$ must match with status UYGUN.
   - Case C: Random company name (e.g., `ACME KODU: ACME-999` and `ÜRETİCİ KODU: BAT54CTA`) $\rightarrow$ must match generically without ACME hardcoded.
