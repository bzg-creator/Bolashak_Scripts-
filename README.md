# 🌍 CountryDictionary Normalization Toolkit

A complete data-cleaning toolkit for normalizing country (and language) reference data in a SQL Server database. The project covers the full pipeline: **audit → analysis → preview → SQL update** — with safe transactions, savepoints, and Excel reports at every stage.

---

## 📋 Table of Contents

- [Background](#background)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Quick Start](#quick-start)
- [Scripts](#scripts)
  - [script1\_f.py — Full Dictionary Audit](#script1_fpy--full-dictionary-audit)
  - [script2\_f.py — Normalization Plan](#script2_fpy--normalization-plan)
  - [script3.py — Change Preview v1](#script3py--change-preview-v1)
  - [script4.py — Change Preview v2 (with "Другое" resolver)](#script4py--change-preview-v2-with-другое-resolver)
- [SQL Scripts](#sql-scripts)
  - [update\_uk.sql — United Kingdom merge](#update_uksql--united-kingdom-merge)
  - [update\_countries.sql — Batch update (7 countries)](#update_countriessql--batch-update-7-countries)
  - [update\_usa.sql — USA with exceptions](#update_usasql--usa-with-exceptions)
  - [update\_batch\_new\_countries.sql — Batch update via temp table](#update_batch_new_countriessql--batch-update-via-temp-table)
  - [query\_script\_languages.sql — Language field update](#query_script_languagessql--language-field-update)
- [Database Tables Affected](#database-tables-affected)
- [Country Name Mapping](#country-name-mapping)
- [Excel Report Structure](#excel-report-structure)
- [Safety & Rollback](#safety--rollback)
- [Workflow](#workflow)

---

## Background

Over time, the `CountryDictionary` reference data became inconsistent across the database:

- The same country stored under multiple names: `США`, `Америка`, `US`, `United States`, `Америка Құрама Штаттары`
- The same country referenced by multiple different GUIDs
- Records with `NULL` in `DictionaryId` — only a raw text value, no GUID binding
- Names in Kazakh, Russian, and English mixed together without a canonical form

This project unifies everything under a **single GUID + single canonical name** per country, across all affected tables.

---

## Project Structure

```
.
├── script1_f.py                        # Step 1 — Full audit, exports Excel report
├── script2_f.py                        # Step 2 — Normalization plan (3-level analysis)
├── script3.py                          # Step 3 — Preview changes from mapping file (v1)
├── script4.py                          # Step 3 — Preview changes from mapping file (v2)
├── Country.xlsx                        # Mapping file: OldId/OldValue → NewId/NewValue
├── update_uk.sql                       # Merge Great Britain + Northern Ireland → UK
├── update_countries.sql                # Batch update 7 countries with savepoints
├── update_usa.sql                      # Update USA with exclusion list
├── update_batch_new_countries.sql      # Batch update via #CountryMap temp table
└── query_script_languages.sql          # Update language dictionary for NULL-IIN records
```

---

## Prerequisites

- Python 3.8+
- SQL Server with ODBC Driver 17
- Windows Authentication (Trusted_Connection)
- Python packages:

```bash
pip install pyodbc openpyxl
```

- For `script4.py`, configure via environment variables (optional):

```bash
set DB_SERVER=your_server_name
set DB_DATABASE=your_database_name
```

---

## Quick Start

1. **Audit the current state** — run `script1_f.py` to get an Excel snapshot of all country values across all tables.
2. **Build the normalization plan** — run `script2_f.py` to identify the three categories of problems.
3. **Fill in `Country.xlsx`** — the mapping sheet `Лист1 (2)` with columns: `Num | NewId | NewValue | OldId | OldValue | SourceEntity | Code | IsDeleted | IsCurrent | Cards | Commission | Placement | LangSchool | Dogovors | Movements | Plans | DictCount | Total`.
4. **Preview changes** — run `script3.py` or `script4.py` to see exactly which records will be touched, without modifying the DB.
5. **Apply SQL scripts** — uncomment `BEGIN TRANSACTION` and `COMMIT` in the relevant SQL file, run ANALYSIS block first, then UPDATE block.
6. **Verify** — use the verification queries at the bottom of each SQL script.

---

## Scripts

### `script1_f.py` — Full Dictionary Audit

Connects to SQL Server and scans all tables for every unique `(DictionaryId, Value)` combination. Results are collected in a temp table `#DictVariants` with a clustered index for performance, then exported to a timestamped Excel file.

**What it does:**
- Pulls data from `Dictionaries`, `Cards` (3 country columns), `Placements`, `Dogovors`, `Movements` (via `WITH NOLOCK` — 14M rows), `Plans`
- Counts usages per table per combination
- Matches values against a built-in `COUNTRY_MAP` (~80 aliases in Russian, Kazakh, English)
- Produces a color-coded Excel report

**Run:**
```bash
python script1_f.py
```

**Output:** `country_analysis_YYYYMMDD_HHMMSS.xlsx`

**Configuration** (edit in script):
```python
DB_SERVER   = "your_server"
DB_DATABASE = "your_database"
```

---

### `script2_f.py` — Normalization Plan

Read-only analysis that classifies all problems into three levels and generates an action plan.

**Three levels:**

| Level | Problem | Fix |
|-------|---------|-----|
| 1 | Wrong name in `Dictionaries.Name` | `UPDATE Dictionaries SET Name = correct` |
| 2 | Same country has multiple GUIDs | Merge all records to one GUID |
| 3 | Records with `NULL` DictionaryId | Find correct GUID and assign it |

**Run:**
```bash
python script2_f.py
```

**Output:** `normalization_plan_YYYYMMDD_HHMMSS.xlsx` — 4 sheets: Level 1, Level 2, Level 3, Summary with step-by-step instructions.

---

### `script3.py` — Change Preview v1

Reads the mapping from `Country.xlsx` and, **without touching the DB**, shows exactly which rows in each table will be updated and what they will change to.

**Features:**
- Loads mapping from Excel (`MAPPING_FILE = "Country.xlsx"`, sheet `Лист1 (2)`)
- Queries each affected table and matches records by OldId
- For `Cards`: shows `CardId` + student full name
- For `Placements`, `Dogovors`, `Plans`: shows record Id
- For `Movements`: shows only COUNT (too large to load row by row)
- Sheet 9: student count per country from `Cards`

**Run:**
```bash
python script3.py
```

**Output:** `normalization_preview_YYYYMMDD_HHMMSS.xlsx`

---

### `script4.py` — Change Preview v2 (with "Другое" resolver)

Extended version of `script3.py` with an additional resolver for records where `NewValue = "Другое"` (literally "Other" — a placeholder in the source data).

**Additional logic:**
- If a `"Другое"` row has an `OldId`, the script finds another row with the same `OldId` and a real name, then substitutes that real name automatically
- Resolved `"Другое"` rows: highlighted in **purple**
- Unresolvable `"Другое"` rows (no OldId): highlighted in **orange** with a warning
- DB credentials via environment variables (`DB_SERVER`, `DB_DATABASE`)

**Run:**
```bash
python script4.py
```

**Output:** `normalization_preview_YYYYMMDD_HHMMSS.xlsx`

---

## SQL Scripts

All SQL scripts follow the same pattern:
1. **ANALYSIS block** — read-only `SELECT`, safe to run anytime
2. **APPLICATION block** — commented-out `UPDATE` statements, uncomment `BEGIN TRANSACTION` / `COMMIT` when ready
3. **VERIFICATION block** — post-update `SELECT` to confirm results (in `/* */` comments)

---

### `update_uk.sql` — United Kingdom merge

Merges two separate entries — *Великобритания* and *Северная Ирландия* — into one canonical record.

```
Old IDs:  12537CCF-...  (Великобритания)
          B43869AB-...  (Северная Ирландия)
New ID:   1021f0c9-4e1c-4f35-9980-def07b2133fe
New Name: Соединенное Королевство Великобритании и Северной Ирландии (Великобритания)
Skip ID:  ECE8B586-...  (different country — excluded)
```

Also handles NULL-Id records matched by old name variants (including Kazakh: `Ұлыбритания және Солтүстік Ирландия Біріккен Патшалығы`).

Tables updated: `Cards` (3 columns), `Placements`, `Dogovors`, `Movements`, `Plans`.

---

### `update_countries.sql` — Batch update (7 countries)

Updates 7 countries in a single script using `@AllOldIds` and `@OldNames` table variables. Each country's old/new Id and name is declared separately, then all UPDATEs run via `JOIN` — one pass per table covers all 7 countries simultaneously.

**Countries:**

| Country | Old Name | New Name |
|---------|----------|----------|
| Абхазия | Абхазия | Республика Абхазия(Абхазия) |
| Австралия | Австралия | Австралийский Союз(Австралия) |
| Австрия | Австрия | Австрийская Республика(Австрия) |
| Аджария | Аджария | Автономная Республика Аджария(Аджария) |
| Азербайджан | Азербайджан | Азербайджанская Республика(Азербайджан) |
| Албания | Албания | Республика Albanия(Албания) |
| Алжир | Алжир | Алжирская Народная Демократическая Республика(Алжир) |

**Savepoints:** `SP_Cards` → `SP_Placements` → `SP_Dogovors` → `SP_Movements` → `SP_Plans`

---

### `update_usa.sql` — USA with exceptions

Updates the United States entry with a hardcoded **exclusion list** — three GUIDs that must never be touched even if they match old name patterns.

```
Old ID:    ECE8B586-FD41-4E10-931E-1C7732BC9E81
New ID:    ca0920bb-e347-4e6c-95db-2a23783fa8b2
Old Name:  Соединенные Штаты Америки
New Name:  Соединенные Штаты Америки (США)

Excluded:  7e49919d-...
           12537ccf-...
           cfea1ddb-...
```

NULL-Id records matched by 8 name variants: `США`, `Америка`, `Америка Құрама Штаттары`, `United States`, `United States of America`, `US`, `USA`, `Соединенные Штаты Америки`.

---

### `update_batch_new_countries.sql` — Batch update via temp table

Uses a `#CountryMap` temporary table as a mapping source instead of inline variables. All countries are loaded in one `INSERT`, then all tables are updated via `JOIN #CountryMap`.

**Countries covered:** Ангола, Андорра, Антигуа, Аргентина, Арктика, Американское Самоа, Антигуа и Барбуда, Афганистан, ОАЭ.

Handles both cases in one structure:
- Records with an `OldId` → matched by GUID
- Records with `NULL` Id → matched by `OldValue` text

---

### `query_script_languages.sql` — Language field update

Targeted update for `UniversityRkLanguageDictionaryId` / `UniversityRkLanguageDictionaryValue` fields in `Cards`, **only for records where `Iin` is NULL or empty** (i.e., incomplete student records).

**Languages updated:**

| Language | Old GUID(s) | New GUID |
|----------|-------------|----------|
| Русский | `516fff8e-...` | `254ab26d-...` |
| Английский | `958590f2-...` | `5cb45a93-...` |
| Казахский | `573f2c5b-...`, `17857e99-...` | `55857469-...` |

Step 1 is a safe read-only `SELECT`. Step 2 is wrapped in a transaction with `ROLLBACK` active by default — change to `COMMIT` only after verifying the preview.

---

## Database Tables Affected

| Table | Columns |
|-------|---------|
| `dbo.Cards` | `CountryDictionaryId`, `CountryDictionaryValue` |
| `dbo.Cards` | `CommissionCountryDictionaryId`, `CommissionCountryDictionaryValue` |
| `dbo.Cards` | `LanguageSchoolCountryDictionaryId`, `LanguageSchoolCountryDictionaryValue` |
| `dbo.Cards` | `UniversityRkLanguageDictionaryId`, `UniversityRkLanguageDictionaryValue` |
| `dbo.Placements` | `CountryDictionaryId`, `CountryDictionaryValue` |
| `dbo.Dogovors` | `CountryDictionaryId`, `CountryDictionaryValue` |
| `dbo.Movements` | `CountryDictionaryId`, `CountryDictionaryValue` |
| `dbo.Plans` | `CountryDictionaryId`, `CountryDictionaryValue` |
| `dbo.Dictionaries` | Referenced as source of truth for `CountryDictionary` type |

> ⚠️ `dbo.Movements` contains ~14 million rows. Updates on this table take significantly longer. Python scripts use `COUNT` only for Movements; SQL scripts update it in a dedicated savepoint block.

---

## Country Name Mapping

The Python scripts contain a `COUNTRY_MAP` / `NAME_MAP` dictionary covering ~80+ aliases across three languages:

```python
"США"                      → "Соединённые Штаты Америки"
"Америка Құрама Штаттары"  → "Соединённые Штаты Америки"
"United States of America" → "Соединённые Штаты Америки"
"Ресей Федерациясы"        → "Российская Федерация"
"Қазақстан"                → "Республика Казахстан"
"Германия Федеративті Республикасы" → "Федеративная Республика Германия"
"Ұлыбритания және Солтүстік Ирландия Біріккен Патшалығы" → "Соединённое Королевство..."
# ... and ~70 more
```

---

## Excel Report Structure

### `script1_f.py` output — `country_analysis_*.xlsx`

| Sheet | Content |
|-------|---------|
| Анализ всех значений | All unique (GUID, Name) pairs with usage counts per table. Color-coded: 🟢 correct, 🟡 needs fix, 🔴 no GUID, grey strikethrough = deleted |
| Требуют исправления | Only the rows that need updating, with current and correct names side by side |
| Легенда | Color legend |

### `script2_f.py` output — `normalization_plan_*.xlsx`

| Sheet | Content |
|-------|---------|
| 1 — Исправить названия | Level 1: wrong names in Dictionaries |
| 2 — Объединить GUID | Level 2: duplicate GUIDs for same country |
| 3 — Записи без Id | Level 3: NULL DictionaryId records |
| Сводка | Summary + step-by-step action plan |

### `script3.py` / `script4.py` output — `normalization_preview_*.xlsx`

| Sheet | Content |
|-------|---------|
| 1. Маппинг замен | Full mapping with change flags (needs Id change? needs Value change?) |
| 2. Cards (Country) | Affected Cards rows with CardId + student name |
| 3. Cards (Commission) | Affected CommissionCountry rows |
| 4. Cards (LangSchool) | Affected LanguageSchool rows |
| 5. Placements | Affected Placements rows |
| 6. Dogovors | Affected Dogovors rows |
| 7. Movements (кол-во) | COUNT only — too large to load |
| 8. Plans | Affected Plans rows |
| 9. Кол-во студентов по странам | Student count per country with % of total |

---

## Safety & Rollback

All SQL update scripts are designed with safety in mind:

**Always run ANALYSIS first:**
```sql
-- Safe to run at any time — read only
PRINT '=== АНАЛИЗ ===';
SELECT ...
```

**Transaction with savepoints:**
```sql
BEGIN TRANSACTION;

    SAVE TRANSACTION SP_Cards;
    -- UPDATE Cards...

    SAVE TRANSACTION SP_Placements;
    -- UPDATE Placements...

    SAVE TRANSACTION SP_Movements;
    -- UPDATE Movements...  ← slowest step

COMMIT TRANSACTION;
```

**Partial rollback (inside open transaction):**
```sql
-- Roll back only Movements and Plans, keep Cards and Placements:
ROLLBACK TRANSACTION SP_Movements;

-- Roll back everything:
ROLLBACK TRANSACTION;
```

**Python scripts are always read-only** — they never execute any `INSERT`, `UPDATE`, or `DELETE`. All writes happen exclusively through the SQL scripts.

---

## Workflow

```
┌─────────────────────────────────────────────────────────────────┐
│  1. script1_f.py  →  country_analysis.xlsx                      │
│     Full audit: what exists, how many usages, what's broken      │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  2. script2_f.py  →  normalization_plan.xlsx                     │
│     Three-level classification: names / GUIDs / NULL records     │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  3. Fill Country.xlsx mapping (OldId → NewId, OldValue → New)   │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  4. script4.py  →  normalization_preview.xlsx                    │
│     See exactly which rows will change — no DB writes            │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  5. Run SQL script — ANALYSIS block only                         │
│     Verify row counts match expectations                         │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  6. Uncomment BEGIN TRANSACTION + COMMIT                         │
│     Run UPDATE blocks — check PRINT row counts at each step      │
└──────────────────────────────┬──────────────────────────────────┘
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│  7. Run VERIFICATION block                                       │
│     New IDs should have rows, old IDs should have 0              │
└─────────────────────────────────────────────────────────────────┘
```
