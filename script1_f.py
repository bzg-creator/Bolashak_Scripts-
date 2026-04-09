
import pyodbc
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime


DB_SERVER   = ""
DB_DATABASE = ""

DICTIONARY_TYPE = "CountryDictionary"

COUNTRY_MAP = {
    "США":                      "Соединённые Штаты Америки",
    "USA":                      "Соединённые Штаты Америки",
    "US":                       "Соединённые Штаты Америки",
    "Америка":                  "Соединённые Штаты Америки",
    "America":                  "Соединённые Штаты Америки",
    "United States":            "Соединённые Штаты Америки",
    "United States of America": "Соединённые Штаты Америки",
    "Америка құрама штаттары":  "Соединённые Штаты Америки",
    "Америка курама штаттары":  "Соединённые Штаты Америки",
    "Россия":                   "Российская Федерация",
    "РФ":                       "Российская Федерация",
    "Russia":                   "Российская Федерация",
    "Ресей":                    "Российская Федерация",
    "Ресей Федерациясы":        "Российская Федерация",
    "Казахстан":                "Республика Казахстан",
    "Казакстан":                "Республика Казахстан",
    "Қазақстан":                "Республика Казахстан",
    "Kazakhstan":               "Республика Казахстан",
    "Германия":                 "Федеративная Республика Германия",
    "ФРГ":                      "Федеративная Республика Германия",
    "Germany":                  "Федеративная Республика Германия",
    "Deutschland":              "Федеративная Республика Германия",
    "Китай":                    "Китайская Народная Республика",
    "КНР":                      "Китайская Народная Республика",
    "China":                    "Китайская Народная Республика",
    "Қытай":                    "Китайская Народная Республика",
    "Великобритания":           "Соединённое Королевство",
    "Англия":                   "Соединённое Королевство",
    "UK":                       "Соединённое Королевство",
    "United Kingdom":           "Соединённое Королевство",
    "Франция":                  "Французская Республика",
    "France":                   "Французская Республика",
    "Турция":                   "Турецкая Республика",
    "Turkey":                   "Турецкая Республика",
    "Türkiye":                  "Турецкая Республика",
    "Узбекистан":               "Республика Узбекистан",
    "Uzbekistan":               "Республика Узбекистан",
    "Кыргызстан":               "Кыргызская Республика",
    "Kyrgyzstan":               "Кыргызская Республика",
    "Таджикистан":              "Республика Таджикистан",
    "Tajikistan":               "Республика Таджикистан",
    "Беларусь":                 "Республика Беларусь",
    "Белоруссия":               "Республика Беларусь",
    "Belarus":                  "Республика Беларусь",
    "Азербайджан":              "Азербайджанская Республика",
    "Azerbaijan":               "Азербайджанская Республика",
    "Армения":                  "Республика Армения",
    "Armenia":                  "Республика Армения",
    "Италия":                   "Итальянская Республика",
    "Italy":                    "Итальянская Республика",
    "Испания":                  "Королевство Испания",
    "Spain":                    "Королевство Испания",
    "Польша":                   "Республика Польша",
    "Poland":                   "Республика Польша",
    "Чехия":                    "Чешская Республика",
    "Czech Republic":           "Чешская Республика",
    "Индия":                    "Республика Индия",
    "India":                    "Республика Индия",
    "Корея":                    "Республика Корея",
    "Южная Корея":              "Республика Корея",
    "South Korea":              "Республика Корея",
    "Иран":                     "Исламская Республика Иран",
    "Iran":                     "Исламская Республика Иран",
}


def get_connection():
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_DATABASE};"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str, timeout=0)  # timeout=0 — без ограничений

STEPS = [

# Шаг 1: создать #DictVariants + сразу индекс
"""
IF OBJECT_ID('tempdb..#DictVariants') IS NOT NULL DROP TABLE #DictVariants;
CREATE TABLE #DictVariants (
    Id           NVARCHAR(36)   NULL,
    [Value]      NVARCHAR(1000) NULL,
    SourceEntity NVARCHAR(400)  NOT NULL DEFAULT(''),
    Code         NVARCHAR(450)  NULL,
    IsDeleted    BIT            NULL,
    IsCurrent    BIT            NULL,
    CardsCount                     INT NOT NULL DEFAULT(0),
    CardCommissionCountryCount     INT NOT NULL DEFAULT(0),
    PlacementCount                 INT NOT NULL DEFAULT(0),
    CardLanguageSchoolCountryCount INT NOT NULL DEFAULT(0),
    DogovorsCountryCount           INT NOT NULL DEFAULT(0),
    MovementsCountryCount          INT NOT NULL DEFAULT(0),
    PlansCountryCount              INT NOT NULL DEFAULT(0),
    DictCount                      INT NOT NULL DEFAULT(0)
)
""",

# Индекс — критично для скорости!
"""
CREATE CLUSTERED INDEX CX_DictVariants
ON #DictVariants (Id, [Value])
""",

# Шаг 2: Dictionaries — источник истины
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, Code, IsDeleted, IsCurrent, DictCount)
SELECT
    CONVERT(NVARCHAR(36), d.Id),
    NULLIF(LTRIM(RTRIM(d.[Name])), N''),
    N'Dictionaries',
    d.Code,
    d.IsDeleted,
    d.IsCurrent,
    COUNT(*) AS DictCount
FROM dbo.Dictionaries d
WHERE d.[Type] = N'CountryDictionary'
  AND (d.Id IS NOT NULL OR NULLIF(LTRIM(RTRIM(d.[Name])), N'') IS NOT NULL)
GROUP BY
    CONVERT(NVARCHAR(36), d.Id),
    NULLIF(LTRIM(RTRIM(d.[Name])), N''),
    d.Code, d.IsDeleted, d.IsCurrent
""",

# Шаг 3: Cards.Country — GROUP BY вместо NOT EXISTS
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, CardsCount)
SELECT Id, [Value], N'Cards.Country', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",

# Обновить CardsCount для уже существующих строк
"""
UPDATE t
SET t.CardsCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Cards.Country%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Cards.Country' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.CardsCount = 0
""",

# Шаг 4: Cards.CommissionCountry
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, CardCommissionCountryCount)
SELECT Id, [Value], N'Cards.CommissionCountry', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",
"""
UPDATE t
SET t.CardCommissionCountryCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Cards.CommissionCountry%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Cards.CommissionCountry' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CommissionCountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.CardCommissionCountryCount = 0
""",

# Шаг 5: Cards.LanguageSchoolCountry
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, CardLanguageSchoolCountryCount)
SELECT Id, [Value], N'Cards.LanguageSchoolCountry', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",
"""
UPDATE t
SET t.CardLanguageSchoolCountryCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Cards.LanguageSchoolCountry%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Cards.LanguageSchoolCountry' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Cards c
    WHERE NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.LanguageSchoolCountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.CardLanguageSchoolCountryCount = 0
""",

# Шаг 6: Placements
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, PlacementCount)
SELECT Id, [Value], N'Placement.Country', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Placements c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",
"""
UPDATE t
SET t.PlacementCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Placement.Country%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Placement.Country' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Placements c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.PlacementCount = 0
""",

# Шаг 7: Dogovors
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, DogovorsCountryCount)
SELECT Id, [Value], N'Dogovors.Country', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Dogovors c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",
"""
UPDATE t
SET t.DogovorsCountryCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Dogovors.Country%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Dogovors.Country' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Dogovors c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.DogovorsCountryCount = 0
""",

# Шаг 8: Movements — 14 млн строк
# Считаем только по уже известным Id из справочника — не сканируем всю таблицу
"""
UPDATE t
SET t.MovementsCountryCount = m.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Movements.Country%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Movements.Country' END
FROM #DictVariants t
JOIN (
    SELECT CountryDictionaryId AS Id, COUNT(*) AS Cnt
    FROM dbo.Movements WITH (NOLOCK)
    WHERE CountryDictionaryId IS NOT NULL
    GROUP BY CountryDictionaryId
) m ON t.Id = m.Id
""",
# Пустой placeholder чтобы нумерация шагов не сбилась
"""SELECT 1""",

# Шаг 9: Plans
"""
INSERT INTO #DictVariants (Id, [Value], SourceEntity, PlansCountryCount)
SELECT Id, [Value], N'Plans.Country', Cnt
FROM (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Plans c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
WHERE NOT EXISTS (
    SELECT 1 FROM #DictVariants t
    WHERE ISNULL(t.Id,'') = ISNULL(src.Id,'')
      AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
)
""",
"""
UPDATE t
SET t.PlansCountryCount = src.Cnt,
    t.SourceEntity = CASE WHEN t.SourceEntity LIKE '%Plans.Country%'
                     THEN t.SourceEntity
                     ELSE t.SourceEntity + N', Plans.Country' END
FROM #DictVariants t
JOIN (
    SELECT
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'')    AS Id,
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') AS [Value],
        COUNT(*) AS Cnt
    FROM dbo.Plans c
    WHERE NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N'') IS NOT NULL
       OR NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'') IS NOT NULL
    GROUP BY
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryId)), N''),
        NULLIF(LTRIM(RTRIM(c.CountryDictionaryValue)), N'')
) src
    ON ISNULL(t.Id,'') = ISNULL(src.Id,'')
   AND ISNULL(t.[Value],'') = ISNULL(src.[Value],'')
WHERE t.PlansCountryCount = 0
""",

# Шаг 10: подтянуть IsDeleted/IsCurrent/Code
"""
UPDATE t
SET t.IsDeleted = d.IsDeleted,
    t.IsCurrent = d.IsCurrent,
    t.Code      = ISNULL(t.Code, d.Code)
FROM #DictVariants t
JOIN dbo.Dictionaries d
    ON CONVERT(NVARCHAR(36), d.Id) = t.Id
WHERE d.[Type] = N'CountryDictionary'
""",

# Финальный SELECT
"""
SELECT
    t.Id, t.[Value], t.SourceEntity, t.Code,
    t.IsDeleted, t.IsCurrent,
    t.CardsCount, t.CardCommissionCountryCount, t.PlacementCount,
    t.CardLanguageSchoolCountryCount, t.DogovorsCountryCount,
    t.MovementsCountryCount, t.PlansCountryCount, t.DictCount,
    (t.CardsCount + t.CardCommissionCountryCount + t.PlacementCount
   + t.CardLanguageSchoolCountryCount + t.DogovorsCountryCount
   + t.MovementsCountryCount + t.PlansCountryCount + t.DictCount) AS TotalCount
FROM #DictVariants t
ORDER BY
    (t.CardsCount + t.CardCommissionCountryCount + t.PlacementCount
   + t.CardLanguageSchoolCountryCount + t.DogovorsCountryCount
   + t.MovementsCountryCount + t.PlansCountryCount + t.DictCount) DESC,
    CASE WHEN t.Id IS NULL THEN 1 ELSE 0 END,
    t.Id, t.[Value]
"""
]

STEP_LABELS = [
    "Создание #DictVariants + индекс",
    "Индекс",
    "Dictionaries",
    "Cards.Country (INSERT)",
    "Cards.Country (UPDATE)",
    "Cards.CommissionCountry (INSERT)",
    "Cards.CommissionCountry (UPDATE)",
    "Cards.LanguageSchoolCountry (INSERT)",
    "Cards.LanguageSchoolCountry (UPDATE)",
    "Placements (INSERT)",
    "Placements (UPDATE)",
    "Dogovors (INSERT)",
    "Dogovors (UPDATE)",
    "Movements — считаем по Id (быстро)...",
    "Movements placeholder",
    "Plans (INSERT)",
    "Plans (UPDATE)",
    "IsDeleted/IsCurrent/Code",
    "Финальный SELECT",
]


# ── Excel стили ──────────────────────────────
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
OK_FILL      = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL    = PatternFill("solid", fgColor="FFF2CC")
ERROR_FILL   = PatternFill("solid", fgColor="FCE4D6")
DELETED_FONT = Font(name="Arial", color="999999", size=10, strike=True)
NORMAL_FONT  = Font(name="Arial", size=10)
BOLD_FONT    = Font(name="Arial", size=10, bold=True)
thin         = Side(style="thin", color="D9D9D9")
THIN_BORDER  = Border(left=thin, right=thin, top=thin, bottom=thin)


def build_excel(rows, output_path):
    wb = openpyxl.Workbook()

    # ── Лист 1: Все значения ─────────────────
    ws1 = wb.active
    ws1.title = "Анализ всех значений"
    ws1.freeze_panes = "A3"

    ws1.merge_cells("A1:O1")
    tc = ws1["A1"]
    tc.value = f"Анализ CountryDictionary — {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    tc.font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 24

    headers = [
        "Id (GUID)", "Название (Value)", "Источники", "Code",
        "Удалён", "Активен",
        "Cards\nCountry", "Cards\nCommission", "Placement",
        "Cards\nLangSchool", "Dogovors", "Movements", "Plans",
        "Dict", "ИТОГО"
    ]
    for col, h in enumerate(headers, 1):
        c = ws1.cell(row=2, column=col, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL
        c.alignment = HEADER_ALIGN; c.border = THIN_BORDER
    ws1.row_dimensions[2].height = 32

    needs_fix = []
    for r_idx, row in enumerate(rows, 3):
        (rid, val, src, code, is_del, is_cur,
         cards, commission, placement, langschool,
         dogovors, movements, plans, dcount, total) = row

        correct = COUNTRY_MAP.get(str(val).strip() if val else "") if val else None
        if is_del:
            row_fill, row_font = None, DELETED_FONT
        elif rid is None:
            row_fill, row_font = ERROR_FILL, NORMAL_FONT
        elif correct and correct != (val or "").strip():
            row_fill, row_font = WARN_FILL, NORMAL_FONT
            needs_fix.append((rid, val, correct, total))
        else:
            row_fill, row_font = OK_FILL, NORMAL_FONT

        data = [rid, val, src, code,
                "Да" if is_del else "Нет",
                "Да" if is_cur else "Нет",
                cards, commission, placement, langschool,
                dogovors, movements, plans, dcount, total]
        for col, v in enumerate(data, 1):
            c = ws1.cell(row=r_idx, column=col, value=v)
            c.border = THIN_BORDER; c.font = row_font
            if row_fill and not is_del: c.fill = row_fill
            c.alignment = Alignment(
                vertical="center",
                horizontal="center" if col >= 7 else "left",
                wrap_text=(col == 3)
            )

    col_widths = [38, 35, 45, 12, 8, 8, 9, 9, 9, 11, 10, 11, 10, 9, 9]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(rows)+2}"

    # ── Лист 2: Требуют исправления ──────────
    ws2 = wb.create_sheet("Требуют исправления")
    ws2.freeze_panes = "A3"
    ws2.merge_cells("A1:F1")
    t2 = ws2["A1"]
    t2.value = f"Требуют нормализации — {len(needs_fix)} шт."
    t2.font = Font(name="Arial", bold=True, size=12, color="C00000")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 24

    h2 = ["Id (GUID)", "Текущее название", "Правильное название", "Разница", "Всего записей", "Статус"]
    for col, h in enumerate(h2, 1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL
        c.alignment = HEADER_ALIGN; c.border = THIN_BORDER
    ws2.row_dimensions[2].height = 28

    for r_idx, (rid, val, correct, total) in enumerate(needs_fix, 3):
        for col, v in enumerate([rid, val, correct, "✓ Найдено в словаре", total, "Исправить"], 1):
            c = ws2.cell(row=r_idx, column=col, value=v)
            c.border = THIN_BORDER; c.font = NORMAL_FONT; c.fill = WARN_FILL
            c.alignment = Alignment(vertical="center",
                                    horizontal="center" if col in (5, 6) else "left")
        ws2.cell(row=r_idx, column=3).font = Font(name="Arial", size=10, color="375623")

    for i, w in enumerate([38, 35, 45, 20, 14, 14], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ── Лист 3: Легенда ───────────────────────
    ws3 = wb.create_sheet("Легенда")
    legend = [
        ("Цвет", "Значение"),
        ("🟢 Зелёный",         "Название корректное"),
        ("🟡 Жёлтый",          "Найдено в словаре — нужно исправить"),
        ("🔴 Красный",         "Нет Id (только текст, нет в Dictionaries)"),
        ("Серый зачёркнутый",  "IsDeleted = 1"),
    ]
    ws3.column_dimensions["A"].width = 25
    ws3.column_dimensions["B"].width = 55
    for r_idx, (a, b) in enumerate(legend, 1):
        ca = ws3.cell(row=r_idx, column=1, value=a)
        cb = ws3.cell(row=r_idx, column=2, value=b)
        f = BOLD_FONT if r_idx == 1 else NORMAL_FONT
        fill = PatternFill("solid", fgColor="D6E4F0") if r_idx == 1 else None
        ca.font = cb.font = f
        if fill: ca.fill = cb.fill = fill
        for c in (ca, cb):
            c.border = THIN_BORDER
            c.alignment = Alignment(vertical="center")

    wb.save(output_path)
    print(f"\n✅ Excel сохранён: {output_path}")
    print(f"   Всего строк:         {len(rows)}")
    print(f"   Требуют исправления: {len(needs_fix)}")


def main():
    print("Подключаемся к БД (Windows Auth)...")
    conn = get_connection()
    cursor = conn.cursor()
    cursor.fast_executemany = True

    total = len(STEPS)
    rows = []

    for i, sql in enumerate(STEPS):
        label = STEP_LABELS[i] if i < len(STEP_LABELS) else f"Шаг {i+1}"
        print(f"[{i+1}/{total}] {label}...", end=" ", flush=True)
        t0 = datetime.now()
        try:
            cursor.execute(sql.strip())
            # Последний шаг — SELECT
            if i == len(STEPS) - 1:
                rows = cursor.fetchall()
                print(f"→ {len(rows)} строк  ({(datetime.now()-t0).seconds}с)")
            else:
                print(f"✓  ({(datetime.now()-t0).seconds}с)")
        except Exception as e:
            print(f"\n❌ Ошибка на шаге {i+1} ({label}):\n   {e}")
            conn.close()
            return

    conn.close()

    print("\nФормируем Excel...")
    output = f"country_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    build_excel(rows, output)


if __name__ == "__main__":
    main()