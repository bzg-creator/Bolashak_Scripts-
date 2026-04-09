
import pyodbc
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ─────────────────────────────────────────────
# НАСТРОЙКИ
# ─────────────────────────────────────────────
DB_SERVER      = ""
DB_DATABASE    = ""
MAPPING_FILE   = "Country.xlsx"          # файл с маппингом рядом со скриптом
MAPPING_SHEET  = "Лист1 (2)"

# ─────────────────────────────────────────────
# Excel стили
# ─────────────────────────────────────────────
def make_font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

HDR_FONT   = make_font(bold=True, color="FFFFFF")
NORM_FONT  = make_font()
BOLD_FONT  = make_font(bold=True)

# Заливки
F_HEADER   = make_fill("1F4E79")
F_CHANGED  = make_fill("FFF2CC")   # жёлтый — нужно изменить
F_OK       = make_fill("E2EFDA")   # зелёный — уже правильно
F_NEW_VAL  = make_fill("DEEAF1")   # голубой — новое значение
F_RED      = make_fill("FCE4D6")   # красный — проблема
F_GRAY     = make_fill("F2F2F2")

thin  = Side(style="thin", color="D9D9D9")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

HDR_ALIGN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(vertical="center", horizontal="left",  wrap_text=False)
CTR_ALIGN  = Alignment(vertical="center", horizontal="center")


def cell(ws, row, col, value, fill=None, font=None, align=None, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font   = font  or NORM_FONT
    c.border = BORDER
    c.alignment = align or LEFT_ALIGN
    if fill:           c.fill = fill
    if number_format:  c.number_format = number_format
    return c


def write_header_row(ws, row, headers, fill=F_HEADER, height=28):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font      = HDR_FONT
        c.fill      = fill
        c.alignment = HDR_ALIGN
        c.border    = BORDER
    ws.row_dimensions[row].height = height


def add_sheet_title(ws, text, ncols):
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value     = text
    c.font      = Font(name="Arial", bold=True, size=11, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill      = make_fill("EBF3FB")
    ws.row_dimensions[1].height = 20


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ─────────────────────────────────────────────
# ПОДКЛЮЧЕНИЕ
# ─────────────────────────────────────────────
def get_conn():
    return pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        f"SERVER={DB_SERVER};DATABASE={DB_DATABASE};"
        "Trusted_Connection=yes;",
        timeout=0
    )


# ─────────────────────────────────────────────
# ЧТЕНИЕ МАППИНГА
# ─────────────────────────────────────────────
def load_mapping(filepath, sheet_name):
    """
    Возвращает список словарей:
    {
      num, new_id, new_value, old_id, old_value, source_entity,
      code, is_deleted, is_current,
      cards, commission, placement, langschool,
      dogovors, movements, plans, dict_count, total,
      needs_id_change, needs_value_change
    }
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[sheet_name]

    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if not any(v for v in row):
            continue
        (num, new_id, new_val, old_id, old_val, src,
         code, is_del, is_cur,
         cards, commission, placement, langschool,
         dogovors, movements, plans, dict_count, total) = row[:18]

        def s(v): return str(v or "").strip()
        def su(v): return s(v).upper()

        old_id_clean  = s(old_id)  if su(old_id)  not in ("NULL", "") else None
        new_id_clean  = s(new_id)  if su(new_id)  not in ("NULL", "") else None
        old_val_clean = s(old_val) if s(old_val)  not in ("NULL", "") else None
        new_val_clean = s(new_val) if s(new_val)  not in ("NULL", "") else None

        needs_id    = (old_id_clean is not None and
                       su(old_id_clean) != su(new_id_clean or ""))
        needs_val   = (old_val_clean is not None and
                       old_val_clean != (new_val_clean or ""))

        sources = [x.strip() for x in s(src).split(",") if x.strip() and x.strip() != "NULL"]

        rows.append({
            "num":          num,
            "new_id":       new_id_clean,
            "new_value":    new_val_clean,
            "old_id":       old_id_clean,
            "old_value":    old_val_clean,
            "sources":      sources,
            "code":         s(code) if s(code) != "NULL" else None,
            "is_deleted":   is_del,
            "is_current":   is_cur,
            "cards":        int(cards or 0),
            "commission":   int(commission or 0),
            "placement":    int(placement or 0),
            "langschool":   int(langschool or 0),
            "dogovors":     int(dogovors or 0),
            "movements":    int(movements or 0),
            "plans":        int(plans or 0),
            "dict_count":   int(dict_count or 0),
            "total":        int(total or 0),
            "needs_id":     needs_id,
            "needs_val":    needs_val,
        })

    print(f"  Маппинг загружен: {len(rows)} строк, "
          f"{sum(1 for r in rows if r['needs_id'] or r['needs_val'])} требуют изменений")
    return rows


# ─────────────────────────────────────────────
# ЗАПРОСЫ К БД
# ─────────────────────────────────────────────

TABLE_CONFIG = {
    # table_key: (sql_table, id_col, value_col, extra_select_cols)
    "Cards.Country": (
        "dbo.Cards",
        "CountryDictionaryId",
        "CountryDictionaryValue",
        ["Id AS CardId", "LTRIM(RTRIM(ISNULL(LastName,'')+' '+ISNULL(FirstName,''))) AS StudentName"]
    ),
    "Cards.CommissionCountry": (
        "dbo.Cards",
        "CommissionCountryDictionaryId",
        "CommissionCountryDictionaryValue",
        ["Id AS CardId", "LTRIM(RTRIM(ISNULL(LastName,'')+' '+ISNULL(FirstName,''))) AS StudentName"]
    ),
    "Cards.LanguageSchoolCountry": (
        "dbo.Cards",
        "LanguageSchoolCountryDictionaryId",
        "LanguageSchoolCountryDictionaryValue",
        ["Id AS CardId", "LTRIM(RTRIM(ISNULL(LastName,'')+' '+ISNULL(FirstName,''))) AS StudentName"]
    ),
    "Placement.Country": (
        "dbo.Placements",
        "CountryDictionaryId",
        "CountryDictionaryValue",
        ["Id AS PlacementId"]
    ),
    "Dogovors.Country": (
        "dbo.Dogovors",
        "CountryDictionaryId",
        "CountryDictionaryValue",
        ["Id AS DogovorId"]
    ),
    "Movements.Country": (
        "dbo.Movements",
        "CountryDictionaryId",
        "CountryDictionaryValue",
        ["Id AS MovementId"]
    ),
    "Plans.Country": (
        "dbo.Plans",
        "CountryDictionaryId",
        "CountryDictionaryValue",
        ["Id AS PlanId"]
    ),
}


def fetch_records_for_table(cursor, table_key, mapping_rows, max_rows=500):
    """
    Для каждой таблицы собираем записи которые нужно изменить.
    Для Movements — только COUNT (слишком много строк).
    Возвращает список строк для Excel.
    """
    if table_key not in TABLE_CONFIG:
        return []

    sql_table, id_col, val_col, extra_cols = TABLE_CONFIG[table_key]
    is_movements = (table_key == "Movements.Country")

    # Строим список пар (old_id, old_value, new_id, new_value) для этой таблицы
    pairs = []
    for r in mapping_rows:
        if table_key not in r["sources"]:
            continue
        if not (r["needs_id"] or r["needs_val"]):
            continue
        pairs.append(r)

    if not pairs:
        return []

    results = []

    for r in pairs:
        old_id  = r["old_id"]
        old_val = r["old_value"]
        new_id  = r["new_id"]
        new_val = r["new_value"]

        # WHERE условие
        conditions = []
        params     = []

        if old_id:
            conditions.append(f"LTRIM(RTRIM({id_col})) = ?")
            params.append(old_id)
        if old_val and old_id is None:
            # только если нет Id — ищем по тексту
            conditions.append(f"LTRIM(RTRIM({val_col})) = ?")
            params.append(old_val)

        if not conditions:
            continue

        where = " AND ".join(conditions)

        if is_movements:
            # Movements — только COUNT
            sql = f"SELECT COUNT(*) FROM {sql_table} WITH(NOLOCK) WHERE {where}"
            try:
                cursor.execute(sql, *params)
                cnt = cursor.fetchone()[0]
                if cnt > 0:
                    results.append({
                        "old_id":    old_id,
                        "old_value": old_val,
                        "new_id":    new_id,
                        "new_value": new_val,
                        "count":     cnt,
                        "record_id": f"[{cnt} записей]",
                        "extra":     {},
                    })
            except Exception as e:
                results.append({
                    "old_id": old_id, "old_value": old_val,
                    "new_id": new_id, "new_value": new_val,
                    "count": 0, "record_id": f"ОШИБКА: {e}", "extra": {},
                })
        else:
            # Остальные таблицы — получаем реальные записи
            extra_str = ", ".join(extra_cols)
            sql = (f"SELECT TOP {max_rows} {extra_str}, "
                   f"LTRIM(RTRIM({id_col})) AS CurId, "
                   f"LTRIM(RTRIM({val_col})) AS CurVal "
                   f"FROM {sql_table} WITH(NOLOCK) WHERE {where}")
            try:
                cursor.execute(sql, *params)
                fetched = cursor.fetchall()
                cols    = [d[0] for d in cursor.description]
                for row in fetched:
                    rec = dict(zip(cols, row))
                    extra = {k: v for k, v in rec.items()
                             if k not in ("CurId", "CurVal")}
                    results.append({
                        "old_id":    rec.get("CurId", old_id),
                        "old_value": rec.get("CurVal", old_val),
                        "new_id":    new_id,
                        "new_value": new_val,
                        "count":     1,
                        "record_id": list(extra.values())[0] if extra else "",
                        "extra":     extra,
                    })
            except Exception as e:
                results.append({
                    "old_id": old_id, "old_value": old_val,
                    "new_id": new_id, "new_value": new_val,
                    "count": 0, "record_id": f"ОШИБКА: {e}", "extra": {},
                })

    return results


def fetch_country_counts(cursor):
    """
    Считаем количество студентов (Cards) по каждой стране.
    Используем CountryDictionaryId — быстро через индекс.
    """
    print("  Считаем студентов по странам...", end=" ")
    sql = """
        SELECT
            c.CountryDictionaryId   AS DictId,
            c.CountryDictionaryValue AS CountryName,
            COUNT(*)                 AS StudentCount
        FROM dbo.Cards c WITH(NOLOCK)
        WHERE c.CountryDictionaryValue IS NOT NULL
          AND LTRIM(RTRIM(c.CountryDictionaryValue)) <> ''
        GROUP BY c.CountryDictionaryId, c.CountryDictionaryValue
        ORDER BY COUNT(*) DESC
    """
    cursor.execute(sql)
    rows = cursor.fetchall()
    print(f"{len(rows)} стран")
    return rows


# ─────────────────────────────────────────────
# ПОСТРОЕНИЕ EXCEL
# ─────────────────────────────────────────────

def build_excel(mapping, db_data, country_counts, output_path):
    wb = openpyxl.Workbook()

    # ── Лист 1: Маппинг ──────────────────────────────────────
    ws_map = wb.active
    ws_map.title = "1. Маппинг замен"
    ws_map.freeze_panes = "A3"

    changed = [r for r in mapping if r["needs_id"] or r["needs_val"]]
    add_sheet_title(ws_map, f"Маппинг замен — {len(changed)} записей требуют изменений  |  {datetime.now().strftime('%d.%m.%Y %H:%M')}", 10)

    h = ["#", "Старый Id (OldId)", "Старое название (OldValue)",
         "Новый Id", "Новое название (Value)",
         "Изменить Id?", "Изменить название?",
         "Таблицы", "Всего записей", "Статус"]
    write_header_row(ws_map, 2, h)

    for r_idx, r in enumerate(mapping, 3):
        needs  = r["needs_id"] or r["needs_val"]
        f_row  = F_CHANGED if needs else F_OK
        f_new  = F_NEW_VAL if needs else F_OK
        status = "Изменить" if needs else "Уже правильно"

        cell(ws_map, r_idx, 1,  r["num"],      f_row, align=CTR_ALIGN)
        cell(ws_map, r_idx, 2,  r["old_id"],   F_RED if r["needs_id"] else f_row)
        cell(ws_map, r_idx, 3,  r["old_value"],F_RED if r["needs_val"] else f_row)
        cell(ws_map, r_idx, 4,  r["new_id"],   f_new)
        cell(ws_map, r_idx, 5,  r["new_value"],f_new, font=BOLD_FONT)
        cell(ws_map, r_idx, 6,  "ДА" if r["needs_id"]  else "—", f_row, align=CTR_ALIGN)
        cell(ws_map, r_idx, 7,  "ДА" if r["needs_val"] else "—", f_row, align=CTR_ALIGN)
        cell(ws_map, r_idx, 8,  ", ".join(r["sources"]), f_row)
        cell(ws_map, r_idx, 9,  r["total"],    f_row, align=CTR_ALIGN)
        cell(ws_map, r_idx, 10, status,        F_CHANGED if needs else F_OK, align=CTR_ALIGN)

    set_col_widths(ws_map, [5, 38, 45, 38, 45, 12, 16, 60, 13, 14])
    ws_map.auto_filter.ref = f"A2:J{len(mapping)+2}"

    # ── Листы 2-6: по таблицам ───────────────────────────────
    TABLE_SHEETS = [
        ("2. Cards (Country)",      "Cards.Country",             ["CardId", "StudentName"]),
        ("3. Cards (Commission)",   "Cards.CommissionCountry",   ["CardId", "StudentName"]),
        ("4. Cards (LangSchool)",   "Cards.LanguageSchoolCountry",["CardId", "StudentName"]),
        ("5. Placements",           "Placement.Country",         ["PlacementId"]),
        ("6. Dogovors",             "Dogovors.Country",          ["DogovorId"]),
        ("7. Movements (кол-во)",   "Movements.Country",         []),
        ("8. Plans",                "Plans.Country",             ["PlanId"]),
    ]

    for sheet_name, table_key, extra_headers in TABLE_SHEETS:
        ws = wb.create_sheet(sheet_name)
        ws.freeze_panes = "A3"
        records = db_data.get(table_key, [])
        is_movements = (table_key == "Movements.Country")

        if is_movements:
            # Movements — только счётчики
            add_sheet_title(ws, f"{sheet_name} — COUNT по Id (данные не загружаются, 14 млн строк)", 6)
            h_mv = ["Старый Id", "Старое название", "Новый Id", "Новое название", "Кол-во записей", "Действие"]
            write_header_row(ws, 2, h_mv)
            for r_idx, rec in enumerate(records, 3):
                f = F_CHANGED if rec["count"] > 0 else F_OK
                cell(ws, r_idx, 1, rec["old_id"],    f)
                cell(ws, r_idx, 2, rec["old_value"],  F_RED)
                cell(ws, r_idx, 3, rec["new_id"],    F_NEW_VAL)
                cell(ws, r_idx, 4, rec["new_value"],  F_NEW_VAL, font=BOLD_FONT)
                cell(ws, r_idx, 5, rec["count"],      f, align=CTR_ALIGN)
                cell(ws, r_idx, 6, "UPDATE CountryDictionaryId + Value", f)
            set_col_widths(ws, [38, 45, 38, 45, 16, 38])
        else:
            total_recs = len(records)
            add_sheet_title(ws, f"{sheet_name} — {total_recs} записей требуют изменений", 6 + len(extra_headers))
            h_extra = extra_headers if extra_headers else []
            h_tbl   = h_extra + ["Старый Id", "Старое название", "Новый Id", "Новое название", "Действие"]
            write_header_row(ws, 2, h_tbl)
            ncols = len(h_tbl)

            for r_idx, rec in enumerate(records, 3):
                col_offset = 1
                for eh in h_extra:
                    val = rec["extra"].get(eh, "")
                    cell(ws, r_idx, col_offset, val, F_GRAY)
                    col_offset += 1
                cell(ws, r_idx, col_offset,   rec["old_id"],    F_RED)
                cell(ws, r_idx, col_offset+1, rec["old_value"],  F_RED)
                cell(ws, r_idx, col_offset+2, rec["new_id"],    F_NEW_VAL)
                cell(ws, r_idx, col_offset+3, rec["new_value"],  F_NEW_VAL, font=BOLD_FONT)
                cell(ws, r_idx, col_offset+4,
                     "UPDATE Id + Value" if rec["old_id"] else "UPDATE Value only",
                     F_CHANGED, align=CTR_ALIGN)

            extra_w = [15, 30] if len(extra_headers) == 2 else [15]
            set_col_widths(ws, extra_w + [38, 42, 38, 42, 22])

        if len(records) == 0:
            ws.cell(row=3, column=1, value="Нет записей требующих изменений").font = NORM_FONT

    # ── Лист 9: Количество студентов по странам ──────────────
    ws_cnt = wb.create_sheet("9. Кол-во студентов по странам")
    ws_cnt.freeze_panes = "A3"
    add_sheet_title(ws_cnt, f"Количество студентов по странам (из таблицы Cards)  —  {len(country_counts)} стран", 5)
    write_header_row(ws_cnt, 2,
        ["#", "Id страны (DictionaryId)", "Название страны", "Кол-во студентов", "% от общего"])

    total_students = sum(r[2] for r in country_counts)
    for r_idx, (dict_id, name, cnt) in enumerate(country_counts, 3):
        pct = round(cnt / total_students * 100, 2) if total_students else 0
        # Чередуем заливку для читаемости
        f = F_GRAY if r_idx % 2 == 0 else None
        cell(ws_cnt, r_idx, 1, r_idx - 2,   f, align=CTR_ALIGN)
        cell(ws_cnt, r_idx, 2, str(dict_id or "NULL"), f)
        cell(ws_cnt, r_idx, 3, name,          f, font=BOLD_FONT if r_idx <= 12 else NORM_FONT)
        cell(ws_cnt, r_idx, 4, cnt,           f, align=CTR_ALIGN)
        pct_cell = cell(ws_cnt, r_idx, 5, pct / 100, f, align=CTR_ALIGN)
        pct_cell.number_format = "0.00%"

    # Итого
    last_row = len(country_counts) + 3
    ws_cnt.cell(row=last_row, column=3, value="ИТОГО").font = BOLD_FONT
    cell(ws_cnt, last_row, 4, total_students, font=BOLD_FONT, align=CTR_ALIGN).fill = make_fill("BDD7EE")
    cell(ws_cnt, last_row, 5, 1.0, align=CTR_ALIGN).number_format = "0.00%"

    set_col_widths(ws_cnt, [5, 38, 42, 18, 14])
    ws_cnt.auto_filter.ref = f"A2:E{len(country_counts)+2}"

    wb.save(output_path)
    print(f"\n✅ Excel сохранён: {output_path}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    print("=" * 55)
    print("normalize_preview.py  — только чтение, БД не меняется")
    print("=" * 55)

    # 1. Загружаем маппинг
    print("\n[1/3] Читаем маппинг из Country.xlsx...")
    mapping = load_mapping(MAPPING_FILE, MAPPING_SHEET)

    # 2. Подключаемся к БД и собираем данные
    print("\n[2/3] Подключаемся к БД и собираем данные...")
    conn   = get_conn()
    cursor = conn.cursor()

    db_data = {}
    table_keys = [
        "Cards.Country",
        "Cards.CommissionCountry",
        "Cards.LanguageSchoolCountry",
        "Placement.Country",
        "Dogovors.Country",
        "Movements.Country",
        "Plans.Country",
    ]

    for tk in table_keys:
        print(f"  {tk}...", end=" ")
        records = fetch_records_for_table(cursor, tk, mapping)
        db_data[tk] = records
        print(f"{len(records)} записей")

    print("  Количество студентов по странам...", end=" ")
    country_counts = fetch_country_counts(cursor)

    conn.close()

    # 3. Строим Excel
    print("\n[3/3] Формируем Excel...")
    output = f"normalization_preview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    build_excel(mapping, db_data, country_counts, output)

    # Итоговая статистика
    total_changes = sum(len(v) for v in db_data.values())
    print(f"\n{'='*55}")
    print(f"Итого записей требующих изменений: {total_changes}")
    for tk in table_keys:
        cnt = len(db_data[tk])
        if cnt > 0:
            print(f"  {tk}: {cnt}")
    print(f"{'='*55}")


if __name__ == "__main__":
    main()
