"""
processor.py — Основной модуль обработки финансовых данных для отчёта 1-korxona
Поддерживает: Форма 1 (Баланс), Форма 2 (ОПУ), ОСВ из 1С (Excel .xlsx)
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import logging
import re

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# МАППИНГ: код строки → источник данных
# ─────────────────────────────────────────────

MAPPING = {
    # ── Глава 1: Доходы ──────────────────────────────────────────────────
    # Код 101: Оборот (выручка без НДС) → Ф2, строка 010 (Чистая выручка)
    101: {"source": "f2", "row_code": "010", "col": "year", "desc": "Оборот организации (без НДС)"},

    # Код 102: Продукция собственного производства → Ф2 строка 020 (если есть)
    102: {"source": "f2", "row_code": "020", "col": "year", "desc": "Продукция собственного производства"},

    # Код 103: Прочие доходы → Ф2, строка 200 (Прочие доходы)
    103: {"source": "f2", "row_code": "200", "col": "year", "desc": "Прочие доходы"},

    # Код 104: Доходы от оперативной аренды → Ф2 строка 210
    104: {"source": "f2", "row_code": "210", "col": "year", "desc": "Доходы от оперативной аренды"},

    # Код 105: Дивиденды полученные → Ф2 строка 220
    105: {"source": "f2", "row_code": "220", "col": "year", "desc": "Доходы в виде дивидендов"},

    # Код 106: Целевые поступления → Ф2 строка 230
    106: {"source": "f2", "row_code": "230", "col": "year", "desc": "Целевые поступления"},

    # Код 109: Общие доходы = 101+102+103 (рассчитывается автоматически)
    109: {"source": "calc", "formula": "101+102+103", "desc": "Общие доходы"},

    # ── Глава 2: Затраты ──────────────────────────────────────────────────
    # Код 110: Затраты всего → Ф2 строка 400 (Совокупные расходы)
    110: {"source": "f2", "row_code": "400", "col": "year", "desc": "Затраты – всего"},

    # Код 111: Себестоимость и расходы периода → Ф2 строка 040+090+110
    111: {"source": "f2_sum", "row_codes": ["040", "090", "110"], "col": "year", "desc": "Себестоимость и расходы периода"},

    # Код 112: Стоимость товаров для перепродажи → Ф2 строка 040
    112: {"source": "f2", "row_code": "040", "col": "year", "desc": "Стоимость товаров для перепродажи"},

    # Код 113: Материальные затраты → ОСВ счета 20,23,25,26 (субстатьи материалов)
    113: {"source": "osv", "accounts": ["2010", "2310", "2510", "2610"], "desc": "Материальные затраты"},

    # Код 117: Затраты на оплату труда → ОСВ счёт 6710 (оборот по Дт) или Ф2
    117: {"source": "osv", "accounts": ["6710"], "col": "debit_turnover", "desc": "Затраты на оплату труда"},

    # Код 118: ЕСН (отчисления соцстрах) → ОСВ счёт 6520
    118: {"source": "osv", "accounts": ["6520"], "col": "debit_turnover", "desc": "Отчисления на соц. страхование"},

    # Код 119: Амортизация ОС и НМА → ОСВ счёт 0200 оборот или Ф1
    119: {"source": "osv", "accounts": ["0200", "0400"], "col": "credit_turnover", "desc": "Амортизация ОС и НМА"},

    # ── Глава 3: Запасы ───────────────────────────────────────────────────
    # Код 140: Производственные запасы → Ф1 строка 140 (нач./кон.)
    140: {"source": "f1", "row_code": "140", "col": "both", "desc": "Производственные запасы"},

    # Код 141: Незавершённое производство → Ф1 строка 150
    141: {"source": "f1", "row_code": "150", "col": "both", "desc": "Незавершённое производство"},

    # Код 142: Готовая продукция → Ф1 строка 160
    142: {"source": "f1", "row_code": "160", "col": "both", "desc": "Готовая продукция"},

    # Код 143: Товары → Ф1 строка 170
    143: {"source": "f1", "row_code": "170", "col": "both", "desc": "Товары"},

    # ── Глава 4: ИКТ ──────────────────────────────────────────────────────
    # Код 150: Затраты на ИКТ → ОСВ счёт 9400 субсчета ИКТ (ручной ввод)
    150: {"source": "manual", "desc": "Затраты на ИКТ – всего"},

    # ── Глава 5: Основные средства ────────────────────────────────────────
    # Код 160: ОС по первоначальной стоимости → Ф1 строка 010 (нач./кон.)
    160: {"source": "f1", "row_code": "010", "col": "both", "desc": "ОС по первоначальной стоимости"},

    # Код 161: Сумма износа ОС → Ф1 строка 011 (нач./кон.)
    161: {"source": "f1", "row_code": "011", "col": "both", "desc": "Сумма износа ОС"},

    # Код 162: Незавершённое строительство → Ф1 строка 030
    162: {"source": "f1", "row_code": "030", "col": "both", "desc": "Незавершённое строительство"},

    # Код 163: НМА по первоначальной стоимости → Ф1 строка 040
    163: {"source": "f1", "row_code": "040", "col": "both", "desc": "НМА по первоначальной стоимости"},

    # Код 164: Амортизация НМА → Ф1 строка 041
    164: {"source": "f1", "row_code": "041", "col": "both", "desc": "Амортизация НМА"},

    # ── Глава 9: Персонал ─────────────────────────────────────────────────
    # Коды 401–416 — из отдельного файла/ручного ввода по персоналу
    401: {"source": "personnel", "field": "avg_headcount_for_salary", "col": "year", "desc": "Численность для исчисления ЗП"},
    403: {"source": "personnel", "field": "total_wage_fund", "col": "year", "desc": "Начисленные доходы (ФОТ)"},
    404: {"source": "personnel", "field": "wage_fund_with_workbooks", "col": "year", "desc": "ФОТ работников с трудовыми книжками"},
    405: {"source": "personnel", "field": "headcount_with_workbooks_endyear", "col": "year", "desc": "Численность с трудкнижками на конец года"},
    409: {"source": "personnel", "field": "avg_headcount_with_workbooks", "col": "year", "desc": "Среднегодовая численность с трудкнижками"},
    411: {"source": "personnel", "field": "avg_external_parttime", "col": "year", "desc": "Внешние совместители (среднегодовые)"},
    412: {"source": "personnel", "field": "avg_gph_workers", "col": "year", "desc": "Работники по ГПХ (среднегодовые)"},
    # Код 413 = 409 + 411 + 412 (рассчитывается)
    413: {"source": "calc", "formula": "409+411+412", "desc": "Среднегодовая численность (включая совместителей и ГПХ)"},
    416: {"source": "personnel", "field": "total_labor_costs", "col": "year", "desc": "Всего расходов на содержание рабочей силы"},

    # ── Глава 10: Выплаты физическим лицам ───────────────────────────────
    417: {"source": "personnel", "field": "interests_paid", "col": "year", "desc": "Проценты"},
    418: {"source": "personnel", "field": "dividends_paid", "col": "year", "desc": "Дивиденды"},
    419: {"source": "personnel", "field": "material_benefit", "col": "year", "desc": "Материальная выгода"},
    420: {"source": "personnel", "field": "material_aid", "col": "year", "desc": "Материальная помощь"},
    421: {"source": "personnel", "field": "author_fees", "col": "year", "desc": "Авторское вознаграждение"},
    422: {"source": "personnel", "field": "severance_pay", "col": "year", "desc": "Выходное пособие"},
}

# Контрольные соотношения (формулы)
CONTROL_RATIOS = {
    "109 = 101+102+103": lambda r: abs(r.get(109, 0) - (r.get(101, 0) + r.get(102, 0) + r.get(103, 0))) < 1,
    "413 = 409+411+412": lambda r: abs(r.get(413, 0) - (r.get(409, 0) + r.get(411, 0) + r.get(412, 0))) < 1,
}


class KorxonaProcessor:
    """Основной класс обработки данных для 1-korxona"""

    def __init__(self, f1_path=None, f2_path=None, osv_path=None, personnel_path=None):
        self.f1_path = Path(f1_path) if f1_path else None
        self.f2_path = Path(f2_path) if f2_path else None
        self.osv_path = Path(osv_path) if osv_path else None
        self.personnel_path = Path(personnel_path) if personnel_path else None

        self.f1_data = {}      # {row_code: {"begin": val, "end": val}}
        self.f2_data = {}      # {row_code: val}
        self.osv_data = {}     # {account: {"debit_turnover": val, "credit_turnover": val, ...}}
        self.personnel_data = {}  # {field: val}
        self.results = {}      # {code: {"begin": val, "end": val, "year": val}}
        self.warnings = []

    # ──────────────────────────────────────────────────────────────────────
    # ПАРСЕРЫ
    # ──────────────────────────────────────────────────────────────────────

    def _find_header_row(self, df, keywords):
        """Найти строку-заголовок по ключевым словам"""
        for i, row in df.iterrows():
            row_str = " ".join(str(v) for v in row.values if pd.notna(v)).lower()
            if all(kw.lower() in row_str for kw in keywords):
                return i
        return None

    def _normalize_code(self, val):
        """Нормализовать код строки к строке '010', '040' и т.д."""
        if pd.isna(val):
            return None
        s = str(val).strip().split(".")[0].split(",")[0]
        s = re.sub(r"[^0-9]", "", s)
        return s.zfill(3) if s else None

    def _read_xltx(self, path):
        """Читает Excel/xltx — пробует лист list02 (формат 1С), иначе первый лист с данными."""
        path = str(path)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()
        # list02 обычно содержит данные в формате 1С
        for sh in ['list02', 'list01'] + sheets:
            if sh not in sheets:
                continue
            try:
                df = pd.read_excel(path, sheet_name=sh, header=None, engine='openpyxl')
                # Есть ли числа?
                has_nums = any(isinstance(v,(int,float)) and pd.notna(v) and v not in (0,)
                               for row in df.values for v in row)
                if has_nums:
                    return df
            except Exception:
                continue
        return pd.read_excel(path, sheet_name=sheets[0], header=None, engine='openpyxl')

    def _find_code_col(self, df):
        """Находит индекс колонки с кодами строк (010, 020 и т.д.)"""
        best_col, best_count = 0, 0
        for ci in range(min(df.shape[1], 6)):
            count = sum(1 for v in df.iloc[:,ci]
                        if self._normalize_code(v) and len(self._normalize_code(v)) >= 2)
            if count > best_count:
                best_col, best_count = ci, count
        return best_col if best_count > 0 else None

    def parse_f1(self):
        """Парсинг Формы 1 (Бухгалтерский баланс).
        Поддерживает .xlsx и .xltx (шаблоны 1С с листами list01/list02).
        Структура: [Наименование, Код, Нач.года, Кон.года]
        """
        if not self.f1_path or not self.f1_path.exists():
            log.warning("Форма 1 не предоставлена, пропускаем.")
            return

        df = self._read_xltx(self.f1_path)
        log.info(f"Форма 1: загружено {len(df)} строк")

        code_col = self._find_code_col(df)
        if code_col is None:
            log.warning("Форма 1: не найдена колонка с кодами строк")
            return

        for _, row in df.iterrows():
            vals = list(row.values)
            if code_col >= len(vals):
                continue
            code = self._normalize_code(vals[code_col])
            if not code or len(code) > 4:
                continue
            # Числа правее колонки кода
            nums = [v for v in vals[code_col+1:code_col+5]
                    if isinstance(v,(int,float)) and pd.notna(v)]
            self.f1_data[code] = {
                "begin": nums[0] if len(nums) > 0 else 0,
                "end":   nums[1] if len(nums) > 1 else 0,
            }

        log.info(f"Форма 1: распознано {len(self.f1_data)} строк")

    def parse_f2(self):
        """Парсинг Формы 2 (ОПУ).
        Поддерживает .xlsx и .xltx.
        В xltx-формате 1С: код в col 2, доходы отч.года в col 5, расходы в col 6.
        В обычном xlsx: код рядом с первым числом.
        """
        if not self.f2_path or not self.f2_path.exists():
            log.warning("Форма 2 не предоставлена, пропускаем.")
            return

        df = self._read_xltx(self.f2_path)
        log.info(f"Форма 2: загружено {len(df)} строк")

        code_col = self._find_code_col(df)
        if code_col is None:
            log.warning("Форма 2: не найдена колонка с кодами строк")
            return

        for _, row in df.iterrows():
            vals = list(row.values)
            if code_col >= len(vals):
                continue
            code = self._normalize_code(vals[code_col])
            if not code or len(code) > 4:
                continue

            # Для xltx 1С: доход идёт перед расходом в парах колонок
            # Берём ВСЕ числа правее кода, выбираем первое ненулевое
            nums_with_idx = [(i, v) for i, v in enumerate(vals[code_col+1:code_col+8], code_col+1)
                             if isinstance(v,(int,float)) and pd.notna(v) and str(v) not in ('x',)]
            if nums_with_idx:
                self.f2_data[code] = nums_with_idx[0][1]

        log.info(f"Форма 2: распознано {len(self.f2_data)} строк")

    def parse_osv(self):
        """Парсинг ОСВ (Оборотно-сальдовая ведомость из 1С).
        Ожидаемая структура: [Счёт, Сальдо нач Дт, Сальдо нач Кт, Оборот Дт, Оборот Кт, Сальдо кон Дт, Сальдо кон Кт]
        """
        if not self.osv_path or not self.osv_path.exists():
            log.warning("ОСВ не предоставлена, пропускаем.")
            return

        df = pd.read_excel(self.osv_path, header=None, engine='openpyxl')
        log.info(f"ОСВ: загружено {len(df)} строк")

        for _, row in df.iterrows():
            vals = list(row.values)
            # Первая колонка — счёт
            acct_raw = str(vals[0]).strip() if pd.notna(vals[0]) else ""
            acct = re.sub(r"[^0-9]", "", acct_raw)
            if not acct or len(acct) < 2:
                continue
            nums = [float(x) if isinstance(x, (int, float)) and pd.notna(x) else 0 for x in vals[1:]]
            if len(nums) >= 6:
                self.osv_data[acct] = {
                    "begin_debit": nums[0], "begin_credit": nums[1],
                    "debit_turnover": nums[2], "credit_turnover": nums[3],
                    "end_debit": nums[4], "end_credit": nums[5],
                }
            elif len(nums) >= 2:
                self.osv_data[acct] = {
                    "begin_debit": 0, "begin_credit": 0,
                    "debit_turnover": nums[0], "credit_turnover": nums[1],
                    "end_debit": 0, "end_credit": 0,
                }

        log.info(f"ОСВ: распознано {len(self.osv_data)} счетов")

    def parse_personnel(self):
        """Парсинг данных по персоналу (отдельный Excel).
        Ожидаемый формат: два столбца — [Показатель / field_name, Значение]
        """
        if not self.personnel_path or not self.personnel_path.exists():
            log.warning("Файл персонала не предоставлен, пропускаем.")
            return

        df = pd.read_excel(self.personnel_path, header=None, engine='openpyxl')
        # Поддерживаем как словарный формат (ключ-значение), так и поиск по ключевым словам
        for _, row in df.iterrows():
            if len(row) >= 2:
                key = str(row.iloc[0]).strip().lower()
                val = row.iloc[1]
                if pd.notna(val):
                    self.personnel_data[key] = float(val) if isinstance(val, (int, float)) else val

        log.info(f"Персонал: загружено {len(self.personnel_data)} полей")

    # ──────────────────────────────────────────────────────────────────────
    # ВЫЧИСЛЕНИЕ ЗНАЧЕНИЙ
    # ──────────────────────────────────────────────────────────────────────

    def _get_f1(self, row_code, col="both"):
        data = self.f1_data.get(row_code.zfill(3), {})
        if col == "both":
            return {"begin": data.get("begin", 0), "end": data.get("end", 0)}
        return data.get(col, 0)

    def _get_f2(self, row_code):
        return self.f2_data.get(row_code.zfill(3), 0)

    def _get_osv_sum(self, accounts, col="credit_turnover"):
        total = 0
        for acct in accounts:
            # Поиск по точному совпадению и по префиксу
            for key, vals in self.osv_data.items():
                if key == acct or key.startswith(acct):
                    total += vals.get(col, 0)
        return total

    def _get_personnel(self, field):
        # Прямой поиск или поиск по ключевым словам
        val = self.personnel_data.get(field, None)
        if val is None:
            # Поиск по частичному совпадению
            for k, v in self.personnel_data.items():
                if field.lower() in k:
                    return float(v) if isinstance(v, (int, float)) else 0
        return float(val) if val is not None else 0

    def compute(self):
        """Вычислить все значения согласно маппингу"""
        computed = {}

        for code, cfg in MAPPING.items():
            src = cfg["source"]
            desc = cfg["desc"]

            try:
                if src == "f1":
                    col = cfg.get("col", "both")
                    if col == "both":
                        v = self._get_f1(cfg["row_code"], "both")
                        computed[code] = {"begin": v["begin"], "end": v["end"], "year": None}
                    else:
                        computed[code] = {"year": self._get_f1(cfg["row_code"], col)}

                elif src == "f2":
                    computed[code] = {"year": self._get_f2(cfg["row_code"])}

                elif src == "f2_sum":
                    total = sum(self._get_f2(rc) for rc in cfg["row_codes"])
                    computed[code] = {"year": total}

                elif src == "osv":
                    col = cfg.get("col", "credit_turnover")
                    total = self._get_osv_sum(cfg["accounts"], col)
                    computed[code] = {"year": total}

                elif src == "personnel":
                    computed[code] = {"year": self._get_personnel(cfg["field"])}

                elif src == "manual":
                    computed[code] = {"year": 0}  # заполняется вручную
                    self.warnings.append(f"Код {code} ({desc}): требует ручного ввода")

                elif src == "calc":
                    # Рассчитывается после первого прохода
                    computed[code] = {"year": None, "_formula": cfg["formula"]}

            except Exception as e:
                log.warning(f"Код {code} ({desc}): ошибка — {e}")
                computed[code] = {"year": 0}

        # Второй проход: вычисляем формулы
        for code, vals in computed.items():
            if vals.get("_formula"):
                formula = vals["_formula"]
                parts = re.findall(r"\d+", formula)
                total = 0
                for p in parts:
                    dep = computed.get(int(p), {})
                    total += dep.get("year") or 0
                computed[code] = {"year": total}

        self.results = computed
        return computed

    # ──────────────────────────────────────────────────────────────────────
    # ВАЛИДАЦИЯ
    # ──────────────────────────────────────────────────────────────────────

    def validate(self):
        """Проверка контрольных соотношений"""
        # Плоский словарь для проверки
        flat = {code: (v.get("year") or 0) for code, v in self.results.items()}
        errors = []
        for name, check_fn in CONTROL_RATIOS.items():
            if not check_fn(flat):
                val_parts = {}
                for p in re.findall(r"\d+", name):
                    val_parts[int(p)] = flat.get(int(p), 0)
                errors.append({"ratio": name, "values": val_parts})
        return errors

    # ──────────────────────────────────────────────────────────────────────
    # ЭКСПОРТ В ШАБЛОН
    # ──────────────────────────────────────────────────────────────────────

    def fill_template(self, template_path: str, output_path: str):
        """Заполнить Excel-шаблон Госкомстата вычисленными значениями.

        Алгоритм поиска ячейки:
        1. Сканируем все строки листа
        2. Ищем ячейку со значением == код строки (целое число или строка)
        3. Значение вставляем в соседние ячейки (по колонкам)
        """
        wb = load_workbook(template_path)

        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    # Попытка трактовать значение ячейки как код строки
                    try:
                        cell_code = int(str(cell.value).strip().split(".")[0])
                    except (ValueError, AttributeError):
                        continue

                    if cell_code not in self.results:
                        continue

                    mapping_cfg = MAPPING[cell_code]
                    result = self.results[cell_code]
                    col_idx = cell.column

                    # Определяем колонки для заполнения
                    if mapping_cfg.get("col") == "both" or ("begin" in result and "end" in result):
                        # Ищем 2 числовых ячейки правее: начало и конец года
                        filled = 0
                        for offset in range(1, 10):
                            neighbor = sheet.cell(row=cell.row, column=col_idx + offset)
                            if neighbor.value is None or isinstance(neighbor.value, (int, float)):
                                if filled == 0 and result.get("begin") is not None:
                                    neighbor.value = result["begin"]
                                    filled += 1
                                elif filled == 1 and result.get("end") is not None:
                                    neighbor.value = result["end"]
                                    filled += 1
                                    break
                    else:
                        # Ищем первую пустую числовую ячейку правее
                        for offset in range(1, 10):
                            neighbor = sheet.cell(row=cell.row, column=col_idx + offset)
                            if neighbor.value is None or isinstance(neighbor.value, (int, float)):
                                neighbor.value = result.get("year") or 0
                                break

        wb.save(output_path)
        log.info(f"Шаблон сохранён: {output_path}")

    # ──────────────────────────────────────────────────────────────────────
    # ПОЛНЫЙ ЦИКЛ
    # ──────────────────────────────────────────────────────────────────────

    def run(self, template_path=None, output_path=None):
        """Запустить полный цикл обработки"""
        self.parse_f1()
        self.parse_f2()
        self.parse_osv()
        self.parse_personnel()
        self.compute()
        errors = self.validate()

        if template_path and output_path:
            self.fill_template(template_path, output_path)

        return {
            "results": self.results,
            "validation_errors": errors,
            "warnings": self.warnings,
        }
