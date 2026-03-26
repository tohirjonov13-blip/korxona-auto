"""
ndfl_processor.py — Парсинг файлов НДФЛ-отчёта и HR-реестров
Поддерживает:
  - Расчёт НДФЛ (годовой) — Приложения 1-7, лист Расчёт, Титул
  - Список ГПХ (из 1С)
  - Список приёма сотрудников (из 1С)
  - Список увольнений (из 1С)
  - Форма 1 (Баланс)
  - Форма 2 (ОПУ)
"""

import pandas as pd
import openpyxl
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
import logging

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────
# Структуры данных
# ─────────────────────────────────────────────────────────────────────

@dataclass
class Employee:
    """Сотрудник из Приложения 4 НДФЛ"""
    num: int
    name: str
    position: str
    pinfl: str
    date_start: str
    resident: int           # 1=резидент, 2=нерезидент
    status: int             # 1=работает, 2=уволен
    contract: int           # 1=осн, 2=совм, 3=ГПХ
    work_rate: str          # 0.25 / 0.5 / 1
    total_income: float     # гр.11 — общий доход за год
    salary_period: float    # гр.12 — ЗП в отчётном периоде
    ndfl_total: float       # гр.13 — НДФЛ всего
    ndfl_period: float      # гр.14 — НДФЛ за период

    @property
    def is_gph(self): return str(self.contract) == '3'
    @property
    def is_nonresident(self): return str(self.resident) == '2'
    @property
    def is_fired(self): return str(self.status) == '2'
    @property
    def name_upper(self): return self.name.strip().upper()


@dataclass
class PrizeEmployee:
    """Получатель приза/необлагаемого дохода — Приложение 5"""
    num: int
    name: str
    pinfl: str
    income: float           # необлагаемый доход
    льгота: str             # код льготы (102912 = ст.378)

    @property
    def name_upper(self): return self.name.strip().upper()


@dataclass
class NDFLCalcData:
    """Ключевые строки из листа 'Расчёт'"""
    total_income: float = 0        # 010 — общая сумма доходов
    labor_income: float = 0        # 011 — доходы в виде оплаты труда
    salary_period: float = 0       # 0110 — ЗП в отчётном периоде
    non_labor_income: float = 0    # 012 — доходы не связанные с ОТ
    exempt_income: float = 0       # 030 — освобождённые доходы
    tax_base: float = 0            # 050 — налоговая база
    ndfl_accrued: float = 0        # 060 — НДФЛ начисленный
    total_tax: float = 0           # 070 — итого НДФЛ+СН
    ndfl_inps: float = 0           # 080 — взносы ИНПС


@dataclass
class NDFLReport:
    """Полный разобранный НДФЛ-отчёт"""
    inn: str = ""
    headcount_avg: int = 0         # ср. численность
    headcount_nonbasic: int = 0    # из них не по осн. месту
    calc: NDFLCalcData = field(default_factory=NDFLCalcData)
    employees: list = field(default_factory=list)    # list[Employee]
    prize_employees: list = field(default_factory=list)  # list[PrizeEmployee]


@dataclass
class GphContract:
    """Договор ГПХ из реестра 1С"""
    name: str
    date_start: str
    date_end: str
    number: str = ""

    @property
    def name_upper(self): return self.name.strip().upper()


# ─────────────────────────────────────────────────────────────────────
# Вспомогательные функции
# ─────────────────────────────────────────────────────────────────────

def _sv(row: pd.Series, idx: int, default=None):
    """Safe value — читает ячейку pandas по индексу"""
    if idx >= len(row):
        return default
    v = row.iloc[idx]
    if pd.isna(v):
        return default
    return v


def _sf(row: pd.Series, idx: int) -> float:
    """Safe float"""
    v = _sv(row, idx)
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _si(row: pd.Series, idx: int) -> int:
    """Safe int"""
    v = _sv(row, idx)
    if v is None:
        return 0
    try:
        return int(float(str(v)))
    except (ValueError, TypeError):
        return 0


def _normalize_name(name: str) -> str:
    """Нормализация ФИО для сравнения"""
    return str(name).strip().upper()


# ─────────────────────────────────────────────────────────────────────
# Парсеры
# ─────────────────────────────────────────────────────────────────────

def parse_ndfl_report(path: str) -> NDFLReport:
    """
    Парсит годовой НДФЛ-отчёт (Расчёт НДФЛ и СН).
    Ожидаемые листы: Титульный лист, Расчет, Приложение 4, Приложение 5
    """
    report = NDFLReport()
    path = str(path)

    try:
        # ── Титульный лист ────────────────────────────────────────────
        df_tit = pd.read_excel(path, sheet_name='Титульный лист',
                               header=None, engine='openpyxl')
        for i in range(len(df_tit)):
            row = df_tit.iloc[i]
            text = ' '.join(str(v) for v in row if pd.notna(v))
            if 'ИНН' in text and not report.inn:
                for v in row:
                    s = str(v).strip()
                    if s.isdigit() and len(s) == 9:
                        report.inn = s
            if 'Численность работников в среднем' in text:
                for v in row:
                    if isinstance(v, (int, float)) and not pd.isna(v) and v > 0:
                        report.headcount_avg = int(v)
            if 'не по месту основной работы' in text:
                for v in row:
                    if isinstance(v, (int, float)) and not pd.isna(v) and v > 0:
                        report.headcount_nonbasic = int(v)

        # ── Расчёт ───────────────────────────────────────────────────
        df_calc = pd.read_excel(path, sheet_name='Расчет',
                                header=None, engine='openpyxl')
        code_to_field = {
            '010': 'total_income',
            '011': 'labor_income',
            '0110': 'salary_period',
            '012': 'non_labor_income',
            '030': 'exempt_income',
            '050': 'tax_base',
            '060': 'ndfl_accrued',
            '070': 'total_tax',
            '080': 'ndfl_inps',
        }
        for i in range(len(df_calc)):
            row = df_calc.iloc[i]
            code = str(_sv(row, 47, '')).strip()
            val  = _sv(row, 55)
            if code in code_to_field and isinstance(val, (int, float)) and not pd.isna(val):
                setattr(report.calc, code_to_field[code], float(val))

        # ── Приложение 4: сотрудники ─────────────────────────────────
        df4 = pd.read_excel(path, sheet_name='Приложение 4',
                            header=None, engine='openpyxl')
        # Колонки (0-based): 2=№, 4=ФИО, 15=должность, 22=ПИНФЛ,
        # 28=дата, 35=резидент, 41=статус, 47=контракт, 57=ставка,
        # 70=доход_всего, 76=ЗП_период, 80=НДФЛ_всего, 85=НДФЛ_период
        COL = dict(num=2, name=4, pos=15, pinfl=22, date=28,
                   res=35, status=41, contract=47, rate=57,
                   inc=70, sal=76, ndfl=80, ndfl_p=85)
        for i in range(9, 50):
            if i >= len(df4):
                break
            row = df4.iloc[i]
            raw_num = _sv(row, COL['num'])
            try:
                num = int(float(str(raw_num)))
            except (ValueError, TypeError):
                continue
            if not (1 <= num <= 200):
                continue
            name = str(_sv(row, COL['name'], '')).strip()
            if not name or name in ('2', 'Х', 'х'):
                continue

            report.employees.append(Employee(
                num=num,
                name=name,
                position=str(_sv(row, COL['pos'], '')).strip(),
                pinfl=str(_sv(row, COL['pinfl'], '')).strip(),
                date_start=str(_sv(row, COL['date'], '')).strip(),
                resident=_si(row, COL['res']) or 1,
                status=_si(row, COL['status']) or 1,
                contract=_si(row, COL['contract']) or 1,
                work_rate=str(_sv(row, COL['rate'], '1')).strip(),
                total_income=_sf(row, COL['inc']),
                salary_period=_sf(row, COL['sal']),
                ndfl_total=_sf(row, COL['ndfl']),
                ndfl_period=_sf(row, COL['ndfl_p']),
            ))

        # ── Приложение 5: призы/необлагаемые ────────────────────────
        df5 = pd.read_excel(path, sheet_name='Приложение 5',
                            header=None, engine='openpyxl')
        for i in range(9, 200):
            if i >= len(df5):
                break
            row = df5.iloc[i]
            raw_num = _sv(row, 2)
            try:
                num = int(float(str(raw_num)))
            except (ValueError, TypeError):
                continue
            if not (1 <= num <= 200):
                continue
            name = str(_sv(row, 4, '')).strip()
            if not name:
                continue
            report.prize_employees.append(PrizeEmployee(
                num=num,
                name=name,
                pinfl=str(_sv(row, 15, '')).strip(),
                income=_sf(row, 27),
                льгота=str(_sv(row, 52, '')).strip(),
            ))

        log.info(f"НДФЛ: ИНН={report.inn}, "
                 f"Прил.4={len(report.employees)}, "
                 f"Прил.5={len(report.prize_employees)}")

    except Exception as e:
        log.error(f"Ошибка парсинга НДФЛ: {e}")
        raise

    return report


def parse_gph_list(path: str) -> list:
    """
    Парсит список ГПХ-договоров из 1С.
    Формат: Дата | Номер | Сотрудник | Дата начала | Дата окончания | ...
    Возвращает list[GphContract]
    """
    df = pd.read_excel(path, engine='openpyxl')
    contracts = []
    name_col = next((c for c in df.columns if 'сотрудник' in c.lower()), None)
    start_col = next((c for c in df.columns if 'начал' in c.lower()), None)
    end_col   = next((c for c in df.columns if 'оконч' in c.lower()), None)
    num_col   = next((c for c in df.columns if 'номер' in c.lower()), None)

    if not name_col:
        log.warning("Не найдена колонка 'Сотрудник' в файле ГПХ")
        return contracts

    for _, row in df.iterrows():
        name = str(row.get(name_col, '')).strip()
        if not name or name.lower() in ('nan', ''):
            continue
        contracts.append(GphContract(
            name=name,
            date_start=str(row.get(start_col, '')).strip() if start_col else '',
            date_end=str(row.get(end_col, '')).strip() if end_col else '',
            number=str(row.get(num_col, '')).strip() if num_col else '',
        ))

    log.info(f"ГПХ: {len(contracts)} договоров")
    return contracts


def parse_hire_list(path: str) -> pd.DataFrame:
    """
    Парсит список принятых сотрудников из 1С.
    Ожидаемые колонки: Сотрудник, Дата приема, Подразделение,
                       Должность, Вид занятости
    """
    df = pd.read_excel(path, engine='openpyxl')
    # Нормализуем дату
    date_col = next((c for c in df.columns if 'прием' in c.lower() or 'приём' in c.lower()), None)
    if date_col:
        df['_date_parsed'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    log.info(f"Приём: {len(df)} записей")
    return df


def parse_fire_list(path: str) -> pd.DataFrame:
    """
    Парсит список уволенных сотрудников из 1С.
    Ожидаемые колонки: Дата, Номер, Номер приказа, Сотрудник,
                       Дата увольнения, Статья ТК РУ
    """
    df = pd.read_excel(path, engine='openpyxl')
    date_col = next((c for c in df.columns if 'уволь' in c.lower()), None)
    if date_col:
        df['_date_parsed'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    log.info(f"Увольнение: {len(df)} записей")
    return df


# ─────────────────────────────────────────────────────────────────────
# Маппинг → коды 1-korxona (Глава 9 и 10)
# ─────────────────────────────────────────────────────────────────────

def extract_korxona_personnel(report: NDFLReport,
                               df_fire: pd.DataFrame = None) -> dict:
    """
    Извлекает данные для заполнения Главы 9 и 10 отчёта 1-korxona.

    Глава 9 (Кадры):
      401 — численность для ЗП (Прил.4, без уволенных до н.г.)
      403 — ФОТ начисленный (Расчёт стр.011)
      404 — ФОТ работников с трудкнижками
      405 — численность с трудкнижками на конец года
      409 — среднегодовая с трудкнижками
      411 — внешние совместители (среднегодовые)
      412 — ГПХ-работники (среднегодовые)
      413 = 409 + 411 + 412
      416 — всего расходов на содержание рабочей силы

    Глава 10 (Выплаты):
      403 — ФОТ (повторяется)
      417 — проценты
      418 — дивиденды
      419 — материальная выгода
      420 — материальная помощь
      421 — авторское вознаграждение
      422 — выходное пособие
    """
    emps = report.employees

    # Основные работники (трудовой договор, основное место)
    main_workers  = [e for e in emps if str(e.contract) == '1']
    extern_part   = [e for e in emps if str(e.contract) == '2']  # совместители
    gph_workers   = [e for e in emps if str(e.contract) == '3']

    # Для 401: все в Прил.4 (включая ГПХ с доходом)
    code_401 = len([e for e in emps if e.total_income > 0])

    # 403 = ФОТ из строки 011 Расчёта
    code_403 = report.calc.labor_income

    # 404 = ФОТ только основных (с трудкнижками)
    code_404 = sum(e.total_income for e in main_workers + extern_part)

    # 405 = численность с трудкнижками на конец года (основные, не уволенные)
    code_405 = len([e for e in main_workers + extern_part
                    if not e.is_fired])

    # 409 = среднегодовая с трудкнижками
    # Если нет данных по помесячной — берём из титула
    code_409 = report.headcount_avg - report.headcount_nonbasic \
        if report.headcount_avg > 0 else len([e for e in main_workers if not e.is_fired])

    # 411 = внешние совместители среднегодовые
    code_411 = len(extern_part)

    # 412 = ГПХ среднегодовые
    code_412 = len(gph_workers)

    # 413 = 409 + 411 + 412
    code_413 = code_409 + code_411 + code_412

    # 416 = все расходы = ФОТ + приравненные выплаты (призы)
    code_416 = report.calc.total_income  # вся сумма начисленных доходов

    return {
        # Глава 9
        401: {'value': code_401, 'desc': 'Численность для исчисления ЗП',
              'source': 'Прил.4 НДФЛ (с доходом > 0)'},
        403: {'value': code_403, 'desc': 'ФОТ начисленный (всего)',
              'source': 'Расчёт НДФЛ стр.011'},
        404: {'value': code_404, 'desc': 'ФОТ работников с трудкнижками',
              'source': 'Прил.4: осн.+совм.'},
        405: {'value': code_405, 'desc': 'Численность с трудкнижками на конец года',
              'source': 'Прил.4: осн.+совм., не уволенные'},
        409: {'value': code_409, 'desc': 'Среднегодовая с трудкнижками',
              'source': 'Титул НДФЛ (ср.численность − совместители)'},
        411: {'value': code_411, 'desc': 'Внешние совместители (среднегодовые)',
              'source': 'Прил.4: контракт=2'},
        412: {'value': code_412, 'desc': 'ГПХ-работники (среднегодовые)',
              'source': 'Прил.4: контракт=3'},
        413: {'value': code_413, 'desc': 'Среднегодовая (вкл. совм. и ГПХ)',
              'source': '409 + 411 + 412'},
        416: {'value': code_416, 'desc': 'Всего расходов на рабочую силу',
              'source': 'Расчёт НДФЛ стр.010'},
        # Глава 10 (выплаты физлицам)
        417: {'value': 0, 'desc': 'Проценты', 'source': 'Ручной ввод'},
        418: {'value': 0, 'desc': 'Дивиденды', 'source': 'Ручной ввод'},
        419: {'value': 0, 'desc': 'Материальная выгода', 'source': 'Ручной ввод'},
        420: {'value': 0, 'desc': 'Материальная помощь', 'source': 'Ручной ввод'},
        421: {'value': 0, 'desc': 'Авторское вознаграждение', 'source': 'Ручной ввод'},
        422: {'value': 0, 'desc': 'Выходное пособие', 'source': 'Ручной ввод'},
    }
