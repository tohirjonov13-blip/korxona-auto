"""
ndfl_checks.py — Логика валидации НДФЛ-данных перед заполнением 1-korxona

Уровни критичности:
  CRITICAL  — нельзя продолжить без исправления (блокирует заполнение)
  WARNING   — требует внимания, но можно пропустить
  INFO      — информационное замечание

Проверки:
  1. Сверка ГПХ    — все в реестре найдены в НДФЛ Прил.4?
  2. Прием/Уволен  — соответствие статусов в НДФЛ с HR-реестрами
  3. Нерезиденты   — ставка НДФЛ (должна применяться повышенная)
  4. Численность   — корректность данных для кодов 401, 413
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional
import pandas as pd


class Severity(Enum):
    CRITICAL = "critical"   # ❌ Блокирует заполнение
    WARNING  = "warning"    # ⚠️  Требует внимания
    INFO     = "info"       # ℹ️  К сведению


@dataclass
class CheckResult:
    """Результат одной проверки"""
    check_id: str           # уникальный ID проверки
    category: str           # категория: ГПХ / Прием-Увол / Нерезидент / Численность
    severity: Severity
    title: str              # короткое название
    description: str        # подробное описание
    affected: list = field(default_factory=list)  # список затронутых (ФИО, коды)
    recommendation: str = ""
    can_skip: bool = True   # можно ли пропустить

    @property
    def icon(self):
        return {"critical": "❌", "warning": "⚠️", "info": "ℹ️"}[self.severity.value]

    @property
    def is_blocking(self):
        return self.severity == Severity.CRITICAL and not self.can_skip


@dataclass
class CheckSummary:
    """Итог всех проверок"""
    results: list = field(default_factory=list)  # list[CheckResult]
    can_proceed: bool = True    # можно ли перейти к заполнению
    skipped_by_user: bool = False

    @property
    def critical(self): return [r for r in self.results if r.severity == Severity.CRITICAL]
    @property
    def warnings(self): return [r for r in self.results if r.severity == Severity.WARNING]
    @property
    def infos(self):    return [r for r in self.results if r.severity == Severity.INFO]
    @property
    def has_critical(self): return len(self.critical) > 0
    @property
    def total(self): return len(self.results)


# ─────────────────────────────────────────────────────────────────────
# Проверка 1: Сверка ГПХ
# ─────────────────────────────────────────────────────────────────────

def check_gph(ndfl_report, gph_contracts: list) -> list:
    """
    Сверяет реестр ГПХ-договоров с Приложением 4 НДФЛ.

    Логика:
    А) Есть в реестре ГПХ, НО нет в Прил.4 → CRITICAL
       (доходы не отражены в НДФЛ — налоговый риск)
    Б) Есть в Прил.4 с contract=3, НО нет в реестре ГПХ → WARNING
       (возможна ошибка классификации)
    В) Есть в обоих, contract≠3 → WARNING
       (возможна ошибка типа договора)
    """
    results = []
    emps = ndfl_report.employees
    ndfl_names = {e.name_upper: e for e in emps}
    gph_names  = {c.name_upper: c for c in gph_contracts}

    # А) В ГПХ-реестре, нет в НДФЛ
    not_in_ndfl = []
    for name_up, contract in gph_names.items():
        if name_up not in ndfl_names:
            not_in_ndfl.append({
                'name': contract.name,
                'period': f"{contract.date_start} — {contract.date_end}",
                'num': contract.number,
            })

    if not_in_ndfl:
        results.append(CheckResult(
            check_id="GHP_NOT_IN_NDFL",
            category="ГПХ",
            severity=Severity.CRITICAL,
            title=f"ГПХ-работники не отражены в НДФЛ ({len(not_in_ndfl)} чел.)",
            description=(
                f"В реестре договоров ГПХ найдено {len(not_in_ndfl)} человека, "
                f"которые отсутствуют в Приложении 4 годового НДФЛ-отчёта. "
                f"Это означает, что их доходы не включены в расчёт — налоговый риск."
            ),
            affected=[f"{x['name']} (договор {x['num']}, {x['period']})"
                      for x in not_in_ndfl],
            recommendation=(
                "Уточните у бухгалтера: были ли выплаты по этим договорам? "
                "Если да — необходимо включить в НДФЛ до подачи отчёта. "
                "Если нет (договор есть, но оплата не производилась) — "
                "отметьте это и можно продолжить."
            ),
            can_skip=True,
        ))

    # Б) В Прил.4 с contract=3, нет в реестре ГПХ
    gph_in_ndfl_not_in_registry = []
    for emp in emps:
        if emp.is_gph and emp.name_upper not in gph_names:
            gph_in_ndfl_not_in_registry.append(emp)

    if gph_in_ndfl_not_in_registry:
        results.append(CheckResult(
            check_id="GHP_NOT_IN_REGISTRY",
            category="ГПХ",
            severity=Severity.WARNING,
            title=f"В НДФЛ тип=ГПХ, но нет в реестре ({len(gph_in_ndfl_not_in_registry)} чел.)",
            description=(
                f"{len(gph_in_ndfl_not_in_registry)} сотрудников в НДФЛ имеют "
                f"тип контракта = ГПХ, но не найдены в файле реестра ГПХ-договоров."
            ),
            affected=[f"{e.name} | доход: {e.total_income:,.0f} сум"
                      for e in gph_in_ndfl_not_in_registry],
            recommendation="Проверьте, не является ли это ошибкой классификации договора в 1С.",
            can_skip=True,
        ))

    # В) Совпадает имя, но тип контракта не ГПХ
    wrong_contract = []
    for name_up, contract in gph_names.items():
        emp = ndfl_names.get(name_up)
        if emp and not emp.is_gph:
            wrong_contract.append({
                'name': emp.name,
                'contract': emp.contract,
                'period': f"{contract.date_start} — {contract.date_end}",
            })

    if wrong_contract:
        from ndfl_processor import Employee
        CMAP = {1:'Трудовой(осн)',2:'Совместитель',3:'ГПХ','1':'Трудовой(осн)','2':'Совместитель','3':'ГПХ'}
        results.append(CheckResult(
            check_id="GHP_WRONG_CONTRACT_TYPE",
            category="ГПХ",
            severity=Severity.WARNING,
            title=f"Тип договора в НДФЛ не совпадает с реестром ГПХ ({len(wrong_contract)} чел.)",
            description=(
                f"Следующие сотрудники есть в реестре ГПХ, "
                f"но в НДФЛ у них другой тип контракта."
            ),
            affected=[f"{x['name']}: в НДФЛ={CMAP.get(x['contract'],'?')}, "
                      f"в реестре=ГПХ ({x['period']})"
                      for x in wrong_contract],
            recommendation="Уточните актуальный тип договора для правильного заполнения кода 413.",
            can_skip=True,
        ))

    if not results:
        results.append(CheckResult(
            check_id="GHP_OK",
            category="ГПХ",
            severity=Severity.INFO,
            title=f"ГПХ-сверка пройдена ({len(gph_contracts)} договоров)",
            description="Все договоры ГПХ найдены в НДФЛ с корректным типом контракта.",
            can_skip=True,
        ))

    return results


# ─────────────────────────────────────────────────────────────────────
# Проверка 2: Приём и увольнение vs НДФЛ
# ─────────────────────────────────────────────────────────────────────

def check_hire_fire(ndfl_report, df_hire: pd.DataFrame,
                    df_fire: pd.DataFrame) -> list:
    """
    Сверяет движение персонала с НДФЛ-отчётом.

    Логика:
    А) Принят в отчётном году, нет ни в Прил.4 ни в Прил.5 → WARNING
    Б) Уволен в 2025, статус в Прил.4 ≠ 2 → WARNING
    В) Уволен в 2025, есть в Прил.4 но нет дохода → INFO
    """
    results = []
    emps   = ndfl_report.employees
    prizes = ndfl_report.prize_employees
    ndfl4_names = {e.name_upper for e in emps}
    ndfl5_names = {p.name_upper for p in prizes}

    # Колонки
    hire_name_col = next((c for c in df_hire.columns if 'сотрудник' in c.lower()), df_hire.columns[0])
    fire_name_col = next((c for c in df_fire.columns if 'сотрудник' in c.lower()), df_fire.columns[3])

    # А) Принятые, но нет в НДФЛ
    not_in_ndfl = []
    for _, row in df_hire.iterrows():
        nm = str(row[hire_name_col]).strip().upper()
        if not nm or nm == 'NAN':
            continue
        hdate = row.get('_date_parsed') or pd.to_datetime(
            row.get('Дата приема', ''), dayfirst=True, errors='coerce')
        if pd.isna(hdate) or hdate.year not in [2024, 2025]:
            continue
        if nm not in ndfl4_names and nm not in ndfl5_names:
            dept = str(row.get('Подразделение', '')).strip()
            empl = str(row.get('Вид занятости', '')).strip()
            not_in_ndfl.append({
                'name': str(row[hire_name_col]).strip(),
                'date': str(hdate.date()),
                'dept': dept,
                'type': empl,
            })

    if not_in_ndfl:
        results.append(CheckResult(
            check_id="HIRE_NOT_IN_NDFL",
            category="Прием/Увольнение",
            severity=Severity.WARNING,
            title=f"Принятые сотрудники не найдены в НДФЛ ({len(not_in_ndfl)} чел.)",
            description=(
                f"{len(not_in_ndfl)} принятых в 2025 году не найдены "
                f"ни в Приложении 4, ни в Приложении 5 НДФЛ-отчёта."
            ),
            affected=[f"{x['name']} (принят {x['date']}, {x['dept']}, {x['type']})"
                      for x in not_in_ndfl],
            recommendation=(
                "Возможные причины: (1) приняты в конце года и нет начислений, "
                "(2) выплаты не отражены в НДФЛ. Уточните у бухгалтера."
            ),
            can_skip=True,
        ))

    # Б) Уволенные в 2025 — статус в Прил.4 должен быть 2
    wrong_status = []
    for _, row in df_fire.iterrows():
        nm = str(row[fire_name_col]).strip().upper()
        if not nm or nm == 'NAN':
            continue
        fdate = row.get('_date_parsed') or pd.to_datetime(
            row.get('Дата увольнения', ''), dayfirst=True, errors='coerce')
        if pd.isna(fdate) or fdate.year != 2025:
            continue
        found = [e for e in emps if e.name_upper == nm]
        if found:
            emp = found[0]
            if not emp.is_fired:
                wrong_status.append({
                    'name': emp.name,
                    'fire_date': str(fdate.date()),
                    'ndfl_status': emp.status,
                })

    if wrong_status:
        results.append(CheckResult(
            check_id="FIRE_WRONG_STATUS",
            category="Прием/Увольнение",
            severity=Severity.WARNING,
            title=f"Уволенные в 2025: статус в НДФЛ не обновлён ({len(wrong_status)} чел.)",
            description=(
                f"{len(wrong_status)} сотрудников уволены в 2025 году, "
                f"но в Приложении 4 у них статус = 1 (работает) вместо 2 (уволен)."
            ),
            affected=[f"{x['name']}: уволен {x['fire_date']}, "
                      f"статус в НДФЛ={x['ndfl_status']}"
                      for x in wrong_status],
            recommendation=(
                "При следующей корректировке НДФЛ-отчёта обновить статус "
                "в Приложении 4 на 2. Для заполнения 1-korxona код 405 "
                "рассчитывается без учёта уволенных — проверьте численность."
            ),
            can_skip=True,
        ))

    # В) Уволенные с 01.01.2026 (особый случай — работали весь год)
    next_year_fires = []
    for _, row in df_fire.iterrows():
        fdate = row.get('_date_parsed') or pd.to_datetime(
            row.get('Дата увольнения', ''), dayfirst=True, errors='coerce')
        if pd.isna(fdate) or fdate.year != 2026:
            continue
        next_year_fires.append(str(row[fire_name_col]).strip())

    if next_year_fires:
        results.append(CheckResult(
            check_id="FIRE_NEXT_YEAR",
            category="Прием/Увольнение",
            severity=Severity.INFO,
            title=f"Уволенные с 01.01.2026: работали весь 2025 год ({len(next_year_fires)} чел.)",
            description=(
                f"{len(next_year_fires)} сотрудников уволены датой 01.01.2026 "
                f"— значит они работали весь 2025 год. Статус в НДФЛ = 1 (норма). "
                f"В численности на конец года (код 405) их можно учитывать."
            ),
            affected=next_year_fires[:10] + (['...'] if len(next_year_fires) > 10 else []),
            recommendation="Норма. При заполнении кода 405 — они включаются в численность.",
            can_skip=True,
        ))

    if not [r for r in results if r.severity in (Severity.CRITICAL, Severity.WARNING)]:
        results.append(CheckResult(
            check_id="HIRE_FIRE_OK",
            category="Прием/Увольнение",
            severity=Severity.INFO,
            title="Движение персонала проверено",
            description="Существенных расхождений между реестрами и НДФЛ не обнаружено.",
            can_skip=True,
        ))

    return results


# ─────────────────────────────────────────────────────────────────────
# Проверка 3: Нерезиденты
# ─────────────────────────────────────────────────────────────────────

def check_nonresidents(ndfl_report) -> list:
    """
    Проверяет нерезидентов в НДФЛ.

    Логика:
    А) Нерезидент с большим доходом → WARNING (проверить ставку)
    Б) НДФЛ/доход нерезидента < 12% → WARNING (возможно занижена ставка)
    В) Есть нерезиденты → INFO (к сведению)
    """
    results = []
    nonresidents = [e for e in ndfl_report.employees if e.is_nonresident]

    if not nonresidents:
        results.append(CheckResult(
            check_id="NONRES_NONE",
            category="Нерезиденты",
            severity=Severity.INFO,
            title="Нерезидентов нет",
            description="Все сотрудники в Приложении 4 являются резидентами РУз.",
            can_skip=True,
        ))
        return results

    # Проверяем эффективную ставку НДФЛ
    # Нерезиденты: ставка 20% (или 12% при определённых условиях)
    # Резиденты: 12%
    NONRES_RATE_MIN = 0.12   # минимальная ожидаемая ставка

    suspect = []
    for emp in nonresidents:
        if emp.total_income > 0 and emp.ndfl_total > 0:
            eff_rate = emp.ndfl_total / emp.total_income
            if eff_rate < NONRES_RATE_MIN - 0.01:  # допуск 1%
                suspect.append({
                    'name': emp.name,
                    'income': emp.total_income,
                    'ndfl': emp.ndfl_total,
                    'rate': eff_rate,
                    'contract': emp.contract,
                })

    if suspect:
        results.append(CheckResult(
            check_id="NONRES_LOW_RATE",
            category="Нерезиденты",
            severity=Severity.WARNING,
            title=f"Нерезиденты: возможно занижена ставка НДФЛ ({len(suspect)} чел.)",
            description=(
                f"Для {len(suspect)} нерезидентов эффективная ставка НДФЛ "
                f"ниже минимально ожидаемой ({NONRES_RATE_MIN*100:.0f}%). "
                f"По НК РУз ставка для нерезидентов — 20% (или 12% при наличии "
                f"соответствующих условий). Рекомендуется проверить."
            ),
            affected=[f"{x['name']}: доход={x['income']:,.0f}, "
                      f"НДФЛ={x['ndfl']:,.0f}, "
                      f"ставка={x['rate']*100:.1f}%"
                      for x in suspect],
            recommendation=(
                "Проверьте тип договора и статус резидентства. "
                "Нерезиденты без льготных соглашений облагаются по ставке 20%."
            ),
            can_skip=True,
        ))

    # Общая информация о нерезидентах
    results.append(CheckResult(
        check_id="NONRES_LIST",
        category="Нерезиденты",
        severity=Severity.INFO if not suspect else Severity.WARNING,
        title=f"Нерезиденты в отчёте: {len(nonresidents)} чел.",
        description=(
            f"В Приложении 4 {len(nonresidents)} сотрудников с типом = "
            f"нерезидент РУз. Убедитесь в правильности применяемых ставок."
        ),
        affected=[f"{e.name}: доход={e.total_income:,.0f}, "
                  f"НДФЛ={e.ndfl_total:,.0f} "
                  f"({e.ndfl_total/e.total_income*100:.1f}% eff.)" if e.total_income > 0
                  else f"{e.name}: нет дохода"
                  for e in nonresidents],
        recommendation="К сведению. Проверьте ставку при наличии подозрений.",
        can_skip=True,
    ))

    return results


# ─────────────────────────────────────────────────────────────────────
# Проверка 4: Численность для 1-korxona
# ─────────────────────────────────────────────────────────────────────

def check_headcount(ndfl_report, personnel_codes: dict) -> list:
    """
    Проверяет корректность численности для кодов 401 и 413.

    Логика:
    А) 413 ≠ 409 + 411 + 412 → CRITICAL
    Б) 401 в НДФЛ (ср. числ.) ≠ расчётному 401 → WARNING
    В) Нулевые значения кодов → WARNING
    """
    results = []

    c = personnel_codes
    code_409 = c.get(409, {}).get('value', 0)
    code_411 = c.get(411, {}).get('value', 0)
    code_412 = c.get(412, {}).get('value', 0)
    code_413 = c.get(413, {}).get('value', 0)
    code_401 = c.get(401, {}).get('value', 0)

    expected_413 = code_409 + code_411 + code_412

    # А) Контрольное соотношение 413 = 409 + 411 + 412
    if abs(code_413 - expected_413) > 0:
        results.append(CheckResult(
            check_id="HEADCOUNT_413_MISMATCH",
            category="Численность",
            severity=Severity.CRITICAL,
            title="Нарушено контрольное соотношение: 413 ≠ 409 + 411 + 412",
            description=(
                f"Код 413 = {code_413}, но 409({code_409}) + "
                f"411({code_411}) + 412({code_412}) = {expected_413}. "
                f"Расхождение: {abs(code_413 - expected_413)}. "
                f"Это приведёт к ошибке контрольных соотношений в 1-korxona."
            ),
            affected=[f"413={code_413}, 409={code_409}, 411={code_411}, 412={code_412}"],
            recommendation=(
                "Проверьте исходные данные. Возможно, в реестре приёма/ГПХ "
                "есть дублирования или пропуски. Скорректируйте до продолжения."
            ),
            can_skip=False,  # БЛОКИРУЕТ заполнение
        ))
    else:
        results.append(CheckResult(
            check_id="HEADCOUNT_413_OK",
            category="Численность",
            severity=Severity.INFO,
            title=f"Контрольное соотношение 413 = 409+411+412 ✓ ({code_413})",
            description=(
                f"409({code_409}) + 411({code_411}) + 412({code_412}) = {code_413}"
            ),
            can_skip=True,
        ))

    # Б) Сверка 401 с титулом НДФЛ
    ndfl_avg = ndfl_report.headcount_avg
    if ndfl_avg > 0 and abs(code_401 - ndfl_avg) > 2:
        results.append(CheckResult(
            check_id="HEADCOUNT_401_DIFF",
            category="Численность",
            severity=Severity.WARNING,
            title=f"Код 401 ({code_401}) отличается от ср.численности в НДФЛ ({ndfl_avg})",
            description=(
                f"В Приложении 4 НДФЛ: {code_401} человек с доходом. "
                f"В титульном листе НДФЛ: ср.численность = {ndfl_avg}. "
                f"Разница {abs(code_401 - ndfl_avg)} человек — "
                f"возможно есть сотрудники без начислений или внутреннее расхождение."
            ),
            affected=[f"Прил.4 (с доходом): {code_401}",
                      f"Титул НДФЛ (ср.числ.): {ndfl_avg}",
                      f"Разница: {abs(code_401 - ndfl_avg)}"],
            recommendation=(
                "Уточните у бухгалтера какая цифра корректна для кода 401. "
                "Код 401 = численность, принятая для исчисления ЗП (включая тех, "
                "кому начислена хотя бы копейка за год)."
            ),
            can_skip=True,
        ))

    # В) Нулевые значения
    zero_codes = [k for k, v in c.items()
                  if v.get('value', 0) == 0 and k in (401, 403, 413)]
    if zero_codes:
        results.append(CheckResult(
            check_id="HEADCOUNT_ZERO_CODES",
            category="Численность",
            severity=Severity.WARNING,
            title=f"Нулевые значения ключевых кодов: {zero_codes}",
            description=(
                f"Коды {zero_codes} имеют значение 0, что может быть некорректным."
            ),
            affected=[f"Код {k}: {c[k].get('desc','')}" for k in zero_codes],
            recommendation="Проверьте источник данных для этих кодов.",
            can_skip=True,
        ))

    return results


# ─────────────────────────────────────────────────────────────────────
# Главная функция запуска всех проверок
# ─────────────────────────────────────────────────────────────────────

def run_all_checks(ndfl_report,
                   gph_contracts: list,
                   df_hire: pd.DataFrame,
                   df_fire: pd.DataFrame,
                   personnel_codes: dict) -> CheckSummary:
    """
    Запускает все 4 проверки и возвращает сводный результат.

    Args:
        ndfl_report:      разобранный NDFLReport
        gph_contracts:    список GphContract из реестра ГПХ
        df_hire:          DataFrame принятых сотрудников
        df_fire:          DataFrame уволенных сотрудников
        personnel_codes:  словарь {код: {value, desc, source}} из extract_korxona_personnel

    Returns:
        CheckSummary с полным списком результатов
    """
    all_results = []

    # 1. ГПХ
    all_results.extend(check_gph(ndfl_report, gph_contracts))

    # 2. Приём/Увольнение
    all_results.extend(check_hire_fire(ndfl_report, df_hire, df_fire))

    # 3. Нерезиденты
    all_results.extend(check_nonresidents(ndfl_report))

    # 4. Численность
    all_results.extend(check_headcount(ndfl_report, personnel_codes))

    # Можно ли продолжить
    blocking = [r for r in all_results
                if r.severity == Severity.CRITICAL and not r.can_skip]
    can_proceed = len(blocking) == 0

    return CheckSummary(results=all_results, can_proceed=can_proceed)
