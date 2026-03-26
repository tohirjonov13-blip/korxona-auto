"""
skp.py — Модуль работы со Статистическим классификатором продукции (СКП) РУз
Файл данных: skp_classifier.json (082-3-006)

Функции:
  - load()            → загрузка и индексирование классификатора
  - search(query)     → полнотекстовый поиск по названию
  - validate(code)    → проверка корректности кода СКП
  - get_tiftn(code)   → получение кода ТНВЭД по коду СКП
  - get_path(code)    → иерархический путь (хлебные крошки)
  - get_children(code)→ дочерние позиции
  - get_by_oked(oked) → все позиции раздела по коду ОКЭД (A-U, 01-99)
"""

import json
import re
from pathlib import Path
from functools import lru_cache
from typing import Optional

# ─── Путь к файлу классификатора ─────────────────────────────────────────────
_DEFAULT_PATH = Path(__file__).parent / "skp_classifier.json"


class SKPClassifier:
    """Классификатор продукции (товаров, работ, услуг) РУз — СКП (082-3-006)"""

    def __init__(self, json_path: str = None):
        path = json_path or _DEFAULT_PATH
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)

        # Индексы
        self._by_code: dict[str, dict] = {}       # CODE → запись
        self._all: list[dict] = []                 # полный список

        for item in raw:
            code = item["CODE"].strip()
            entry = {
                "code": code,
                "name": item["NAME"].strip(),
                "tiftn": item.get("TIFTN", "").strip(),
                "level": self._level(code),
            }
            self._by_code[code] = entry
            self._all.append(entry)

        # Нормализованные названия для поиска
        self._search_index = [
            (item["code"], item["name"].lower(), item)
            for item in self._all
        ]

    # ─── Утилиты ─────────────────────────────────────────────────────────────

    @staticmethod
    def _level(code: str) -> str:
        """Определить уровень иерархии кода"""
        if len(code) == 1 and code.isalpha():
            return "section"     # A, B, C...
        parts = code.split(".")
        return {1: "division", 2: "group", 3: "class", 4: "subclass"}.get(len(parts), "unknown")

    @staticmethod
    def _normalize(text: str) -> str:
        return re.sub(r"\s+", " ", text.lower().strip())

    # ─── Поиск ───────────────────────────────────────────────────────────────

    def search(self, query: str, max_results: int = 20,
               section: str = None, level: str = None) -> list[dict]:
        """Полнотекстовый поиск по наименованию продукции.

        Args:
            query:       строка поиска (русский текст)
            max_results: максимальное количество результатов
            section:     фильтр по разделу ('A', 'C', 'F' и т.д.)
            level:       фильтр по уровню ('section','division','group','class','subclass')

        Returns:
            список словарей {code, name, tiftn, level, score}
        """
        q = self._normalize(query)
        words = q.split()
        results = []

        for code, name_low, item in self._search_index:
            # Фильтр по разделу
            if section:
                sec = code[0].upper() if code[0].isalpha() else None
                if sec != section.upper():
                    # Также проверяем числовые коды — находим раздел по первым цифрам
                    prefix = code.split(".")[0]
                    parent_sec = self._find_section_for_code(code)
                    if parent_sec != section.upper():
                        continue

            # Фильтр по уровню (возвращаем только конечные позиции если не указано)
            if level:
                if item["level"] != level:
                    continue

            # Скоринг: полное совпадение > все слова > частичное
            score = 0
            if q == name_low:
                score = 100
            elif q in name_low:
                score = 80
            elif all(w in name_low for w in words):
                score = 60 + sum(1 for w in words if name_low.startswith(w))
            elif any(w in name_low for w in words if len(w) > 3):
                score = 20 + sum(10 for w in words if w in name_low)
            else:
                continue

            results.append({**item, "score": score})

        results.sort(key=lambda x: (-x["score"], x["code"]))
        return results[:max_results]

    def _find_section_for_code(self, code: str) -> Optional[str]:
        """Найти раздел (A-U) для числового кода"""
        # Разделы сопоставляются с диапазонами ОКЭД
        division = code.split(".")[0]
        try:
            num = int(division)
        except ValueError:
            return code[0] if code[0].isalpha() else None

        # Маппинг ОКЭД → Раздел СКП
        DIVISION_TO_SECTION = {
            range(1, 4): "A",     # 01-03: Сельское хозяйство
            range(5, 10): "B",    # 05-09: Добыча
            range(10, 34): "C",   # 10-33: Обрабатывающая
            range(35, 36): "D",   # 35: Электроэнергия
            range(36, 40): "E",   # 36-39: Водоснабжение
            range(41, 44): "F",   # 41-43: Строительство
            range(45, 48): "G",   # 45-47: Торговля
            range(49, 54): "H",   # 49-53: Транспорт
            range(55, 57): "I",   # 55-56: Гостиницы
            range(58, 64): "J",   # 58-63: ИКТ
            range(64, 67): "K",   # 64-66: Финансы
            range(68, 69): "L",   # 68: Недвижимость
            range(69, 76): "M",   # 69-75: Профессиональные услуги
            range(77, 83): "N",   # 77-82: Административные
            range(84, 85): "O",   # 84: Госуправление
            range(85, 86): "P",   # 85: Образование
            range(86, 89): "Q",   # 86-88: Здравоохранение
            range(90, 94): "R",   # 90-93: Искусство
            range(94, 97): "S",   # 94-96: Прочие услуги
            range(97, 99): "T",   # 97-98: Домашние хозяйства
            range(99, 100): "U",  # 99: Экстерриториальные
        }
        for r, sec in DIVISION_TO_SECTION.items():
            if num in r:
                return sec
        return None

    # ─── Валидация ───────────────────────────────────────────────────────────

    def validate(self, code: str) -> dict:
        """Проверить корректность кода СКП.

        Returns:
            {"valid": bool, "entry": dict|None, "error": str|None}
        """
        code = code.strip()
        entry = self._by_code.get(code)
        if entry:
            return {"valid": True, "entry": entry, "error": None}

        # Попробуем найти похожий
        suggestions = self.search(code, max_results=3)
        err_msg = f"Код '{code}' не найден в СКП."
        if suggestions:
            err_msg += f" Возможно: {', '.join(s['code'] for s in suggestions)}"
        return {"valid": False, "entry": None, "error": err_msg}

    # ─── ТНВЭД ───────────────────────────────────────────────────────────────

    def get_tiftn(self, skp_code: str) -> Optional[str]:
        """Получить код ТНВЭД (ТИФ ТН) по коду СКП.

        Если на данном уровне ТНВЭД нет — ищем у родителей.
        """
        code = skp_code.strip()
        entry = self._by_code.get(code)
        if not entry:
            return None
        if entry["tiftn"]:
            return entry["tiftn"]

        # Подъём к родителям
        parts = code.split(".")
        while len(parts) > 1:
            parts = parts[:-1]
            parent_code = ".".join(parts)
            parent = self._by_code.get(parent_code)
            if parent and parent["tiftn"]:
                return f"(от родителя) {parent['tiftn']}"
        return None

    # ─── Иерархия ────────────────────────────────────────────────────────────

    def get_path(self, code: str) -> list[dict]:
        """Получить иерархический путь (хлебные крошки) для кода СКП.

        Пример для '01.11.11.1':
          A → 01 → 01.1 → 01.11 → 01.11.1 → 01.11.11 → 01.11.11.1
        """
        code = code.strip()
        path = []

        # Числовые коды: строим путь через разбиение по точкам
        if "." in code or code.isdigit():
            parts = code.split(".")
            # Добавляем раздел (секцию)
            section = self._find_section_for_code(code)
            if section and section in self._by_code:
                path.append(self._by_code[section])

            # Добавляем каждый уровень
            for i in range(1, len(parts) + 1):
                partial = ".".join(parts[:i])
                if partial in self._by_code:
                    path.append(self._by_code[partial])
        else:
            # Секция (A, B...)
            if code in self._by_code:
                path.append(self._by_code[code])

        return path

    def get_children(self, code: str, direct_only: bool = True) -> list[dict]:
        """Получить дочерние позиции для кода.

        Args:
            direct_only: только прямые дети (следующий уровень)
        """
        code = code.strip()
        results = []
        code_dots = code.count(".")

        for item in self._all:
            c = item["code"]
            if c == code:
                continue
            # Дочерний код начинается с родительского
            if c.startswith(code + ".") or (code.isalpha() and c[:len(code)] == code and not c[len(code):len(code)+1].isalpha()):
                if direct_only:
                    child_dots = c.count(".")
                    if child_dots == code_dots + 1:
                        results.append(item)
                else:
                    results.append(item)
        return results

    def get_by_oked(self, oked: str) -> list[dict]:
        """Все позиции СКП соответствующие коду ОКЭД (раздел A-U или числовой 01-99)"""
        oked = oked.strip().upper()
        if oked.isalpha():
            # Раздел: возвращаем всё дерево раздела
            return [item for item in self._all
                    if item["code"].startswith(oked) and item["code"] != oked]
        else:
            # Числовой ОКЭД — возвращаем дивизион
            return [item for item in self._all
                    if item["code"].startswith(oked)]

    def get_entry(self, code: str) -> Optional[dict]:
        """Получить запись по точному коду"""
        return self._by_code.get(code.strip())

    @property
    def stats(self) -> dict:
        """Статистика по классификатору"""
        from collections import Counter
        levels = Counter(item["level"] for item in self._all)
        sections = [item for item in self._all if item["level"] == "section"]
        with_tiftn = sum(1 for item in self._all if item["tiftn"])
        return {
            "total": len(self._all),
            "sections": len(sections),
            "with_tiftn": with_tiftn,
            "by_level": dict(levels),
        }


# ─── Синглтон ─────────────────────────────────────────────────────────────────

_instance: Optional[SKPClassifier] = None


def get_skp(json_path: str = None) -> SKPClassifier:
    """Получить экземпляр классификатора (загружается один раз)"""
    global _instance
    if _instance is None:
        _instance = SKPClassifier(json_path)
    return _instance


# ─── CLI быстрой проверки ─────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    skp = get_skp()
    print(f"СКП загружен: {skp.stats}")

    if len(sys.argv) > 1:
        query = " ".join(sys.argv[1:])
        print(f"\nПоиск: '{query}'")
        results = skp.search(query, max_results=10)
        for r in results:
            tiftn = f"  ТНВЭД: {r['tiftn']}" if r["tiftn"] else ""
            print(f"  [{r['code']}] {r['name']}{tiftn}")
    else:
        # Демо
        print("\n--- Поиск 'хлопок' ---")
        for r in skp.search("хлопок", max_results=5):
            print(f"  [{r['code']}] {r['name']}  ТНВЭД: {r.get('tiftn','—')}")

        print("\n--- Путь для кода 01.11.11.1 ---")
        for p in skp.get_path("01.11.11.1"):
            indent = "  " * p["level"].count(".")
            print(f"  {p['code']} → {p['name']}")

        print("\n--- ТНВЭД для 01.11.11.1 ---")
        print(" ", skp.get_tiftn("01.11.11.1"))

        print("\n--- Валидация кода 99.99.99.9 ---")
        print(" ", skp.validate("99.99.99.9"))
